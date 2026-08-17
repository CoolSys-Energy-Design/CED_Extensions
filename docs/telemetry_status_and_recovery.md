# CED Telemetry: Current State and Recovery Plan

Status reviewed: 2026-08-05

## Executive summary

Telemetry is intentionally paused for distributed builds.

The extension still checks the ACC/Desktop Connector workspace route and updates the local route-status file. It does not create or configure pyRevit telemetry files, and it does not transfer telemetry to Desktop Connector.

The retained telemetry implementation is still in the repository behind two switches so it can be tested later on a controlled machine.

## Relevant files

- [`startup.py`](../AE%20pyTools.extension/startup.py) — pyRevit startup hook, emergency switches, route check, and shutdown transfer hook.
- [`telemetry_route.py`](../AE%20pyTools.extension/telemetry_route.py) — workspace discovery, route scoring, local route state, and transfer helpers.
- [`adc_startup_diagnostics.py`](../CEDLib.lib/UnitTests/adc_startup_diagnostics.py) — developer test harness and read-only route diagnostics.
- `ADC Diagnostics.pushbutton/script.py` — separate user-facing diagnostics command tested by the harness.

## Current release switches

At the top of `startup.py`:

```python
ENABLE_PYREVIT_TELEMETRY = False
ENABLE_DESKTOP_CONNECTOR_TELEMETRY_TRANSFER = False
```

Both values must remain `False` in distributed builds.

### Startup behavior while disabled

The startup hook runs in this order:

```text
configure telemetry state
check ACC/Desktop Connector route
register shutdown hook
```

With telemetry disabled, the first step:

1. Reads the configured pyRevit telemetry active state.
2. Reads the live pyRevit telemetry active state.
3. If either state is enabled, calls pyRevit's native `set_telemetry_state(False)`.
4. Saves the user configuration only when the persisted state needed changing.

The disabled path does not:

- Set `telemetry_file_dir`.
- Set `telemetry_file_path`.
- Call `setup_telemetry()`.
- Create a telemetry directory or telemetry JSON.
- Modify telemetry server URLs.
- Transfer or delete telemetry files.

### Route behavior while disabled

`_check_acc_sync()` remains active. It calls `telemetry_route.resolve_usage_route(persist=True)` so workspace discovery and route status continue to work.

This may create or update:

```text
%APPDATA%\pyRevit\Extensions\CED_pyTelemetry\.ced_usage_route_status.json
```

That file is local route metadata, not a pyRevit telemetry session file. It is intentionally retained because route checking is still required.

### Shutdown behavior while disabled

The `ApplicationClosing` transfer hook is not registered. `_on_app_closing()` also has an early return so an old handler cannot transfer files if it remains attached in the current Revit process.

No Desktop Connector destination is resolved during shutdown, and no source telemetry files are copied, deleted, or migrated.

## How pyRevit telemetry actually works

There are three separate values involved:

| Value | Meaning |
| --- | --- |
| `telemetry_file_dir` in `pyRevit_config.ini` | Persisted directory configuration. |
| `TELEMETRYDIR` runtime value | Directory currently loaded into the Revit process. |
| `TELEMETRYFILE` runtime value | The current session's specific JSON path. |

The pyRevit Settings window displays the live `TELEMETRYFILE` value. It does not search the telemetry directory for JSON files.

During native pyRevit startup, `setup_telemetry()`:

1. Reads the persisted telemetry settings.
2. Pushes the configured directory into the runtime environment with `persist=False`.
3. Creates and assigns a current JSON path only when telemetry is active and the directory is valid.
4. Clears the current file path when the directory is invalid or file initialization fails.

Therefore, a JSON file existing on disk does not prove that the current Revit session has a nonblank `TELEMETRYFILE` value.

When the emergency switch is active, `active = false` is the intended result. On the next native startup, pyRevit should not assign a current telemetry JSON path, so Settings showing a blank current-file field is expected.

## What caused the previous problems

The failures came from mixing three different responsibilities and startup timings.

### Startup ordering

pyRevit's native loader initializes telemetry before extension `startup.py` runs. CED cannot retroactively control the native setup that already occurred during that process.

This explains why a first launch could create a file before CED disabled telemetry, or why changing a folder from CED did not automatically create a current session path.

### Configuration serialization

The earlier implementation used generic config access and save calls for telemetry settings. Repeated read/write cycles interacted badly with pyRevit's config serialization and produced progressively escaped paths such as doubled backslashes. Server URL values were also observed in malformed forms such as `""/`.

Relevant upstream history:

- [pyRevit issue #3334](https://github.com/pyrevitlabs/pyRevit/issues/3334) — configuration escape-doubling behavior.
- [pyRevit issue #3534](https://github.com/pyrevitlabs/pyRevit/issues/3534) — short malformed values not being repaired by the earlier recovery logic.

### Persisted path versus live path

Writing a correct directory to config does not initialize the current session's JSON path. Calling `setup_telemetry()` from CED to compensate would risk creating a second session file or re-running native initialization in the wrong context.

The retained local-test path uses pyRevit 6.5's native setter and deliberately avoids calling `setup_telemetry()`. Native pyRevit owns session-file creation on the next startup.

### Server URL handling

CED startup must not rewrite `telemetry_server_url` or `apptelemetry_server_url`. Existing malformed URL values are a separate pyRevit configuration problem and are intentionally left untouched while telemetry is disabled.

## What is currently persisted

With the emergency switch taking effect, CED may persist this one setting:

```ini
[telemetry]
active = false
```

CED does not clear or replace the existing `telemetry_file_dir`. An old custom directory may therefore remain visible in pyRevit Settings even though telemetry is inactive and no current JSON path is assigned.

## Recovery plan

Do not begin by re-enabling telemetry for all users. Use one controlled test machine and a backed-up pyRevit config.

### Phase 1: establish native pyRevit behavior

Keep both CED switches `False`.

1. Record the installed pyRevit version and active engine.
2. Back up the user pyRevit config before changing anything.
3. Confirm that `active = false` persists after a Revit restart.
4. Confirm that no new telemetry JSON is created on the following startup.
5. Confirm that Settings shows a blank current telemetry file because telemetry is inactive.
6. Confirm that `.ced_usage_route_status.json` still updates independently.

Do not use the presence of a JSON file alone as proof of current-session initialization.

### Phase 2: test telemetry setup locally

On the controlled test machine only:

1. Set `ENABLE_PYREVIT_TELEMETRY = True`.
2. Leave `ENABLE_DESKTOP_CONNECTOR_TELEMETRY_TRANSFER = False`.
3. Start from a backed-up, known-valid pyRevit config.
4. Test one directory change at a time.
5. Let native pyRevit create the session file on the next Revit startup; do not call `setup_telemetry()` from CED.
6. Verify all of the following after restart:
   - Persisted `telemetry_file_dir` is a normal Windows path.
   - Live `TELEMETRYDIR` matches the persisted path.
   - Live `TELEMETRYFILE` is a single current session path.
   - The current file exists and is writable.
   - Only one new session file is created.
   - Server URL values are unchanged.
   - Settings displays the same current path returned by pyRevit's getter.

If any of these fail, stop and capture the config file, pyRevit version, startup log, live getters, and directory contents before trying another fix.

### Phase 3: test route and transfer separately

Keep telemetry setup enabled only on the controlled machine, but leave transfer disabled while validating route behavior.

1. Test route discovery and manual route approval.
2. Verify `.ced_usage_route_status.json` contents and updates.
3. Use a test telemetry file and test Desktop Connector destination.
4. Set `ENABLE_DESKTOP_CONNECTOR_TELEMETRY_TRANSFER = True` only after local file setup is reliable.
5. Validate copy, collision, deletion, failure, and multiple-Revit-process behavior.
6. Return the transfer switch to `False` immediately after testing.

Do not test transfer against production telemetry while the source/destination contract is still being changed.

### Phase 4: update automated coverage

Before enabling anything for users, update `adc_startup_diagnostics.py` to match the current startup API.

The test shim needs:

- A `pyrevit.userconfig` module.
- A fake `telemetry_status` property.
- `save_changes()` tracking.
- Native telemetry state persistence behavior.
- A Python 3.12-compatible `imp` shim or an import-safe test loader.

Required tests:

- Active telemetry is changed to `False` only when needed.
- No telemetry folder helper runs while disabled.
- No telemetry setup or file-path setter runs while disabled.
- An already-disabled configuration produces no config write.
- Route checking still runs while telemetry is disabled.
- Shutdown transfer returns immediately while disabled.
- Legacy enabled-mode tests run only with the switches explicitly set to `True`.

## Rules for the eventual fix

When telemetry work resumes:

- Do not use raw `script.get_config()` and `script.save_config()` for telemetry settings.
- Do not rewrite server URLs from CED startup.
- Use pyRevit's native telemetry API and version-aware behavior.
- Treat `telemetry_file_dir`, `TELEMETRYDIR`, and `TELEMETRYFILE` as separate values.
- Use `persist=False` for native startup-time environment initialization.
- Use `persist=True` only for an intentional directory change, followed by one controlled config save when required.
- Do not call `setup_telemetry()` from the extension startup hook.
- Never transfer or delete a file unless its current-session identity and destination are both verified.
- Keep route status updates independent from telemetry session-file creation.

## Future installer configuration direction

Do not return telemetry to distributed builds until the relevant pyRevit telemetry issues are resolved upstream.

When telemetry work resumes, prefer having the installer invoke supported pyRevit CLI configuration commands once during installation rather than directly editing `pyRevit_config.ini` or reconfiguring telemetry from the CED startup hook. The intended model is:

1. The installer uses the pyRevit CLI to apply the telemetry settings once.
2. CED startup does not set the telemetry directory, rewrite telemetry configuration, or call `setup_telemetry()`.
3. Native pyRevit startup remains responsible for loading the persisted settings and creating the current session file.

This should avoid CED competing with pyRevit's startup initialization and should prevent the duplicate-session-file and stale-active-file behavior previously observed. Before enabling it for users, verify the supported CLI command and behavior against the target pyRevit versions, confirm the CLI is available in the installer's per-user context, and test first on a controlled machine with a backed-up configuration.

## Local re-enable switches

For a controlled test only:

```python
ENABLE_PYREVIT_TELEMETRY = True
ENABLE_DESKTOP_CONNECTOR_TELEMETRY_TRANSFER = False
```

This enables the retained telemetry setup but keeps shutdown transfer off. Only after setup is proven should the second switch be changed to `True` locally.

Return both switches to `False` before committing or distributing the extension.
