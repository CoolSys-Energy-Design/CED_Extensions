# Alerts Browser (User Guide)

## What This Tool Does
Alerts Browser is for troubleshooting circuits with alerts.  
You can pick a circuit, read active/hidden alerts, jump to related model elements, and refresh that one circuit after fixing issues.

## Quick Start
1. Open **Alerts Browser** from Circuit Manager.
2. Pick a circuit from the left list.
3. Review alerts on the right (**Active** and **Hidden** tabs).
4. Use **Panel / Circuit / Device** to find elements in model.
5. Fix model issue, then click **Refresh**.

## Main Layout
- **Left list**: circuits that currently have alerts.
- **Right tabs**:
  - **Active**: currently active alerts.
  - **Hidden**: hidden/suppressed alerts.
- **Grid Display**: switches to an **All Alerts** grid with one row per alert and no circuit picker, or back to the by-circuit display.
- The header summary always shows total **Alerts** and **Hidden** counts; `Show Hidden` only controls which rows are displayed.
- **All Alerts columns**: Circuit (`Panel / Circuit`), Load Name, Severity, Group, Alert ID, and Message. Click headers to sort.
- Hidden rows are shown with subdued/italic text when **Show Hidden** is enabled.
- **Bottom commands**: model selection buttons + refresh.

## Bottom Buttons
- **Panel**: selects upstream panel in model.
- **Circuit**: selects circuit element in model.
- **Device**: selects downstream connected devices.
- **X**: clears model selection.
- **Refresh**: recalculates selected circuit and reloads its alerts.

## Why Refresh Might Be Disabled
Refresh is disabled when:
- no circuit is selected
- an operation is already running
- recalculation is blocked by ownership/worksharing

Hover the disabled button to see the reason.

## Typical Workflow
1. Select a circuit with alerts.
2. Read alert message and group.
3. Jump to panel/circuit/device in model.
4. Make correction.
5. Click **Refresh**.
6. Confirm alert count drops or clears.

## Tips
- Keep this window open while editing; it is modeless.
- If nothing updates after edits, click **Refresh** again.
- If blocked by ownership, retry after checkout/ownership is available.
