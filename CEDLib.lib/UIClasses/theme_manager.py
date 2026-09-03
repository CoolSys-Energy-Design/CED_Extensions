# -*- coding: utf-8 -*-
"""Shared CED theme catalog, unlock codes, and pyRevit persistence."""

import os


THEME_CONFIG_SECTION = "AE-pyTools-Theme"
THEME_CONFIG_THEME_KEY = "theme_mode"
THEME_CONFIG_ACCENT_KEY = "accent_mode"
DEFAULT_ACCENT_MODE = "neutral"
THEME_UNLOCKS_KEY = "unlocked_theme_ids"
THEME_UNLOCK_SCHEMA_KEY = "theme_unlock_schema"
THEME_UNLOCK_SCHEMA_VERSION = 1


class ThemeDescriptor(object):
    """Metadata for one theme and its access requirements."""

    def __init__(
        self,
        mode,
        label,
        description,
        relative_paths,
        unlock_id=None,
        grandfather_existing=False,
    ):
        self.Mode = mode
        self.Label = label
        self.Description = description
        self.RelativePaths = tuple(relative_paths or ())
        self.UnlockId = unlock_id
        self.GrandfatherExisting = bool(grandfather_existing)

    @property
    def mode(self):
        return self.Mode

    @property
    def label(self):
        return self.Label

    @property
    def description(self):
        return self.Description

    @property
    def relative_paths(self):
        return self.RelativePaths

    def __getitem__(self, key):
        values = {
            "mode": self.Mode,
            "label": self.Label,
            "description": self.Description,
            "relative_paths": self.RelativePaths,
            "unlock_id": self.UnlockId,
        }
        return values[key]


class ThemeUnlockCode(object):
    """A key sequence that unlocks one hidden theme."""

    def __init__(self, unlock_id, theme_mode, sequence, success_message):
        self.UnlockId = unlock_id
        self.ThemeMode = theme_mode
        self.Sequence = tuple(sequence or ())
        self.SuccessMessage = success_message


class UnlockKeyDebouncer(object):
    """Emit one normalized key per physical press until that key is released."""

    def __init__(self):
        self._pressed = set()

    def reset(self):
        self._pressed.clear()

    def release(self, value):
        key = normalize_unlock_key(value)
        if key:
            self._pressed.discard(key)

    def press(self, value, is_repeat=False):
        key = normalize_unlock_key(value)
        if not key or bool(is_repeat) or key in self._pressed:
            return None
        self._pressed.add(key)
        return key


ALL_THEME_DESCRIPTORS = (
    ThemeDescriptor(
        "light",
        "Light",
        "Bright workspace",
        (os.path.join("Themes", "CEDTheme.Light.xaml"),),
    ),
    ThemeDescriptor(
        "dark",
        "Dark",
        "Deep blue-gray",
        (
            os.path.join("Themes", "CEDTheme.Dark.xaml"),
            os.path.join("Themes", "CEDTheme.Light.xaml"),
        ),
    ),
    ThemeDescriptor(
        "dark_alt",
        "Dark Alt",
        "Neutral dark",
        (
            os.path.join("Themes", "CEDTheme.DarkAlt.xaml"),
            os.path.join("Themes", "CEDTheme.Dark.xaml"),
            os.path.join("Themes", "CEDTheme.Light.xaml"),
        ),
    ),
    ThemeDescriptor(
        "cursed",
        "Cursed",
        "Neon goblin mode",
        (os.path.join("Themes", "CEDTheme.Cursed.xaml"),),
        unlock_id="konami",
        grandfather_existing=True,
    ),
    ThemeDescriptor(
        "chatgpt",
        "chatGPT",
        "Charcoal conversation mode",
        (os.path.join("Themes", "CEDTheme.ChatGPT.xaml"),),
        unlock_id="chatgpt_aislop",
    ),
)


THEME_UNLOCK_CODES = (
    ThemeUnlockCode(
        "konami",
        "cursed",
        ("UP", "UP", "DOWN", "DOWN", "LEFT", "RIGHT", "LEFT", "RIGHT", "B", "A", "ENTER"),
        "Cursed theme unlocked.",
    ),
    ThemeUnlockCode(
        "chatgpt_aislop",
        "chatgpt",
        ("A", "I", "S", "L", "O", "P"),
        "chatGPT theme unlocked.",
    ),
)


def all_theme_descriptors():
    """Return every registered theme, including locked themes."""
    return ALL_THEME_DESCRIPTORS


def _descriptor_for_mode(mode):
    normalized = str(mode or "").strip().lower()
    for descriptor in ALL_THEME_DESCRIPTORS:
        if descriptor.Mode == normalized:
            return descriptor
    return None


def _get_config(config=None):
    if config is not None:
        return config
    try:
        from pyrevit import script

        return script.get_config(THEME_CONFIG_SECTION)
    except Exception:
        return None


def _save_config(save_callback=None):
    try:
        if save_callback is not None:
            save_callback()
        else:
            from pyrevit import script

            script.save_config()
        return True
    except Exception:
        return False


def _parse_unlock_ids(value):
    if value is None:
        return set()
    if isinstance(value, (list, tuple, set)):
        values = value
    else:
        text = str(value or "").strip().strip("[]()")
        values = text.replace(";", ",").split(",")
    return set(
        str(item or "").strip().strip("'\"").lower()
        for item in values
        if str(item or "").strip().strip("'\"")
    )


def _write_unlock_ids(config, unlock_ids):
    config.set_option(THEME_UNLOCKS_KEY, ",".join(sorted(set(unlock_ids or ()))))


def unlocked_theme_ids(config=None, migrate=True, save_callback=None):
    """Return persisted unlock IDs and optionally perform one-time migration."""
    cfg = _get_config(config)
    if cfg is None:
        return set()
    unlock_ids = _parse_unlock_ids(cfg.get_option(THEME_UNLOCKS_KEY, ""))
    if not migrate:
        return unlock_ids

    try:
        schema_version = int(cfg.get_option(THEME_UNLOCK_SCHEMA_KEY, 0) or 0)
    except Exception:
        schema_version = 0
    if schema_version >= THEME_UNLOCK_SCHEMA_VERSION:
        return unlock_ids

    saved_mode = str(cfg.get_option(THEME_CONFIG_THEME_KEY, "light") or "light").strip().lower()
    descriptor = _descriptor_for_mode(saved_mode)
    if descriptor and descriptor.GrandfatherExisting and descriptor.UnlockId:
        unlock_ids.add(descriptor.UnlockId)
    _write_unlock_ids(cfg, unlock_ids)
    cfg.set_option(THEME_UNLOCK_SCHEMA_KEY, THEME_UNLOCK_SCHEMA_VERSION)
    _save_config(save_callback=save_callback)
    return unlock_ids


def theme_descriptors(config=None, save_callback=None):
    """Return only themes selectable by this pyRevit user."""
    unlock_ids = unlocked_theme_ids(config=config, migrate=True, save_callback=save_callback)
    return tuple(
        descriptor
        for descriptor in ALL_THEME_DESCRIPTORS
        if not descriptor.UnlockId or descriptor.UnlockId in unlock_ids
    )


def is_theme_unlocked(mode, config=None):
    descriptor = _descriptor_for_mode(mode)
    if descriptor is None:
        return False
    if not descriptor.UnlockId:
        return True
    return descriptor.UnlockId in unlocked_theme_ids(config=config)


def available_theme_mode(value, fallback="light", config=None):
    """Resolve a requested theme without permitting locked config injection."""
    descriptor = _descriptor_for_mode(value)
    if descriptor is not None and is_theme_unlocked(descriptor.Mode, config=config):
        return descriptor.Mode
    fallback_descriptor = _descriptor_for_mode(fallback)
    if fallback_descriptor is not None and is_theme_unlocked(fallback_descriptor.Mode, config=config):
        return fallback_descriptor.Mode
    return "light"


def normalize_unlock_key(value):
    """Normalize a WPF Key or string for code matching."""
    text = str(value or "").strip().upper()
    aliases = {
        "RETURN": "ENTER",
        "KEY.ENTER": "ENTER",
        "KEY.RETURN": "ENTER",
        "KEY.UP": "UP",
        "KEY.DOWN": "DOWN",
        "KEY.LEFT": "LEFT",
        "KEY.RIGHT": "RIGHT",
    }
    text = aliases.get(text, text)
    if text.startswith("KEY."):
        text = text[4:]
    ignored = ("", "NONE", "SYSTEM", "LEFTSHIFT", "RIGHTSHIFT", "LEFTCTRL", "RIGHTCTRL", "LEFTALT", "RIGHTALT")
    return "" if text in ignored else text


def display_unlock_key(value):
    key = normalize_unlock_key(value)
    return {"UP": "↑", "DOWN": "↓", "LEFT": "←", "RIGHT": "→", "ENTER": "Enter"}.get(key, key)


def match_unlock_code(keys):
    """Return the matching code when any registered sequence is a suffix."""
    normalized = tuple(normalize_unlock_key(key) for key in list(keys or ()))
    for code in THEME_UNLOCK_CODES:
        if len(normalized) >= len(code.Sequence) and normalized[-len(code.Sequence):] == code.Sequence:
            return code
    return None


def unlock_theme(code_or_unlock_id, config=None, save_callback=None):
    """Persist an unlock and return its theme descriptor, or None."""
    unlock_id = getattr(code_or_unlock_id, "UnlockId", code_or_unlock_id)
    unlock_id = str(unlock_id or "").strip().lower()
    descriptor = None
    for candidate in ALL_THEME_DESCRIPTORS:
        if candidate.UnlockId == unlock_id:
            descriptor = candidate
            break
    cfg = _get_config(config)
    if descriptor is None or cfg is None:
        return None
    unlock_ids = unlocked_theme_ids(config=cfg, migrate=True, save_callback=save_callback)
    unlock_ids.add(unlock_id)
    _write_unlock_ids(cfg, unlock_ids)
    if not _save_config(save_callback=save_callback):
        return None
    return descriptor
