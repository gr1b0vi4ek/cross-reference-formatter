"""Портативный конфиг: JSON рядом с exe."""

import json
import sys
from pathlib import Path


def _get_app_dir() -> Path:
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _get_config_path() -> Path:
    app_dir = _get_app_dir()
    portable = app_dir / 'crf_config.json'
    try:
        test = app_dir / '.crf_write_test'
        test.touch()
        test.unlink()
        return portable
    except (OSError, PermissionError):
        return Path.home() / '.crf_config.json'


CONFIG_FILE = _get_config_path()


def load_config() -> dict:
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return {}


def save_config(cfg: dict) -> None:
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except OSError:
        pass