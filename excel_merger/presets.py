import json
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
PRESET_FILE = PROJECT_ROOT / "presets" / "settings_presets.json"


def load_presets(path: Path = PRESET_FILE) -> dict[str, dict]:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {str(name): value for name, value in data.items() if isinstance(value, dict)}


def save_presets(presets: dict[str, dict], path: Path = PRESET_FILE) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(presets, indent=2, sort_keys=True), encoding="utf-8")


def upsert_preset(name: str, payload: dict, path: Path = PRESET_FILE) -> None:
    presets = load_presets(path)
    presets[name] = payload
    save_presets(presets, path)


def delete_preset(name: str, path: Path = PRESET_FILE) -> bool:
    presets = load_presets(path)
    if name not in presets:
        return False
    del presets[name]
    save_presets(presets, path)
    return True
