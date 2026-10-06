import hashlib
import json
from pathlib import Path

from .models import Plan, Source, Theme


def _no_duplicate_keys(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def read_json(path):
    path = Path(path)
    if path.stat().st_size > 5_000_000:
        raise ValueError("JSON exceeds the 5 MB limit")
    return json.loads(path.read_text(encoding="utf-8-sig"), object_pairs_hook=_no_duplicate_keys)


def load_source(path):
    return Source.model_validate(read_json(path))


def load_plan(path):
    return Plan.model_validate(read_json(path))


def load_theme(path=None):
    return Theme.model_validate(read_json(path)) if path else Theme()


def fingerprint(source):
    canonical = json.dumps(source.model_dump(mode="json"), ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def write_json(path, data, force=False):
    path = Path(path)
    if path.exists() and not force:
        raise ValueError(f"output exists; use --force: {path.name}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def asset_path(asset, base):
    path = Path(asset.path)
    if path.is_absolute() or ":" in asset.path or "\\" in asset.path:
        raise ValueError("images require portable relative paths")
    base = Path(base).resolve()
    resolved = (base / path).resolve()
    if not resolved.is_relative_to(base):
        raise ValueError("image escapes the source folder")
    if resolved.suffix.lower() not in (".png", ".jpg", ".jpeg"):
        raise ValueError("only PNG/JPEG images are supported")
    return resolved
