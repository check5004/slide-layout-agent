"""Trusted, versioned native template registry. No network, Office or font dependency."""
import hashlib
import json
from functools import lru_cache
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
CATALOG = ROOT / "catalog"
INVENTORY = json.loads((CATALOG / "inventory.json").read_text(encoding="utf-8"))
REFERENCE_IDS = tuple(item["layout_id"] for item in INVENTORY["layouts"])
from .editorial_specs import LAYOUT_IDS as EDITORIAL_IDS
LAYOUT_IDS = REFERENCE_IDS + EDITORIAL_IDS
LICENSE_NOTICE = (CATALOG / "upstream/LICENSE").read_text(encoding="utf-8")


@lru_cache(maxsize=1)
def registry():
    data = json.loads((CATALOG / "manifest.json").read_text(encoding="utf-8"))
    return {entry["layout_id"]: entry for entry in data["layouts"]}


@lru_cache(maxsize=1)
def editorial_registry():
    data = json.loads((CATALOG / "editorial/manifest.json").read_text(encoding="utf-8"))
    return {entry["layout_id"]: entry for entry in data["layouts"]}


def all_registry():
    return {**registry(), **editorial_registry()}


def template_path(entry):
    path = (ROOT / entry["template"]).resolve()
    root = CATALOG / ("editorial/templates" if entry.get("catalog") == "editorial" else "templates")
    if not path.is_relative_to(root.resolve()):
        raise ValueError("template outside trusted catalog")
    if hashlib.sha256(path.read_bytes()).hexdigest() != entry["template_sha256"]:
        raise ValueError("template hash changed; rebuild and review catalog")
    return path


def is_catalog(layout_id):
    return layout_id in LAYOUT_IDS
