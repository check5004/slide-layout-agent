"""Freeze the complete upstream inventory; never fetch or execute source HTML."""
import argparse
import hashlib
import json
import re
import shutil
import subprocess
from pathlib import Path

from lxml import html

ROOT = Path(__file__).resolve().parents[1]
COMMIT = "e5dd04dc6f1acf010046d2fc7319bcf5242a849a"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("upstream", type=Path)
    parser.add_argument("--git", default="git")
    args = parser.parse_args()
    upstream = args.upstream.resolve()
    head = subprocess.check_output([args.git, "-C", str(upstream), "rev-parse", "HEAD"], text=True).strip()
    if head != COMMIT:
        raise ValueError(f"expected {COMMIT}; got {head}")
    files = subprocess.check_output([args.git, "-C", str(upstream), "ls-files"], text=True).splitlines()
    catalog = (upstream / "references/archetype-catalog.md").read_text(encoding="utf-8")
    rows = re.findall(r"^\| (\d+) \| `([^`]+)` \| ([^|]+) \|(.+)$", catalog, re.M)
    assert len(rows) == 62
    records, sourcefiles = [], []
    for rel in files:
        if not rel.endswith(".html"):
            continue
        data = (upstream / rel).read_bytes()
        root = html.fromstring(data)
        sections = root.xpath("//section")
        is_layout = rel.startswith("templates/")
        sourcefiles.append({"path": rel, "sha256": hashlib.sha256(data).hexdigest(), "sections": len(sections),
                            "classification": "layout_catalog" if is_layout else "test_fixture_not_layout"})
        if not is_layout:
            continue
        offset, prefix = (27, "m") if "more" in rel else (0, "b")
        for i, section in enumerate(sections):
            _, key, title, tail = rows[offset + i]
            serialized = html.tostring(section, encoding="utf-8")
            records.append({"layout_id": key, "source_file": rel, "source_section": i + 1,
                            "source_selector": f"section:nth-of-type({i + 1})", "part_id": f"{prefix}{i+1:02}",
                            "name": title.strip(), "source_section_sha256": hashlib.sha256(serialized).hexdigest(),
                            "source_svg_count": len(section.xpath(".//svg")), "source_table_count": len(section.xpath(".//table")),
                            "source_image_count": len(section.xpath(".//img")),
                            "variants": ["warm", "cool"], "status": "inventoried"})
    dest = ROOT / "catalog" / "upstream"
    dest.mkdir(parents=True, exist_ok=True)
    copied = ["LICENSE", "references/archetype-catalog.md", "assets/SuperTemplate_62type.pptx"]
    copied += [f["path"] for f in sourcefiles]
    for rel in copied:
        target = dest / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(upstream / rel, target)
    result = {"repository": "https://github.com/carnot-tech/consulting-pptx-skill", "commit": COMMIT,
              "license": "MIT", "copyright": "Copyright (c) 2026 Carnot AI Inc.",
              "layout_count": len(records), "render_variant_count": 2 * len(records),
              "source_html_files": sourcefiles,
              "non_layouts": {"shared_partials": [], "shared_styles": "inline CSS in both template documents",
                              "fixture_sections": 6, "pptx_catalog_navigation_slides": [1, 2, 3, 13, 27, 40, 54, 64]},
              "variant_note": "warm is active upstream CSS; cool is the documented commented palette alternative, not a new structural layout",
              "layouts": records}
    (ROOT / "catalog" / "inventory.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Pinned {head}: {len(records)} layouts; {len(sourcefiles)} HTML files")


if __name__ == "__main__":
    main()
