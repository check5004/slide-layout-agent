"""Conservative preflight. This is not a substitute for a real renderer."""
import hashlib
import math
import re
import unicodedata
from collections import defaultdict
from decimal import Decimal

from PIL import Image
from pydantic import BaseModel

from .io import asset_path, fingerprint
from .models import Datum, Text
from .catalog import is_catalog
from .catalog_validate import validate_catalog

WIDTH, HEIGHT = 13.333333, 7.5
BODY_Y, BODY_H = 2.0, 4.55
MARGIN, CONTENT_W = 0.6, 12.133333


def walk_evidence(obj):
    if isinstance(obj, (Text, Datum)):
        yield obj
    elif isinstance(obj, BaseModel):
        for value in obj.__dict__.values():
            yield from walk_evidence(value)
    elif isinstance(obj, list):
        for value in obj:
            yield from walk_evidence(value)
    elif isinstance(obj, dict):
        for value in obj.values():
            yield from walk_evidence(value)


def numbers(text):
    text = unicodedata.normalize("NFKC", text)
    return {str(Decimal(s.replace(",", "")).normalize()) for s in re.findall(r"(?<![\w.])[-+]?\d[\d,]*(?:\.\d+)?|(?<=[^\x00-\x7f])[-+]?\d[\d,]*(?:\.\d+)?", text)}


def units(text):
    return sum(1 if unicodedata.east_asian_width(c) in "WF" else 0.6 for c in text)


def fits(text, width, height, pt):
    # 10% horizontal / 20% vertical safety; font-specific rendering comes later.
    capacity = max(1, (width * 72 - 12) / (pt * 1.10))
    lines = sum(max(1, math.ceil(units(line) / capacity)) for line in text.split("\n"))
    return lines * pt * 1.35 + 8 <= height * 72


def validate(plan, source, base, theme):
    issues = []

    def add(code, message, slide=None, severity="error"):
        issues.append({"severity": severity, "code": code, "slide": slide, "message": message})

    if plan.source_sha256 != fingerprint(source):
        add("SOURCE_CHANGED", "source hash does not match; review the original source before replanning")
    segments = {s.id: s for s in source.segments}
    assets = {a.id: a for a in source.images}
    seen, quotes, rendered_numbers = set(), defaultdict(list), defaultdict(set)
    omitted = {o.source_id for o in plan.omissions}
    if any(is_catalog(s.layout_id) for s in plan.slides):
        add("CATALOG_FONT_REVIEW", "Catalog preserves compact source typography. Font installation is not required for writing PPTX; the viewer may substitute fonts. Review the result on the target device.", severity="warning")
    for omission in plan.omissions:
        if omission.source_id not in segments:
            add("UNKNOWN_OMISSION", omission.source_id)
        add("OMISSION_REVIEW", f"{omission.source_id}: {omission.reason}", severity="warning")

    def capacity(text, w, h, pt, slide, label):
        if not fits(text, w, h, pt):
            add("TEXT_OVERFLOW", f"{label}: split/rewrite the plan; do not shrink below theme size", slide.id)

    for slide in plan.slides:
        if slide.origin not in {s.group for s in source.segments}:
            add("UNKNOWN_ORIGIN", slide.origin, slide.id)
        for item in walk_evidence(slide):
            evidence = []
            for ref in item.refs:
                segment = segments.get(ref.source_id)
                if not segment:
                    add("UNKNOWN_SOURCE", ref.source_id, slide.id)
                    continue
                if ref.quote not in segment.text:
                    add("QUOTE_MISMATCH", f"quote not found in {ref.source_id}", slide.id)
                    continue
                seen.add(ref.source_id)
                # Coverage is based on visible verbatim text, not a long evidence
                # quote which could hide omitted qualifiers. Paraphrases always
                # require explicit semantic review.
                if isinstance(item, Text) and item.text in segment.text:
                    quotes[ref.source_id].append(item.text)
                elif isinstance(item, Datum):
                    quotes[ref.source_id].append(ref.quote)
                evidence.append(ref.quote)
                rendered_numbers[ref.source_id].update(numbers(item.text) if isinstance(item, Text) else {str(Decimal(str(item.value)).normalize())})
            visible = item.text if isinstance(item, Text) else str(item.value)
            if numbers(visible) - numbers(" ".join(evidence)):
                add("UNSUPPORTED_NUMBER", "visible number absent from cited quotes", slide.id)
            if isinstance(item, Text):
                if item.mode == "verbatim" and not any(item.text in quote for quote in evidence):
                    add("NOT_VERBATIM", "verbatim text differs from every cited quote", slide.id)
                if item.mode == "paraphrase":
                    add("PARAPHRASE_REVIEW", "check meaning, qualifiers and causal relationships", slide.id, "warning")
        if is_catalog(slide.layout_id):
            validate_catalog(slide, add)
            continue
        capacity(slide.title.text, CONTENT_W, 1.0, theme.title_pt, slide, "title")
        if slide.lead:
            if not theme.lead_mode:
                add("LEAD_DISABLED", "lead requires a theme with lead_mode=true", slide.id)
            capacity(slide.lead.text, CONTENT_W, 0.45, 18, slide, "lead")
        c, kind = slide.contents, slide.layout_id
        if kind in ("title", "bullets", "closing", "text_image"):
            maximum = {"title": 3, "bullets": 5, "closing": 3, "text_image": 4}[kind]
            minimum = 0 if kind == "title" else 1
            if not minimum <= len(c.items) <= maximum:
                add("ITEM_LIMIT", f"{kind} requires {minimum}–{maximum} items; split the plan", slide.id)
            w = 5.65 if kind == "text_image" else (10.7 if kind in ("title", "closing") else 11.5)
            h = (3.5 if kind in ("title", "closing") else BODY_H) / max(1, len(c.items)) - 0.12
            for t in c.items:
                capacity(t.text, w, h, theme.body_pt, slide, "item")
        if kind == "text_image":
            if c.image_id not in assets:
                add("UNKNOWN_IMAGE", c.image_id, slide.id)
        elif kind == "comparison":
            for title, items in ((c.left_title, c.left), (c.right_title, c.right)):
                capacity(title.text, 5.35, 0.65, theme.body_pt, slide, "comparison heading")
                if len(items) > 3:
                    add("ITEM_LIMIT", "comparison permits at most three items per column", slide.id)
                for t in items:
                    capacity(t.text, 5.1, 3.2 / len(items) - 0.12, theme.body_pt, slide, "comparison item")
        elif kind == "process":
            if len(c.steps) > 4:
                add("ITEM_LIMIT", "process permits at most four steps", slide.id)
            for t in c.steps:
                if len(t.text) > 48:
                    add("TEXT_OVERFLOW", "process steps must be at most 48 characters", slide.id)
                capacity(t.text, 10.9 / len(c.steps) - 0.35, 2.1, theme.body_pt, slide, "step")
        elif kind == "table":
            if len(c.rows) > 6:
                add("ROW_LIMIT", "table permits at most six data rows; use split", slide.id)
            for row in [c.headers] + c.rows:
                for t in row:
                    capacity(t.text, CONTENT_W / len(c.headers) - 0.20, BODY_H / (len(c.rows) + 1) - 0.07, theme.table_pt, slide, "table cell")
        elif kind == "chart":
            for t in c.categories + [s.name for s in c.series]:
                if units(t.text) > 10:
                    add("CHART_LABEL_LIMIT", "chart labels must fit ten full-width characters", slide.id)
            capacity(c.note.text, CONTENT_W, 0.8, 18, slide, "chart note")

    for segment in source.segments:
        sid = segment.id
        if sid in seen and sid in omitted:
            add("OMISSION_CONFLICT", f"{sid} is both used and omitted")
        if sid not in seen and sid not in omitted:
            add("UNCOVERED_SOURCE", f"{sid} is neither cited nor explicitly omitted")
        if sid not in omitted and numbers(segment.text) - rendered_numbers[sid]:
            add("MISSING_NUMBER", f"{sid} has numbers missing from visible content")
        if sid in seen:
            # Exact quoted spans, not merely the existence of a source ID.
            covered = [False] * len(segment.text)
            for quote in quotes[sid]:
                start = segment.text.find(quote)
                for i in range(start, start + len(quote)):
                    covered[i] = True
            if any(not hit and ch.isalnum() for ch, hit in zip(segment.text, covered)):
                add("PARTIAL_SOURCE_REVIEW", f"{sid}: some original text is not represented by quoted evidence", severity="warning")
    for asset in source.images:
        if not any(s.layout_id == "text_image" and s.contents.image_id == asset.id for s in plan.slides):
            add("UNUSED_IMAGE_REVIEW", f"{asset.id}: supplied image is not used", severity="warning")
        try:
            path = asset_path(asset, base)
            if path.stat().st_size > 20_000_000:
                raise ValueError("image exceeds 20 MB")
            if hashlib.sha256(path.read_bytes()).hexdigest() != asset.sha256:
                raise ValueError("image hash changed")
            with Image.open(path) as im:
                if im.width * im.height > 40_000_000:
                    raise ValueError("image exceeds 40 megapixels")
                if im.format not in ("PNG", "JPEG"):
                    raise ValueError("image bytes must be PNG/JPEG")
                im.verify()
        except (OSError, ValueError, Image.DecompressionBombError) as exc:
            add("IMAGE_INVALID", f"{asset.id}: {type(exc).__name__}: {exc}")
    return {"ok": not any(i["severity"] == "error" for i in issues), "issues": issues, "slide_count": len(plan.slides), "visual_review": "not_performed"}
