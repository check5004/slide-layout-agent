import argparse
import hashlib
import json
import re
import shutil
import sys
import tempfile
import uuid
from pathlib import Path

from pydantic import ValidationError

from .audit import audit
from .io import fingerprint, load_plan, load_source, load_theme, write_json
from .models import Asset, Plan, Segment, Source
from .render import render
from .validate import BODY_H, fits, validate


def ingest(path, mode, image_args, destination):
    path = Path(path)
    if path.stat().st_size > 2_000_000:
        raise ValueError("text exceeds 2 MB")
    raw = path.read_text(encoding="utf-8-sig")
    segments, group = [], "document"
    title = "入力文書"
    blocks = re.split(r"\n\s*\n", raw.strip()) if mode == "prose" else re.split(r"(?m)(?=^## )|\n\s*\n", raw.strip())
    group_n = 0
    for block in blocks:
        block = block.strip()
        if not block:
            continue
        if mode == "slides" and block.startswith("## "):
            heading, _, rest = block.partition("\n")
            group_n += 1
            group = f"slide{group_n}"
            if group_n == 1:
                title = heading[3:]
            segments.append(Segment(id=f"s{len(segments)+1:03d}", text=heading[3:], group=group))
            block = rest.strip()
        if block:
            segments.append(Segment(id=f"s{len(segments)+1:03d}", text=block, group=group))
    if mode == "slides" and group_n == 0:
        raise ValueError("slides mode requires at least one Markdown ## heading")
    images = []
    base = Path(destination).parent.resolve()
    for spec in image_args:
        image_id, delimiter, filename = spec.partition("=")
        if not delimiter:
            raise ValueError("--image requires ID=PATH")
        p = Path(filename).resolve()
        if not p.is_relative_to(base):
            raise ValueError("place images inside the source.json directory")
        if p.stat().st_size > 20_000_000:
            raise ValueError("image exceeds 20 MB")
        images.append(Asset(id=image_id, path=p.relative_to(base).as_posix(), sha256=hashlib.sha256(p.read_bytes()).hexdigest()))
    return Source(title=title, segments=segments, images=images)


def split_plan(plan, theme):
    output = []
    for slide in plan.slides:
        if slide.layout_id == "table":
            groups = [slide.contents.rows[i:i+6] for i in range(0, len(slide.contents.rows), 6)]
            key = "rows"
        elif slide.layout_id == "bullets":
            groups, pending = [], list(slide.contents.items)
            while pending:
                count = min(5, len(pending))
                while count and not all(fits(t.text, 11.5, BODY_H / count - 0.12, theme.body_pt) for t in pending[:count]):
                    count -= 1
                if count == 0:
                    raise ValueError(f"{slide.id}: one item exceeds a full slide; split that text in the plan with retained references")
                groups.append(pending[:count])
                pending = pending[count:]
            key = "items"
        else:
            output.append(slide.model_dump(mode="json"))
            continue
        for index, group in enumerate(groups):
            page = slide.model_dump(mode="json")
            page["id"] = slide.id if index == 0 else f"{slide.id}-p{index+1}"
            page["contents"][key] = [v.model_dump(mode="json") if hasattr(v, "model_dump") else [t.model_dump(mode="json") for t in v] for v in group]
            if len(groups) > 1:
                page["rationale"] += f" [pagination {index+1}/{len(groups)} from {slide.id}; original order retained]"
            output.append(page)
    data = plan.model_dump(mode="json")
    data["slides"] = output
    return Plan.model_validate(data)


def review(plan, source, report):
    lines = [f"# {plan.title}", "", f"原文 SHA-256: `{fingerprint(source)}`", "", "これは構成確認用。描画確認はまだ行っていません。", ""]
    for i, slide in enumerate(plan.slides, 1):
        lines.extend([f"## {i}. {slide.title.text}", "", f"layout: `{slide.layout_id}` / id: `{slide.id}` / origin: `{slide.origin}`", "", slide.rationale, "", "```json", json.dumps(slide.contents.model_dump(mode="json"), ensure_ascii=False, indent=2), "```", ""])
    lines.extend(["## 検証", "", "```json", json.dumps(report, ensure_ascii=False, indent=2), "```", ""])
    return "\n".join(lines)


def build_parser():
    parser = argparse.ArgumentParser(description="Source-traceable editable PowerPoint workflow; no LLM/API calls")
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("ingest")
    p.add_argument("input")
    p.add_argument("--mode", choices=["prose", "slides"], default="prose")
    p.add_argument("--image", action="append", default=[])
    p.add_argument("--out", required=True)
    p.add_argument("--force", action="store_true")
    p = sub.add_parser("fingerprint")
    p.add_argument("source")
    p = sub.add_parser("schema")
    p.add_argument("--out", required=True)
    p.add_argument("--force", action="store_true")
    for name in ("validate", "review", "render", "split"):
        p = sub.add_parser(name)
        p.add_argument("plan")
        p.add_argument("--source", required=True)
        p.add_argument("--theme")
        if name != "validate":
            p.add_argument("--out", required=True)
            p.add_argument("--force", action="store_true")
        if name == "render":
            p.add_argument("--accept-omissions", action="store_true")
    return parser


def main(argv=None):
    args = build_parser().parse_args(argv)
    try:
        if args.command == "ingest":
            source = ingest(args.input, args.mode, args.image, args.out)
            write_json(args.out, source.model_dump(mode="json"), args.force)
            print(f"Source written; SHA-256: {fingerprint(source)}")
            return 0
        if args.command == "schema":
            write_json(args.out, Plan.model_json_schema(), args.force)
            return 0
        source = load_source(args.source)
        if args.command == "fingerprint":
            print(fingerprint(source))
            return 0
        plan, theme = load_plan(args.plan), load_theme(args.theme)
        if args.command == "split":
            plan = split_plan(plan, theme)
        report = validate(plan, source, Path(args.source).parent, theme)
        if args.command == "validate":
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return 0 if report["ok"] else 2
        dest = Path(args.out)
        if dest.exists() and not args.force:
            raise ValueError("output exists; use --force")
        dest.parent.mkdir(parents=True, exist_ok=True)
        if args.command == "review":
            dest.write_text(review(plan, source, report), encoding="utf-8")
            return 0 if report["ok"] else 2
        if not report["ok"]:
            print(json.dumps(report, ensure_ascii=False, indent=2))
            return 2
        if args.command == "split":
            write_json(dest, plan.model_dump(mode="json"), args.force)
            return 0
        if plan.omissions and not args.accept_omissions:
            raise ValueError("explicit omissions require review and --accept-omissions")
        report_path = dest.with_suffix(".validation.json")
        if report_path.exists() and not args.force:
            raise ValueError("validation output exists; use --force")
        # Only publish the file after its structural audit succeeds.
        with tempfile.TemporaryDirectory(prefix="slide-agent-", dir=dest.parent) as temp:
            temp_pptx = Path(temp) / "deck.pptx"
            render(plan, source, Path(args.source).parent, theme, temp_pptx)
            report["ooxml"] = audit(temp_pptx, theme)
            if not report["ooxml"]["ok"]:
                print(json.dumps(report, ensure_ascii=False, indent=2))
                return 2
            # A Windows TemporaryDirectory has an owner-only ACL. Moving its
            # child would keep that ACL and prevent the user's PowerPoint from
            # reading the result. Copy into the destination directory first so
            # the normal inherited permissions apply, then publish atomically.
            staged = dest.with_name(f".{dest.name}.{uuid.uuid4().hex}.pending")
            try:
                shutil.copyfile(temp_pptx, staged)
                staged.replace(dest)
            finally:
                if staged.exists():
                    staged.unlink()
        write_json(report_path, report, args.force)
        print(f"Rendered {len(plan.slides)} slides. Structural checks passed; real visual review is still required.")
        return 0
    except (ValueError, OSError, ValidationError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
