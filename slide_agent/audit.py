"""Inspect generated geometry and editable OOXML parts, without claiming rendering."""
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from pptx import Presentation
from pptx.util import Inches


def audit(path, theme):
    prs = Presentation(str(path))
    issues, counts = [], {"slides": len(prs.slides), "text_shapes": 0, "tables": 0, "charts": 0, "pictures": 0}
    for i, slide in enumerate(prs.slides, 1):
        content_shapes = 0
        text_boxes = []
        for shape in slide.shapes:
            if min(shape.left, shape.top) < -2 or shape.left + shape.width > prs.slide_width + 2 or shape.top + shape.height > prs.slide_height + 2:
                issues.append(f"slide {i}: out-of-bounds shape {shape.name}")
            if shape.has_chart:
                counts["charts"] += 1
                content_shapes += 1
            if shape.has_table:
                counts["tables"] += 1
                content_shapes += 1
                frames = [(cell.text_frame, theme.table_pt) for row in shape.table.rows for cell in row.cells]
            elif shape.has_text_frame and shape.text:
                if shape.name in ("body", "title", "lead", "label", "footer"):
                    text_boxes.append(shape)
                counts["text_shapes"] += 1
                content_shapes += int(shape.name in ("body", "title"))
                minimum = theme.title_pt if shape.name == "title" else theme.body_pt if shape.name == "body" else 11
                frames = [(shape.text_frame, minimum)]
            else:
                frames = []
            if shape.shape_type == 13:
                counts["pictures"] += 1
            for tf, minimum in frames:
                for p in tf.paragraphs:
                    for run in p.runs:
                        if run.font.size is None or run.font.size.pt < minimum:
                            issues.append(f"slide {i}: missing/undersized font in {shape.name}")
        if not content_shapes:
            issues.append(f"slide {i}: no native editable content")
        for a_index, a in enumerate(text_boxes):
            for b in text_boxes[a_index+1:]:
                overlap_x = min(a.left+a.width, b.left+b.width) - max(a.left, b.left)
                overlap_y = min(a.top+a.height, b.top+b.height) - max(a.top, b.top)
                if overlap_x > 2 and overlap_y > 2:
                    issues.append(f"slide {i}: overlapping text boxes {a.name}/{b.name}")
    with ZipFile(path) as archive:
        if archive.testzip():
            issues.append("invalid ZIP CRC")
        for name in archive.namelist():
            if name.endswith((".xml", ".rels")):
                root = etree.fromstring(archive.read(name))
                if name.startswith("ppt/slides/slide") and name.endswith(".xml") and root.xpath("//*[local-name()='normAutofit' or local-name()='spAutoFit']"):
                    issues.append(f"auto-resize present in {name}")
                if name.endswith(".rels") and root.xpath("//*[@TargetMode='External']"):
                    issues.append(f"unexpected external relationship in {name}")
        counts["embedded_workbooks"] = sum(name.startswith("ppt/embeddings/") and name.endswith(".xlsx") for name in archive.namelist())
        if counts["embedded_workbooks"] != counts["charts"]:
            issues.append("native charts must each have an editable embedded workbook")
    return {"ok": not issues, "issues": issues, "counts": counts, "rendering": "not_performed"}
