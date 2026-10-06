"""Reproduce fictional fixtures. This script is deterministic, not an AI planner."""
import hashlib
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from PIL import Image, ImageDraw

from slide_agent.io import fingerprint, write_json
from slide_agent.models import Plan, Source

ROOT = Path(__file__).resolve().parents[1]


def make_demo(destination=None):
    root = Path(destination) if destination else ROOT / "examples"
    root.mkdir(parents=True, exist_ok=True)
    image = Image.new("RGB", (1200, 800), "#DBE8ED")
    draw = ImageDraw.Draw(image)
    for row in range(4):
        for col in range(6):
            x, y = 30 + col * 195, 30 + row * 190
            draw.rectangle((x, y, x + 165, y + 155), fill=("#007C83" if (row + col) % 2 else "#47699B"))
    draw.ellipse((470, 270, 730, 530), fill="#F5F7FA", outline="#13233B", width=10)
    draw.text((535, 385), "IMAGE FIXTURE", fill="#13233B")
    image.save(root / "grid.png")
    segments, slides = [], []

    def t(value, group):
        sid = f"s{len(segments)+1:03d}"
        segments.append({"id": sid, "text": value, "citation": "検証用の架空設定。実在の調査結果ではありません。", "group": group})
        return {"text": value, "refs": [{"source_id": sid, "quote": value}], "mode": "verbatim"}

    def slide(kind, title, contents_factory, why):
        group = f"section{len(slides)+1}"
        heading = t(title, group)
        slides.append({"id": f"demo{len(slides)+1}", "origin": group, "layout_id": kind,
                       "title": heading, "rationale": why, "contents": contents_factory(lambda v: t(v, group))})

    slide("title", "地域図書室の運営改善", lambda t: {"items": [t("架空の活動を使ったレイアウト検証"), t("数値・団体・活動はすべて架空の設定です。")]}, "表紙。架空設定を明示する。")
    slide("bullets", "運営の前提をそろえる", lambda t: {"items": [t("利用者の要望を記録し、担当者間で共有する。"), t("既存の資料を整理し、必要な情報を見つけやすくする。"), t("小さな試行の結果を確認してから、運用を見直す。") ]}, "同格の前提を箇条書きで示す。")
    slide("text_image", "記録と資料を並べて確認", lambda t: {"items": [t("右側は添付画像の比率を検証する図です。"), t("画像全体を表示し、縦横比を保ちます。"), t("説明文は PowerPoint 上で編集できます。")], "image_id": "grid", "image_mode": "fit"}, "支給画像と説明の対比。fitで全体を保持。")
    slide("comparison", "受付方法の比較", lambda t: {"left_title": t("紙の受付"), "right_title": t("共通の記録表"), "left": [t("その場で書き込める。"), t("集計には転記が必要。")], "right": [t("担当者間で参照しやすい。"), t("入力手順の共有が必要。") ]}, "入力にある二つの選択肢を対置する。")
    slide("process", "試行から振り返りへ", lambda t: {"steps": [t("要望の整理"), t("小さく試す"), t("結果の記録"), t("運用の見直し")]}, "原文の順序をnative図形と矢印で示す。")
    slide("table", "試行の予定を共有する", lambda t: {"headers": [t("段階"), t("作業"), t("確認内容")], "rows": [[t("準備"), t("記録表を作成"), t("項目の不足")], [t("試行"), t("受付で利用"), t("記入の負担")], [t("振り返り"), t("記録を確認"), t("改善する箇所")]]}, "既存の対応関係を編集可能な表にする。")

    def chart_contents(t):
        cats = [t("試行前"), t("第1回"), t("第2回")]
        name = t("記録件数")
        values = []
        for value in (12, 18, 24):
            evidence = t(f"{value}")
            values.append({"value": float(value), "refs": evidence["refs"]})
        return {"categories": cats, "series": [{"name": name, "values": values}], "note": t("単位：件。架空の数値であり、改善効果を示すものではありません。")}

    slide("chart", "記録件数を確認する", chart_contents, "与えられた件数だけをnative縦棒グラフにする。因果を補完しない。")
    slide("closing", "次の試行に向けて", lambda t: {"items": [t("記録の取り方を担当者間で確認する。"), t("不明な点を残したまま効果を断定しない。"), t("運用に合わせて文章や表を手直しする。") ]}, "入力にある次の行動をまとめる。")
    source = Source.model_validate({"title": "地域図書室の運営改善（架空例）", "segments": segments,
                                   "images": [{"id": "grid", "path": "grid.png", "caption": "Original synthetic aspect-ratio fixture", "sha256": hashlib.sha256((root / "grid.png").read_bytes()).hexdigest()}]})
    plan = Plan.model_validate({"source_sha256": fingerprint(source), "title": source.title, "slides": slides})
    write_json(root / "demo.source.json", source.model_dump(mode="json"), force=True)
    write_json(root / "demo.plan.json", plan.model_dump(mode="json"), force=True)
    return source, plan


if __name__ == "__main__":
    make_demo()
    print("Wrote fictional source, deterministic layout fixture and original image fixture.")
