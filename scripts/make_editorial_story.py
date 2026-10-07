"""Reproduce a Codex-authored fictional internal research briefing (16 slides).

The family/content choices below are authored fixtures, not an autonomous AI.
Four adjacent pages intentionally begin with the same three-card structure.
"""
import copy
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from scripts.make_editorial_samples import editorial_sample_plan
from slide_agent.catalog import editorial_registry
from slide_agent.selection import select_variants
from slide_agent.models import Theme
from slide_agent.validate import validate
from slide_agent.render import render
from slide_agent.audit import audit
from slide_agent.cli import review

DEST = ROOT / 'examples/editorial-story'
LAYOUTS = ['ed_cover_brief_2', 'ed_agenda_rows_3'] + ['ed_cards_grid_3'] * 4 + [
    'ed_cards_grid_6', 'ed_diagram_left_3', 'ed_comparison_columns_2', 'ed_steps_vertical_4',
    'ed_timeline_horizontal_3', 'ed_metrics_breakdown_top_3', 'ed_table_three_columns_4',
    'ed_chart_bar_left_4', 'ed_quote_left_2', 'ed_summary_rows_3']
TITLES = ['社内ナレッジ共有の調査メモ', '資料共有を整理するための三つの観点',
          '見つけやすさは入口の設計で変わる', '検索で詰まる場面を分けて考える',
          '要約から根拠へ戻れることが必要', '更新を続けるための役割を整理する',
          '運用前に確認する六つの観点', '資料から共有メモへ変換する流れ',
          '集約型と分散型は運用条件で選ぶ', '小さく試し、条件を揃えて見直す',
          '段階を分けて確認する', '文書の確認状況を棚卸しする',
          '確認項目と残る留保を一覧にする', '言及件数だけで優先度を決めない',
          '発言と分析者の解釈を区別する', '共有時に残すべき要点を整理する']
CARD_PAGES = [
    [('入口をまとめる', '部署ごとの入口を一覧化し、どの資料を探す場所かを明記する。既存の原文はそのまま保管する。', '対象：検索の入口'),
     ('呼称を対応させる', '同じ対象を示す用語の違いを整理する。専門用語を置き換えず、別の呼び方からも辿れるようにする。', '根拠：用語の照合'),
     ('範囲を先に見せる', '検索結果に対象部署と適用範囲を添える。本文を開く前に、自分の状況に関係する資料かを判断できる。', '留保：効果は未検証')],
    [('場所が分からない', '資料の保管先が部署ごとに異なる。入口の整理と、原文を管理する責任の所在は分けて扱う。', '観点：保管場所'),
     ('用語が一致しない', '検索語と資料内の呼称が一致しない。索引に同義の表現を残し、業務上の定義の違いも確認する。', '観点：用語の違い'),
     ('使えるか迷う', '資料は見つかっても適用条件が分からない。更新日だけでなく、対象範囲と確認先も示す。', '留保：場面差あり')],
    [('引用箇所を残す', '要約と原文の位置を対応させる。後から読んだ人が、結論の根拠となった条件を確かめられる形にする。', '根拠：原文を保持'),
     ('留保を近くに置く', '限定条件や未確認事項を本文の近くに残す。脚注だけに追いやらず、判断する場面で読める位置に置く。', '観点：読み落とし'),
     ('解釈を分けて示す', '資料に書かれた内容と分析者の解釈を区別する。推測を確定事項に言い換えず、確認が必要な点を残す。', '留保：解釈を区別')],
    [('原文の担当を決める', '原文を更新する担当と、共有メモを確認する担当を明記する。どちらかだけが更新される状態を見つける。', '観点：担当の分担'),
     ('更新の契機を決める', '定期確認に加え、制度や対象範囲が変わった際の見直しを決める。頻度は業務の変化に応じて調整する。', '留保：頻度は要調整'),
     ('保留の理由を残す', '確認が終わらない論点は理由と確認先を残す。無理に結論を埋めず、次の担当者が続きを調べられるようにする。', '観点：引き継ぎ')],
]


def story():
    registry = editorial_registry()
    entries = [copy.deepcopy(registry[key]) for key in LAYOUTS]
    for index, entry in enumerate(entries):
        entry['title']['sample'] = TITLES[index]
    for page, triples in enumerate(CARD_PAGES, 2):
        for i, (heading, body, evidence) in enumerate(triples, 1):
            for suffix, value in [('heading', heading), ('body', body), ('evidence', evidence)]:
                entries[page]['texts'][f'item_{i}_{suffix}']['sample'] = value
    source, initial = editorial_sample_plan(entries, DEST)
    source.title = initial.title = '架空調査：社内ナレッジ共有'
    from slide_agent.io import fingerprint
    initial.source_sha256 = fingerprint(source)
    for slide, entry in zip(initial.slides, entries):
        slide.repetition_reason = None
        slide.rationale = (f"{entry['name']}を選択。原文の{entry['item_count']}項目と意味上の関係を保持し、"
                           '見出し・説明・留保を対応づける。架空の調査共有であり、現実の効果を主張しない。')
    selected, decisions = select_variants(initial)
    return source, initial, selected, decisions


def main():
    source, initial, plan, decisions = story()
    DEST.mkdir(parents=True, exist_ok=True)
    report = validate(plan, source, DEST, Theme())
    if not report['ok']: raise ValueError([i for i in report['issues'] if i['severity']=='error'])
    for name, model in [('source', source), ('initial.plan', initial), ('plan', plan)]:
        (DEST / f'story.{name}.json').write_text(model.model_dump_json(indent=2), encoding='utf-8')
    (DEST / 'input.txt').write_text('\n\n'.join(f'## {i+1}. {TITLES[i]}\n{segment.text}' for i, segment in enumerate(source.segments)), encoding='utf-8')
    (DEST / 'story.selection.json').write_text(json.dumps({'decisions': decisions, 'layout_selection': report['layout_selection']}, ensure_ascii=False, indent=2), encoding='utf-8')
    render(plan, source, DEST, Theme(), DEST / 'story.pptx')
    report['ooxml'] = audit(DEST / 'story.pptx', Theme())
    if not report['ooxml']['ok']: raise ValueError(report['ooxml'])
    (DEST / 'story.validation.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    (DEST / 'story.review.md').write_text(review(plan, source, report), encoding='utf-8')
    print('16-slide fictional research story: selected, validated, native editable PPTX generated')
    for item in decisions[2:6]: print(item['slide_id'], item['from'], '->', item['to'])


if __name__ == '__main__': main()
