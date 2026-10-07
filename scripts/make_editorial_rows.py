"""Matching-content row-alignment fixtures, including explicit top/middle choices."""
import copy
import json
import sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from scripts.make_editorial_samples import editorial_sample_plan
from slide_agent.catalog import editorial_registry
from slide_agent.models import Theme
from slide_agent.validate import validate
from slide_agent.render import render
from slide_agent.audit import audit

DEST=ROOT/'examples/editorial-rows'
LAYOUTS=[f'ed_cards_rows_{n}' for n in range(2,7)]+[
    'ed_bullets_rows_3','ed_steps_vertical_4','ed_timeline_vertical_3',
    'ed_metrics_stack_3','ed_table_three_columns_4','ed_cards_rows_3','ed_cards_rows_3']


def row_fixture():
    entries=[copy.deepcopy(editorial_registry()[key]) for key in LAYOUTS]
    for page,entry in enumerate(entries,1):
        entry['title']['sample']='短い見出しと説明を同じ行のまとまりで読む'
        if page==11:entry['title']['sample']='行ごとに上寄せと中央寄せを指定する'
        if page==12:entry['title']['sample']='読み始めを揃える説明は上寄せを選ぶ'
        for key,spec in entry['texts'].items():
            if key.startswith('item_'):
                item=int(key.split('_')[1]);role=key.split('_')[2]
                if role=='heading':
                    value='確認の入口' if item%3==1 else '原文の条件を整理' if item%3==2 else '対象範囲と\n例外条件を確認'
                elif role=='body':
                    value='原文を確認する。' if item%3==1 else '確認日と対象範囲を揃える。\n条件の違いも残す。' if item%3==2 else '資料への参照を残す。\n未確認の条件は分けて扱う。\n次の確認先も記録する。'
                elif role=='evidence':
                    value='根拠：架空資料' if item%2 else '根拠：架空資料\n留保：効果は未検証'
                else:continue
                # Preserve the established capacity. Dense four-row body slots
                # have one line; no hidden shrink to force a multiline example.
                spec['sample']='\n'.join(value.split('\n')[:spec['max_lines']])
    source,plan=editorial_sample_plan(entries,DEST)
    plan.slides[10].row_alignments={'item_1':'top','item_2':'middle','item_3':'top'}
    plan.slides[11].row_alignments={f'item_{i}':'top' for i in range(1,4)}
    return source,plan


def main():
    source,plan=row_fixture();DEST.mkdir(exist_ok=True)
    report=validate(plan,source,DEST,Theme())
    if not report['ok']:raise ValueError(report['issues'])
    (DEST/'rows.source.json').write_text(source.model_dump_json(indent=2),encoding='utf-8')
    (DEST/'rows.after.plan.json').write_text(plan.model_dump_json(indent=2),encoding='utf-8')
    before=plan.model_dump(mode='json')
    for slide in before['slides']:slide.pop('row_alignments',None)
    (DEST/'rows.before.plan.json').write_text(json.dumps(before,ensure_ascii=False,indent=2),encoding='utf-8')
    render(plan,source,DEST,Theme(),DEST/'rows.after.pptx')
    report['ooxml']=audit(DEST/'rows.after.pptx',Theme())
    if not report['ooxml']['ok']:raise ValueError(report['ooxml'])
    (DEST/'rows.validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print('12 row fixtures: mixed line lengths, counts 2-6, evidence, numbers, dates, metrics and table; native audit passed')


if __name__=='__main__':main()
