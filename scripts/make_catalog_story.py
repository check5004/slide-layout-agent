"""Reproduce a Codex-authored fictional plan; this is not an AI planner."""
import json
import sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from slide_agent.io import fingerprint
from slide_agent.models import Plan,Source,Theme
from slide_agent.validate import validate
from slide_agent.render import render
from slide_agent.audit import audit


def make_story(destination=None):
    out=Path(destination) if destination else ROOT/'examples/catalog-story'
    out.mkdir(parents=True,exist_ok=True)
    segments=[]
    def text(value,group):
        sid=f's{len(segments)+1:03}'
        segments.append({'id':sid,'text':value,'group':group,'citation':'すべて架空の図書室・件数・提案。実在の調査結果ではありません。'})
        return {'text':value,'refs':[{'source_id':sid,'quote':value}],'mode':'verbatim'}
    def slide(sid,layout,title,slots,rationale):
        return {'id':sid,'origin':sid,'layout_id':layout,'title':text(title,sid),'rationale':rationale,
                'contents':{'texts':{k:text(v,sid) for k,v in slots.items()},'charts':{},'metrics':{},'states':{}}}
    slides=[slide('intro','title_page','図書室の記録改善',
                  {'text_007':'架空の試行記録を用いた構成例'},'表紙の大きな見出しで目的と架空設定を示す。'),
            slide('evidence','chart_insight','記録件数と判断の留保',
                  {'text_006':'件数は架空です。増加の原因や改善効果は確認していません。',
                   'text_010':'受付記録件数（件）','text_025':'効果はまだ断定しない',
                   'text_026':'記録件数の推移を確認します。\n増加の原因は未確認です。'},
                  '4時点のnative chartと右側の留保を組み合わせ、件数増加を因果と取り違えない。'),
            slide('decision','decision_page','次の試行で確認すること',
                  {'text_006':'架空の運用案です。実施結果を示すものではありません。','text_011':'提案する次の行動',
                   'text_012':'記録手順をそろえて\n小さく試行する',
                   'text_013':'1','text_014':'入力項目を担当者で確認する',
                   'text_016':'2','text_017':'同じ手順で記録を続ける',
                   'text_019':'3','text_020':'記入負担と欠落を振り返る'},
                  '提案1点と確認手順3点を分け、意思決定の構造に合わせる。')]
    values=[]
    for v in [12,18,21,24]:
        evidence=text(str(v),'evidence'); values.append({'value':v,'refs':evidence['refs']})
    slides[1]['contents']['charts']['chart_028']={
        'categories':[text(x,'evidence') for x in ['試行前','初回','中間','最終']],
        'series':[{'name':text('受付記録','evidence'),'values':values,'x_values':[]}]}
    source=Source.model_validate({'title':'図書室の記録改善（架空例）','segments':segments})
    plan=Plan.model_validate({'title':source.title,'source_sha256':fingerprint(source),'slides':slides})
    report=validate(plan,source,out,Theme())
    if not report['ok']: raise ValueError(report['issues'])
    for name,obj in [('source',source),('plan',plan)]:
        (out/f'story.{name}.json').write_text(obj.model_dump_json(indent=2),encoding='utf-8')
    (out/'input.txt').write_text('\n\n'.join(s.text for s in source.segments)+'\n',encoding='utf-8')
    dest=out/'story.pptx';render(plan,source,out,Theme(),dest)
    report['ooxml']=audit(dest,Theme())
    if not report['ooxml']['ok']:raise ValueError(report['ooxml'])
    (out/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    return source,plan


if __name__=='__main__':
    make_story();print('Fictional authored plan: 3 selected saved templates, source-backed native output.')
