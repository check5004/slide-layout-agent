"""Fictional structure fixtures only; never imports private research material."""
import argparse
import json
import math
import sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from slide_agent.catalog import relation_registry
from slide_agent.models import Source,Plan,Theme
from slide_agent.io import fingerprint
from slide_agent.relations import sequence_geometry,bounds_capacity
from slide_agent.validate import validate
from slide_agent.render import render
from slide_agent.audit import audit


def fixture(entries, stress=False, qa=False, usage=False):
    segments=[];slides=[]
    for page,entry in enumerate(entries,1):
        group=f'page{page:02}';kind=entry['relation']['kind'];n=entry['item_count']
        def text(value,quote=None):
            sid=f's{len(segments)+1:04}';quote=quote or value
            segments.append({'id':sid,'text':quote,'group':group,'citation':'架空設計書・構造検証用'})
            return {'text':value,'refs':[{'source_id':sid,'quote':quote}],'mode':'verbatim'}
        def fill(spec,value):
            return '\n'.join(['調'*math.floor(spec['max_units_per_line'])]*spec['max_lines']) if stress else value
        title=fill(entry['title'],'主体の関係と通信の向きを\n根拠とともに確認する')
        contents={'texts':{k:text(fill(s,s['sample'])) for k,s in entry['texts'].items()}}
        if kind=='compare':
            r=entry['relation'];axes=['識別の扱い','確認する主体','更新時の条件']
            values=[['事前に識別子を登録','受付側が登録内容を照合','登録内容を更新'],['共通の参照先を提示','照会側が参照先を確認','参照先の内容を更新'],['処理ごとに識別子を渡す','各処理の担当が確認','処理条件を引き継ぐ']]
            contents['comparison']={'axes':[text(fill(r['cells'][f'{i+1}:0'],v)) for i,v in enumerate(axes)],'methods':[
                {'id':f'method_{j}','label':text(fill(r['cells'][f'0:{j+1}'],f'方式{chr(65+j)}')),
                 'values':[text(fill(r['cells'][f'{i+1}:{j+1}'],v)) for i,v in enumerate(values[j])]} for j in range(3)]}
        else:
            if kind=='spoke':labels=['受付機能','利用者','照会先','発行窓口']
            elif kind=='pair_fan':labels=['依頼元','照会先','配信元','宛先A','宛先B']
            elif kind=='fan' and n==3:labels=['https://id.test/x','主体A','主体B']
            else:labels=['受付機能','照会先','確認窓口','発行窓口','保管先'][:n]
            nodes=[{'id':f'actor_{i}','label':text(fill(entry['relation']['nodes'][i]['label'],label))} for i,label in enumerate(labels)]
            edges=[];groups=[]
            if kind=='sequence':
                pairs=[(0,1),(1,2),(2,1),(1,1)] if entry['max_events']==4 else [(0,1),(1,2),(2,1),(1,1),(1,n-1),(n-1,1),(1,0),(0,0)]
                # Five-lane fixtures exercise every actor and a long crossing span.
                if n==5:pairs[-2:]=[(3,4),(4,3)]
                if n==4 and entry['max_events']==4:pairs[-1]=(1,3)
                labels_e=['参照先を渡す','文書を要求','文書を応答','内部で確認','受付を依頼','結果を応答','資料を提示','内部で保存']
                if n==5:
                    labels_e[len(pairs)-2:len(pairs)]=['別窓口へ照会','結果を返す']
                if n==4 and entry['max_events']==4:labels_e[3]='処理を引き継ぐ'
            else:
                pairs=[tuple(map(int,key.split('>'))) for key in entry['relation']['routes']]
                if kind=='fan' and n==3 and not stress:pairs=[(1,0),(2,0)]
                labels_e=['同じURLを提示']*len(pairs) if kind=='fan' and n==3 and not stress else ['照会' if i%2==0 else '応答' for i in range(len(pairs))]
            if usage:
                if kind=='pair':pairs=[(0,1)];labels_e=['照会を開始']
                elif kind=='spoke':pairs=[(1,0),(0,2),(3,0)];labels_e=['操作','照会','付与']
                elif kind=='pair_fan':pairs=[(0,1),(2,3),(2,4)];labels_e=['照会','配信','配信']
            for i,((a,z),label) in enumerate(zip(pairs,labels_e)):
                if stress:
                    if kind=='sequence':
                        from types import SimpleNamespace
                        mock=SimpleNamespace(nodes=[SimpleNamespace(id=f'actor_{j}') for j in range(n)],edges=[SimpleNamespace(source=f'actor_{x}',target=f'actor_{y}') for x,y in pairs])
                        bounds,_=sequence_geometry(entry,mock,i)
                        cap=bounds_capacity(bounds,16,2 if entry['max_events']==4 else 1)
                    else:cap=entry['relation']['routes'][f'{a}>{z}']['label']
                    label=fill(cap,label)
                edges.append({'id':f'edge_{i}','source':f'actor_{a}','target':f'actor_{z}','label':text(label, None if stress else f'{labels[a]}から{labels[z]}へ「{label}」を伝える。')})
            if kind in ('spoke','sequence') and not (usage and kind=='sequence') or kind=='fan' and n>3:
                members=['actor_0'] if kind!='sequence' else ['actor_0','actor_1']
                label='受付アプリ' if kind!='sequence' else '共通運用'
                if stress:
                    from types import SimpleNamespace
                    from slide_agent.relations import boundary_geometry
                    from slide_agent.validate import units
                    mock=SimpleNamespace(nodes=[SimpleNamespace(id=f'actor_{j}') for j in range(n)],boundaries=[SimpleNamespace(members=members)])
                    _,bounds=boundary_geometry(entry,mock)
                    prefix='内包：' if kind!='sequence' else '同管理：'
                    label='調'*math.floor(bounds_capacity(bounds,14,1)['max_units_per_line']-units(prefix))
                groups=[{'id':'owner','kind':'contains' if kind!='sequence' else 'same_owner','label':text(label,f'{label}は対象の機能を含む管理範囲を表す。'),'members':members}]
            contents['network']={'nodes':nodes,'edges':edges,'boundaries':groups}
        slides.append({'id':f'slide{page:02}','origin':group,'layout_id':entry['layout_id'],'title':text(title),'rationale':'架空の構造例。矢印の向き・主体ID・包含と同管理・比較軸を保持する。',
                       'repetition_reason':'全型の構造QAのため、登録済み各layoutを一度ずつ描画する。','contents':contents})
    source=Source.model_validate({'title':'架空の通信・関係図検証','segments':segments})
    display={'mode':'qa' if qa else 'reader','citations':[{'source_ids':[s['id'] for s in segments],'title':'架空設計書','url':'https://example.com/fictional-design','version':'v1','accessed_on':'2026-10-07'}]}
    plan=Plan.model_validate({'title':'架空の構造検証','source_sha256':fingerprint(source),'slides':slides,'display':display})
    return source,plan


def main():
    p=argparse.ArgumentParser();p.add_argument('--out',type=Path,default=ROOT/'examples/relations');p.add_argument('--stress',action='store_true');p.add_argument('--usage',action='store_true');args=p.parse_args()
    entries=list(relation_registry().values())
    if args.usage:entries=[relation_registry()[k] for k in ('rel_pair_2_lr','rel_spoke_4_lr','rel_sequence_4_8_lr','rel_pair_fan_5_lr','rel_fan_3_lr','rel_compare_three')]
    args.out.mkdir(parents=True,exist_ok=True);source,plan=fixture(entries,args.stress,usage=args.usage)
    report=validate(plan,source,args.out,Theme())
    if not report['ok']:raise ValueError([i for i in report['issues'] if i['severity']=='error'])
    for name,value in [('source',source),('plan',plan)]: (args.out/f'relations.{name}.json').write_text(value.model_dump_json(indent=2),encoding='utf-8')
    deck=args.out/'relations.pptx';render(plan,source,args.out,Theme(),deck)
    report['ooxml']=audit(deck,Theme())
    if not report['ooxml']['ok']:raise ValueError(report['ooxml'])
    (args.out/'relations.validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(f'{len(plan.slides)} native structure fixtures: validation and OOXML audit passed')


if __name__=='__main__':main()
