"""Publish SHA-matched real renders, never an HTML approximation of a slide."""
import argparse
import hashlib
import html
import json
import shutil
import sys
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from scripts.publish_editorial_assets import contact
from slide_agent.catalog import relation_registry


def main():
    p=argparse.ArgumentParser();p.add_argument('--pptx',type=Path,required=True);p.add_argument('--plan',type=Path,required=True);p.add_argument('--rendered',type=Path,required=True);p.add_argument('--out',type=Path,required=True);a=p.parse_args()
    report=json.loads((a.rendered/'render-report.json').read_text(encoding='utf-8-sig'));plan=json.loads(a.plan.read_text(encoding='utf-8'))
    if report['pptxSha256']!=hashlib.sha256(a.pptx.read_bytes()).hexdigest() or report['slideCount']!=len(plan['slides']) or report['overflowCount'] or report['positionOverflowCount']:raise ValueError('render SHA/page count/overflow verification failed')
    a.out.mkdir(parents=True,exist_ok=True);(a.out/'previews').mkdir(exist_ok=True)
    cards=[];images=[];labels=[];coverage=[]
    for i,slide in enumerate(plan['slides'],1):
        e=relation_registry()[slide['layout_id']]
        name=f'{e["layout_id"]}.png' if a.out.resolve()==(ROOT/'catalog/relations').resolve() else f'{i:02}-{slide["layout_id"]}.png'
        target=a.out/'previews'/name
        shutil.copyfile(a.rendered/f'slide-{i:02}.png',target)
        images.append(target);labels.append(f'{i:02} {slide["layout_id"]}')
        coverage.append({'page':i,'layout_id':e['layout_id'],'template_sha256':e['template_sha256'],'png_sha256':hashlib.sha256(target.read_bytes()).hexdigest(),'preview':f'previews/{name}'})
        cards.append(f'<article><h2>{html.escape(e["layout_id"])}</h2><a href="previews/{name}"><img src="previews/{name}" alt="{html.escape(slide["title"]["text"])}"></a><details><summary>構造・容量・native slot</summary><pre>{html.escape(json.dumps(e["relation"],ensure_ascii=False,indent=2))}</pre></details></article>')
    contact(images,labels,a.out/'contact-sheet.png',columns=2)
    for start in range(0,len(images),6):contact(images[start:start+6],labels[start:start+6],a.out/f'contact-{start//6+1:02}.png',columns=2)
    (a.out/'index.html').write_text('<!doctype html><html lang="ja"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Native relationship layouts</title><style>body{margin:0;background:#eef2f5;color:#213547;font:16px/1.6 system-ui}header{padding:28px 4vw;background:#213547;color:white}main{display:grid;grid-template-columns:repeat(auto-fit,minmax(480px,1fr));gap:24px;padding:24px}article{background:white;padding:16px;border-radius:8px}h1{margin:0}h2{font-size:18px}img{width:100%;display:block}pre{overflow:auto;max-height:400px;font-size:13px}a{color:#087e83}@media(max-width:600px){main{grid-template-columns:1fr}}</style><header><h1>関係図・通信図 / native editable</h1><p>架空の構造例。保存済みPPTXからPowerPointで実描画。線・矢印・文字・表は編集できます。HTMLはPNGの一覧です。</p></header><main>'+''.join(cards)+'</main></html>',encoding='utf-8')
    (a.out/'coverage.json').write_text(json.dumps({'pptx_sha256':report['pptxSha256'],'pages':coverage},indent=2),encoding='utf-8')
    shutil.copyfile(a.rendered/'render-report.json',a.out/'render-report.json')
    print(f'Published {len(images)} verified real-render previews')


if __name__=='__main__':main()
