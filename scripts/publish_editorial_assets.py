"""Publish only verified real-engine PNGs as an offline, searchable gallery."""
import argparse
import base64
import hashlib
import html
import json
import math
import shutil
import sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from PIL import Image,ImageDraw,ImageFont
from slide_agent.catalog import editorial_registry


def contact(images, labels, destination, columns=3):
    w,h,gap=640,360,16;top=36
    sheet=Image.new('RGB',(columns*(w+gap)+gap,math.ceil(len(images)/columns)*(h+top+gap)+gap),'#E7ECF0')
    d=ImageDraw.Draw(sheet);font=ImageFont.load_default(size=18)
    for i,(image,label) in enumerate(zip(images,labels)):
        x=gap+(i%columns)*(w+gap);y=gap+(i//columns)*(h+top+gap)
        d.text((x,y+5),label,font=font,fill='#213547')
        with Image.open(image) as im:sheet.paste(im.convert('RGB').resize((w,h),Image.Resampling.LANCZOS),(x,y+top))
    destination.parent.mkdir(parents=True,exist_ok=True);sheet.save(destination)


def publish(pptx, plan_path, rendered, destination, gallery=False):
    report=json.loads((rendered/'render-report.json').read_text(encoding='utf-8-sig'))
    digest=hashlib.sha256(pptx.read_bytes()).hexdigest()
    if report['pptxSha256']!=digest: raise ValueError('rendered PNGs do not match the current PPTX')
    if report['overflowCount']:raise ValueError('real engine measured text overflow; repair before publishing')
    plan=json.loads(plan_path.read_text(encoding='utf-8')); entries=editorial_registry()
    if report['slideCount']!=len(plan['slides']):raise ValueError('render page count mismatch')
    destination.mkdir(parents=True,exist_ok=True)
    images,labels,cards,coverage=[],[],[],[]
    for i,slide in enumerate(plan['slides'],1):
        entry=entries[slide['layout_id']];image=rendered/f'slide-{i:02}.png'
        if not image.is_file():raise ValueError(f'missing real preview {i}')
        filename=f'{slide["layout_id"]}.png' if gallery else f'slide-{i:02}.png'
        target=destination/'previews'/filename;target.parent.mkdir(exist_ok=True);shutil.copyfile(image,target)
        images.append(target);labels.append(f'{i:02} {slide["layout_id"]}')
        coverage.append({'slide':i,'layout_id':slide['layout_id'],'family':entry['family'],'structural_variant':entry['structural_variant'],
                         'template_sha256':entry['template_sha256'],'preview':f'previews/{filename}',
                         'png_sha256':hashlib.sha256(target.read_bytes()).hexdigest()})
        ranges=' / '.join(f'{k}: {s["font_pt"]}pt, {s["max_lines"]} lines × {s["max_units_per_line"]} units' for k,s in entry['texts'].items())
        text=html.escape(slide['title']['text'])
        details=html.escape(json.dumps({'family':entry['family'],'variant':entry['structural_variant'],'count':entry['item_count'],
                                      'visual_signature':entry['visual_signature'],'slots':entry['texts'],'images':entry.get('images',{})},ensure_ascii=False,indent=2))
        src=f'previews/{filename}'
        cards.append(f'<article data-family="{entry["family"]}" data-count="{entry["item_count"]}"><a href="{src}"><img loading="lazy" src="{src}" alt="{text}"></a><div class="body"><small>{i:02} · {entry["family"]} · {entry["structural_variant"]} · {entry["item_count"]} items</small><h2>{text}</h2><code>{slide["layout_id"]}</code><details><summary>Semantic slots / capacity</summary><pre>{details}</pre></details></div></article>')
    contact(images,labels,destination/'contact-sheet.png')
    for start in range(0,len(images),12):contact(images[start:start+12],labels[start:start+12],destination/f'contact-{start//12+1:02}.png')
    options=''.join(f'<option>{f}</option>' for f in sorted({entries[s['layout_id']]['family'] for s in plan['slides']}))
    document='''<!doctype html><html lang="ja"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Editorial research layouts</title><style>
*{box-sizing:border-box}body{margin:0;background:#eef2f5;color:#213547;font:16px/1.6 system-ui,sans-serif}header{padding:36px max(24px,4vw);background:#213547;color:white}h1{font-size:30px;margin:0 0 8px}header p{max-width:850px;margin:0;color:#d9e3ec}nav{position:sticky;top:0;padding:14px 4vw;background:#fff;box-shadow:0 2px 8px #21354718;display:flex;gap:14px;align-items:center;flex-wrap:wrap;z-index:1}label{display:flex;gap:8px;align-items:center}select,input{font:inherit;padding:6px;border:1px solid #b7c7d3;border-radius:5px}main{padding:24px 3vw;display:grid;grid-template-columns:repeat(auto-fit,minmax(430px,1fr));gap:24px}article{background:white;border:1px solid #d7e1e8;border-radius:9px;overflow:hidden}article img{width:100%;display:block}article .body{padding:16px}h2{font-size:19px;margin:5px 0}small{color:#566777}code{font-size:13px;color:#087e83}details{margin-top:12px}pre{overflow:auto;font-size:12px;max-height:400px}article[hidden]{display:none}@media(max-width:600px){main{grid-template-columns:1fr}}a{color:#087e83}</style>
<header><h1>Editorial / 社内調査共有のレイアウト</h1><p>14の意味上のfamily。左右・項目数・図文比率・配置を選べます。本文を縮めず、内容が収まる型を選びます。画像はMicrosoft PowerPointの実描画。すべて架空例です。</p></header>
<nav><label>Family <select id="family"><option value="">すべて</option>OPTIONS</select></label><label>項目数 <select id="count"><option value="">すべて</option><option>2</option><option>3</option><option>4</option><option>5</option><option>6</option></select></label><label>検索 <input id="search" placeholder="cards / 比較 / 所見"></label><span id="total"></span><a href="contact-sheet.png">全型コンタクトシート</a></nav><main>CARDS</main><script>
function filter(){const family=document.querySelector('#family').value,count=document.querySelector('#count').value,q=document.querySelector('#search').value.toLowerCase();let n=0;document.querySelectorAll('article').forEach(a=>{a.hidden=!!((family&&a.dataset.family!==family)||(count&&a.dataset.count!==count)||(q&&!a.innerText.toLowerCase().includes(q)));if(!a.hidden)n++});document.querySelector('#total').textContent=n+' layouts'}document.querySelectorAll('select,input').forEach(e=>e.addEventListener('input',filter));filter();</script></html>'''.replace('OPTIONS',options).replace('CARDS',''.join(cards))
    (destination/'index.html').write_text(document,encoding='utf-8')
    (destination/'coverage.json').write_text(json.dumps({'engine':report['engine'],'version':report['version'],'pptx_sha256':digest,'pages':coverage},ensure_ascii=False,indent=2),encoding='utf-8')
    shutil.copyfile(rendered/'render-report.json',destination/'render-report.json')
    print(f'Published {len(images)} real previews, searchable HTML and contact sheets')


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--pptx',type=Path,required=True);p.add_argument('--plan',type=Path,required=True);p.add_argument('--rendered',type=Path,required=True);p.add_argument('--out',type=Path,required=True);p.add_argument('--gallery',action='store_true');a=p.parse_args()
    publish(a.pptx,a.plan,a.rendered,a.out,a.gallery)
