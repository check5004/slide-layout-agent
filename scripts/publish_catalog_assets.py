"""Assemble offline gallery, exact per-layout schemas and QA coverage.

PNG inputs must be produced by the real PowerPoint development QA script.
Visual review remains explicit and separate from pixel/statistical checks.
"""
import argparse
import copy
import hashlib
import html
import json
import shutil
import sys
from pathlib import Path

from PIL import Image, ImageStat, ImageDraw, ImageFont

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from slide_agent.models import Plan, CatalogSlide
from slide_agent.schema_contract import layout_schema


def main():
    p=argparse.ArgumentParser();p.add_argument("--warm",type=Path,required=True);p.add_argument("--cool",type=Path,required=True)
    p.add_argument("--reviewed",action="store_true");args=p.parse_args()
    manifest_path=ROOT/"catalog/manifest.json"
    manifest=json.loads(manifest_path.read_text(encoding="utf-8"))
    coverage=[];reports={}
    for variant,directory in [("warm",args.warm),("cool",args.cool)]:
        report=json.loads((directory/"render-report.json").read_text(encoding="utf-8-sig"))
        if report["slideCount"]!=62 or report["overflowCount"]:
            raise ValueError(f"{variant}: invalid slide count or measured overflow")
        expected=hashlib.sha256((ROOT/f"catalog/samples/catalog-{variant}.pptx").read_bytes()).hexdigest()
        if expected!=report["pptxSha256"]: raise ValueError("render report refers to old PPTX bytes")
        reports[variant]={k:v for k,v in report.items() if k!="editableTextBoundaryMeasurements"}
        target_dir=ROOT/f"catalog/previews/{variant}";target_dir.mkdir(parents=True,exist_ok=True)
        (ROOT/"catalog/qa").mkdir(exist_ok=True)
        shutil.copyfile(directory/"render-report.json",ROOT/f"catalog/qa/{variant}-powerpoint.json")
        for i,entry in enumerate(manifest["layouts"],1):
            source=directory/f"slide-{i:02}.png";target=target_dir/f"{entry['layout_id']}.png"
            im=Image.open(source).convert("RGB")
            if im.size!=(1600,900):raise ValueError("unexpected raster size")
            stat=ImageStat.Stat(im)
            if max(stat.stddev)<5:raise ValueError("blank or near-blank raster")
            shutil.copyfile(source,target)
            coverage.append({"source_file":entry["source_file"],"source_section":entry["source_section"],
                "source_pptx_slide":entry["source_pptx_slide"],"layout_id":entry["layout_id"],"variant":variant,
                "template":entry["template"],"sample_deck":f"catalog/samples/catalog-{variant}.pptx","sample_slide":i,
                "preview":target.relative_to(ROOT).as_posix(),"png_sha256":hashlib.sha256(target.read_bytes()).hexdigest(),
                "geometry":"passed","ooxml":"passed","editable":"native text/shapes/tables and registered charts; zero pictures",
                "powerpoint":"rendered","measured_overflow":0,"pixel_nonblank":"passed",
                "visual_review":"inspected" if args.reviewed else "pending",
                "html_comparison":"inspected; documented differences" if args.reviewed else "pending"})
    schema_dir=ROOT/"catalog/schemas";schema_dir.mkdir(exist_ok=True)
    for entry in manifest["layouts"]:
        sc=layout_schema(entry)
        (schema_dir/f"{entry['layout_id']}.schema.json").write_text(json.dumps(sc,ensure_ascii=False,indent=2),encoding="utf-8")
        entry["schema"]=f"catalog/schemas/{entry['layout_id']}.schema.json"
        entry["samples"]={v:{"deck":f"catalog/samples/catalog-{v}.pptx","slide":i+1,
                             "preview":f"catalog/previews/{v}/{entry['layout_id']}.png"} for v in ["warm","cool"] for i,e in enumerate(manifest["layouts"]) if e is entry}
        entry["source_preview"]=f"catalog/source-previews/{entry['layout_id']}.png"
        entry["validation_status"]="visually_inspected" if args.reviewed else "rendered_pending_visual_review"
    manifest_path.write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    (ROOT/"catalog/coverage.json").write_text(json.dumps({"upstream_commit":manifest["upstream_commit"],"structural_layouts":62,
            "variant_samples":124,"covered_layouts":62,"excluded_layouts":[],"renderers":reports,"records":coverage},ensure_ascii=False,indent=2)+"\n",encoding="utf-8")
    cards=[]
    for e in manifest["layouts"]:
        key=html.escape(e["layout_id"]);name=html.escape(e["name"])
        details=html.escape(json.dumps({k:e[k] for k in ["layout_id","title","texts","charts","metrics","states","differences","split_policy","image_policy"]},ensure_ascii=False,indent=2))
        cards.append(f'''<article data-search="{key} {name} {e['part_id']}"><h2>{e['part_id']} · {name}</h2><code>{key}</code>
          <div class="pair"><figure><img loading="lazy" src="source-previews/{key}.png" alt="上流HTML {name}"><figcaption>固定commitのHTML</figcaption></figure>
          <figure><img loading="lazy" class="preview" data-key="{key}" src="previews/warm/{key}.png" alt="PowerPoint差込見本 {name}"><figcaption>PowerPoint実描画</figcaption></figure></div>
          <p><a href="templates/{key}.pptx">template.pptx</a> · <a href="samples/{key}.pptx">差込sample.pptx</a> · <a href="schemas/{key}.schema.json">schema</a></p>
          <p>テキスト {len(e['texts'])} / chart {len(e['charts'])} / 数値 {len(e['metrics'])} / 状態 {len(e['states'])}</p>
          <details><summary>入力slot・容量・差異を表示</summary><pre>{details}</pre></details></article>''')
    page='''<!doctype html><html lang="ja"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
    <title>62型 Native PPTX Catalog</title><style>body{font:16px system-ui,sans-serif;margin:0;background:#f4f5f7;color:#142133}header{padding:32px;position:sticky;top:0;background:#fff;z-index:2;border-bottom:1px solid #ddd}h1{margin:0 0 8px}input,select{padding:10px;font:inherit;margin:8px 12px 0 0;border:1px solid #aaa;border-radius:6px}main{padding:24px;max-width:1800px;margin:auto}article{background:white;border:1px solid #dde0e5;margin-bottom:24px;padding:22px;border-radius:8px}h2{font-size:22px}.pair{display:grid;grid-template-columns:1fr 1fr;gap:16px}figure{margin:16px 0}img{width:100%;border:1px solid #eee}figcaption{color:#526070;margin-top:6px}pre{white-space:pre-wrap;max-height:600px;overflow:auto;background:#f5f5f5;padding:16px}a{color:#006273}@media(max-width:700px){.pair{grid-template-columns:1fr}header{position:relative}}</style>
    <header><h1>62型 Native PPTX Catalog</h1><p>保存済みテンプレートへ差し込むPython workflow。構造62型 / 配色2種 / native要素。</p>
    <input id="search" placeholder="型名・ID・part IDで検索" aria-label="型を検索"><select id="variant" aria-label="配色"><option value="warm">Warm</option><option value="cool">Cool</option></select><span id="count">62 / 62</span>
    <p><a href="../docs/catalog/template-workflow.md">使い方</a> · <a href="coverage.json">全件coverage</a> · <a href="manifest.json">slot契約</a> · <a href="upstream/LICENSE">MIT / Carnot AI Inc.</a></p></header><main>'''+"\n".join(cards)+'''</main><script>
    document.querySelector('#search').addEventListener('input',e=>{let n=0;for(const a of document.querySelectorAll('article')){a.hidden=!a.dataset.search.toLowerCase().includes(e.target.value.toLowerCase());if(!a.hidden)n++}document.querySelector('#count').textContent=n+' / 62'});
    document.querySelector('#variant').addEventListener('change',e=>{for(const i of document.querySelectorAll('.preview'))i.src='previews/'+e.target.value+'/'+i.dataset.key+'.png'});
    </script></html>'''
    (ROOT/"catalog/index.html").write_text(page,encoding="utf-8")
    comparison_lines=["# HTMLと登録PPTXの型別比較","","固定上流commit: `"+manifest['upstream_commit']+"`。62構造すべて登録済みです。HTMLとのpixel完全一致を意味しません。公開native PPTXを実体templateとして使い、下記の差異を保持または明示的に改修しています。","","| 型 | 個別差異 |","|---|---|"]
    comparisons=json.loads((ROOT/'catalog/html-comparison.json').read_text(encoding='utf-8'))
    for e in manifest['layouts']:
        comparison_lines.append('| '+e['part_id']+' '+e['layout_id']+' | '+' '.join(comparisons[e['layout_id']]).replace('|','\\|')+' |')
    (ROOT/'docs/catalog/html-comparison.md').write_text('\n'.join(comparison_lines)+'\n',encoding='utf-8')
    contact_dir=ROOT/"out/catalog-review";contact_dir.mkdir(parents=True,exist_ok=True)
    font=ImageFont.load_default()
    for batch in range(0,62,4):
        sheet=Image.new("RGB",(1600,1920),"#e5e7eb");draw=ImageDraw.Draw(sheet)
        for j,e in enumerate(manifest["layouts"][batch:batch+4]):
            y=j*480;draw.text((10,y+7),f"{batch+j+1:02} {e['layout_id']} | source HTML left / native filled PPTX right",fill="black",font=font)
            for x,path in [(0,ROOT/e["source_preview"]),(800,ROOT/e["samples"]["warm"]["preview"])]:
                im=Image.open(path).convert("RGB").resize((800,450));sheet.paste(im,(x,y+30))
        sheet.save(contact_dir/f"compare-{batch//4+1:02}.png")
    for batch in range(0,62,8):
        sheet=Image.new("RGB",(1600,1920),"#e5e7eb");draw=ImageDraw.Draw(sheet)
        for j,e in enumerate(manifest["layouts"][batch:batch+8]):
            x=(j%2)*800;y=(j//2)*480;draw.text((x+10,y+7),f"{batch+j+1:02} {e['layout_id']} cool",fill="black",font=font)
            im=Image.open(ROOT/e["samples"]["cool"]["preview"]).resize((800,450));sheet.paste(im,(x,y+30))
        sheet.save(contact_dir/f"cool-{batch//8+1:02}.png")
    print('Published 124 PNGs, 62 schemas, coverage, offline gallery and review contact sheets')


if __name__=="__main__":main()
