"""Reproduce all native sample decks using installed Python dependencies only."""
import argparse
import json
import re
import sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT))
from slide_agent.catalog import registry
from slide_agent.io import fingerprint
from slide_agent.models import Source, Plan, Theme
from slide_agent.render import render
from slide_agent.validate import validate
from slide_agent.audit import audit
from slide_agent.numeric_format import exact_number


def example_text(text):
    text=re.sub(r"Text\s*(\d+)",r"記載例\1",text)
    text=re.sub(r"ラベル\s*(\d+)",r"項目\1",text)
    text=text.replace("YYYY","2026").replace("Source 1","公開見本").replace("タイトル", "見出し")
    return text


def sample_plan(entries, variant="warm"):
    segments,slides=[],[]
    for index,entry in enumerate(entries,1):
        sid=f"s{index:03}"
        source_text=[]
        def text(value):
            value=example_text(value)
            source_text.append(value)
            return {"text":value,"refs":[{"source_id":sid,"quote":value}],"mode":"verbatim"}
        def datum(value):
            value=float(value); raw=exact_number(value)
            source_text.append(raw)
            return {"value":value,"refs":[{"source_id":sid,"quote":raw}]}
        title=text(entry["title"]["sample"])
        contents={"texts":{k:text(v["sample"]) for k,v in entry["texts"].items()},"charts":{},"metrics":{},"states":{k:text(v["sample"]) for k,v in entry["states"].items()}}
        for key,spec in entry["charts"].items():
            sample=spec["sample"]
            contents["charts"][key]={"categories":[text(x) for x in sample["categories"]],
                                     "series":[{"name":text(s["name"]),"values":[datum(v) for v in s["values"]],
                                                "x_values":[datum(x) for x in s["x_values"]]} for s in sample["series"]]}
            if "elapsed_years" in sample: contents["charts"][key]["elapsed_years"]=datum(sample["elapsed_years"])
        for key,spec in entry["metrics"].items(): contents["metrics"][key]=datum(spec["sample"])
        segments.append({"id":sid,"group":"catalog","text":"\n".join(source_text),
                         "citation":f"Public MIT sample: {entry['source_file']} section {entry['source_section']}"})
        slides.append({"id":f"slide{index:03}","origin":"catalog","layout_id":entry["layout_id"],"variant":variant,
                       "title":title,"rationale":"Public template fixture; short fictional slot content; no private data.","contents":contents})
    source=Source.model_validate({"title":"公開レイアウト見本","segments":segments})
    plan=Plan.model_validate({"title":"公開レイアウト見本","source_sha256":fingerprint(source),"slides":slides})
    return source,plan


def main():
    parser=argparse.ArgumentParser(); parser.add_argument("--out",type=Path,default=ROOT/"catalog/samples")
    parser.add_argument("--individual",action="store_true"); args=parser.parse_args()
    args.out.mkdir(parents=True,exist_ok=True)
    entries=list(registry().values())
    for variant in ["warm","cool"]:
        source,plan=sample_plan(entries,variant)
        report=validate(plan,source,args.out,Theme())
        errors=[i for i in report["issues"] if i["severity"]=="error"]
        if errors:
            print(json.dumps(errors,ensure_ascii=False,indent=2)); raise SystemExit(2)
        for name,obj in [("source",source),("plan",plan)]:
            (args.out/f"catalog-{variant}.{name}.json").write_text(obj.model_dump_json(indent=2),encoding="utf-8")
        dest=args.out/f"catalog-{variant}.pptx"
        render(plan,source,args.out,Theme(),dest)
        report["ooxml"]=audit(dest,Theme())
        (args.out/f"catalog-{variant}.validation.json").write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding="utf-8")
        if not report["ooxml"]["ok"]:
            print(json.dumps(report["ooxml"],ensure_ascii=False,indent=2)); raise SystemExit(2)
        if args.individual and variant=="warm":
            for entry in entries:
                one_source,one_plan=sample_plan([entry],variant)
                render(one_plan,one_source,args.out,Theme(),args.out/f"{entry['layout_id']}.pptx")
        print(f"{variant}: {len(plan.slides)} template slides generated and structurally audited without Office")


if __name__=="__main__": main()
