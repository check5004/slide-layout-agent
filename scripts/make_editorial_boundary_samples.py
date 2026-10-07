"""Exercise every Japanese text slot at its declared line/width boundary."""
import copy
import math
import json
import sys
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1];sys.path.insert(0,str(ROOT))
from slide_agent.catalog import editorial_registry
from scripts.make_editorial_samples import save_samples


def boundary_entries():
    entries=copy.deepcopy(list(editorial_registry().values()))
    for entry in entries:
        for spec in [entry['title'],*entry['texts'].values()]:
            spec['sample']='\n'.join(['調'*math.floor(spec['max_units_per_line'])]*spec['max_lines'])
    return entries


def main():
    entries=boundary_entries()
    dest=ROOT/'out/editorial-boundary/boundary.pptx'
    save_samples(entries,dest,boundary_fixture=True)
    plan=json.loads(dest.with_suffix('.plan.json').read_text(encoding='utf-8'))
    coverage=[]
    for slide,entry in zip(plan['slides'],entries):
        for key,spec in [('title',entry['title']),*entry['texts'].items()]:
            actual=slide['title']['text'] if key=='title' else slide['contents']['texts'][key]['text']
            if actual!=spec['sample']:raise ValueError(f'Boundary fixture changed: {entry["layout_id"]}/{key}')
            coverage.append({'layout_id':entry['layout_id'],'slot':key,'lines':spec['max_lines'],
                             'full_width_characters_per_line':math.floor(spec['max_units_per_line'])})
    dest.with_suffix('.coverage.json').write_text(json.dumps({'slot_count':len(coverage),'slots':coverage},indent=2),encoding='utf-8')


if __name__=='__main__':main()
