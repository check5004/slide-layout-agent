"""Saved AI schemas must agree with the runtime models and trusted catalog contract.

The runtime uses Pydantic and semantic preflight, not a third-party JSON-schema
interpreter. This comparison detects stale/edited AI-facing files.
"""
import copy
import hashlib
import json
from functools import lru_cache

from .catalog import ROOT,all_registry as registry
from .models import CatalogSlide,EditorialSlide,Plan


@lru_cache(maxsize=1)
def _catalog_model_schema():return CatalogSlide.model_json_schema()


@lru_cache(maxsize=1)
def _editorial_model_schema():return EditorialSlide.model_json_schema()


@lru_cache(maxsize=1)
def _plan_model_schema():return Plan.model_json_schema()


def layout_schema(entry):
    editorial = entry.get('catalog') == 'editorial'
    sc=copy.deepcopy(_editorial_model_schema() if editorial else _catalog_model_schema())
    sc['properties']['layout_id']={'const':entry['layout_id'],'type':'string'}
    contents={'type':'object','additionalProperties':False,'required':['texts','charts','metrics','states'],'properties':{}}
    kinds = [('texts','Text'),('charts','TemplateChart'),('metrics','SignedDatum'),('states','Text')]
    if editorial:
        kinds.append(('images', 'TemplateImage')); contents['required'].append('images')
    for kind,model in kinds:
        contents['properties'][kind]={'type':'object','additionalProperties':False,'required':list(entry[kind]),
            'properties':{key:{'$ref':f'#/$defs/{model}'} for key in entry[kind]}}
    sc['properties']['contents']=contents
    chart=sc['$defs']['TemplateChart']
    if entry['layout_id']=='scenario_lines_cagr':
        chart['required']=sorted(set(chart.get('required',[])+['elapsed_years']))
        chart['properties']['elapsed_years']={'$ref':'#/$defs/ElapsedYears'}
    else:chart['properties']['elapsed_years']={'type':'null','default':None}
    keys=['template_sha256','title','texts','charts','metrics','states','chart_aliases','reference_line','bridge','calculation','split_policy','image_policy', 'family', 'structural_variant', 'item_count', 'semantic_signature', 'slot_mapping', 'visual_signature', 'images']
    contract={k:entry[k] for k in keys+['emphasis_item', 'text_flow', 'metric_total', 'comparison_axes'] if k in entry}
    sc['x-contract-sha256']=hashlib.sha256(json.dumps(contract,sort_keys=True,ensure_ascii=False,separators=(',',':')).encode()).hexdigest()
    sc['x-runtime-validation']='Pydantic plus semantic preflight; full-width capacities and relations are recorded in the trusted catalog'
    return sc


def schema_error(layout_id):
    entry = registry()[layout_id]
    expected={'catalog/plan.schema.json':_plan_model_schema(),entry.get('schema',f'catalog/schemas/{layout_id}.schema.json'):layout_schema(entry)}
    for path,contract in expected.items():
        try: actual=json.loads((ROOT/path).read_text(encoding='utf-8-sig'))
        except (OSError,ValueError):return f'{path}: missing or invalid saved schema; regenerate and review the catalog'
        if actual!=contract:return f'{path}: saved AI schema differs from runtime/catalog contract; regenerate and review'
    return None
