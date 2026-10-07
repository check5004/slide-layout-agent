"""Generate all samples and the first visual-review milestone without Office."""
import argparse
import hashlib
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from scripts.make_catalog_samples import sample_plan
from slide_agent.catalog import editorial_registry
from slide_agent.models import Theme
from slide_agent.render import render
from slide_agent.validate import validate
from slide_agent.audit import audit
from slide_agent.models import Source, Plan
from slide_agent.io import fingerprint

REPRESENTATIVE = ['ed_cover_brief_2', 'ed_cards_grid_3', 'ed_cards_grid_6',
                  'ed_cards_rows_4', 'ed_cards_feature_left_3', 'ed_cards_feature_right_3',
                  'ed_diagram_left_3', 'ed_comparison_columns_2', 'ed_steps_horizontal_3',
                  'ed_metrics_breakdown_top_3', 'ed_table_three_columns_4', 'ed_chart_bar_right_4']


def editorial_sample_plan(entries, base=None, *, boundary_fixture=False):
    import copy
    entries = copy.deepcopy(entries)
    for entry in entries:
        if entry.get('emphasis_item') and not boundary_fixture:
            entry['title']['sample'] = '最も重視する論点は検索入口を揃えること'
    source, plan = sample_plan(entries)
    from slide_agent.models import EditorialEmphasis
    for slide, entry in zip(plan.slides, entries):
        slide.repetition_reason = 'Catalog inventory: adjacent entries intentionally show every count/orientation for comparison; this is not a presentation sequence.'
        if entry.get('emphasis_item'):
            slide.emphasis = EditorialEmphasis(item_id='item_1', refs=slide.title.refs,
                                              reason=('Synthetic capacity stress fixture, not a semantic example of priority.' if boundary_fixture else
                                                      'The visible source-backed title explicitly identifies the first item as the main focus.'))
    if any(e.get('images') for e in entries):
        image = ROOT / 'catalog/editorial/assets/research-screen.png'
        if not image.exists(): make_sample_image(image)
        base = base or ROOT / 'catalog/editorial/samples'
        # Sources only permit assets within their directory. Each deliverable
        # receives its own small fictional asset beside source.json.
        import shutil
        base.mkdir(parents=True, exist_ok=True)
        target = base / 'research-screen.png'
        if image.resolve() != target.resolve(): shutil.copyfile(image, target)
        source = source.model_copy(update={'images': []})
        raw = source.model_dump(mode='json')
        raw['images'] = [{'id': 'research_screen', 'path': target.name, 'caption': 'Fictional research search screen; no private information', 'sha256': hashlib.sha256(target.read_bytes()).hexdigest()}]
        source = Source.model_validate(raw)
        data = plan.model_dump(mode='json'); data['source_sha256'] = fingerprint(source)
        for slide, entry in zip(data['slides'], entries):
            slide['contents']['images'] = {key: {'image_id': 'research_screen', 'image_mode': 'fit'} for key in entry.get('images', {})}
        plan = Plan.model_validate(data)
    return source, plan


def make_sample_image(path):
    from PIL import Image, ImageDraw, ImageFont
    image = Image.new('RGB', (1100, 800), '#F0F4F7'); d = ImageDraw.Draw(image)
    # Portable Latin-only artificial UI; fixture is not an actual screenshot.
    font = ImageFont.load_default(size=27); small = ImageFont.load_default(size=21)
    d.rectangle((30, 28, 1070, 105), fill='#213547')
    d.text((62, 51), 'RESEARCH LIBRARY / FICTIONAL UI', font=font, fill='white')
    d.rounded_rectangle((62, 139, 1038, 205), radius=10, fill='white', outline='#C8D4DF', width=2)
    d.text((84, 160), 'Search: knowledge sharing', font=small, fill='#566777')
    for k, title in enumerate(['Search entry points', 'Evidence and context', 'Ownership and updates']):
        y = 243 + k * 163
        d.rectangle((62, y, 1038, y + 139), fill='white')
        d.rectangle((62, y, 69, y + 139), fill='#087E83')
        d.text((90, y + 22), title, font=font, fill='#213547')
        d.text((90, y + 66), 'Scope  /  source  /  review owner', font=small, fill='#566777')
        d.text((90, y + 102), 'Illustrative record; not a real research result', font=small, fill='#566777')
    path.parent.mkdir(parents=True, exist_ok=True); image.save(path)


def save_samples(entries, dest, *, boundary_fixture=False):
    source, plan = editorial_sample_plan(entries, dest.parent, boundary_fixture=boundary_fixture)
    report = validate(plan, source, dest.parent, Theme())
    errors = [i for i in report['issues'] if i['severity'] == 'error']
    if errors:
        for error in errors: print(json.dumps(error, ensure_ascii=True))
        raise ValueError(f'{len(errors)} validation errors')
    dest.parent.mkdir(parents=True, exist_ok=True)
    for label, model in [('source', source), ('plan', plan)]:
        dest.with_suffix(f'.{label}.json').write_text(model.model_dump_json(indent=2), encoding='utf-8')
    render(plan, source, dest.parent, Theme(), dest)
    report['ooxml'] = audit(dest, Theme())
    dest.with_suffix('.validation.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    if not report['ooxml']['ok']: raise ValueError(report['ooxml'])
    print(f'{dest.name}: {len(entries)} slides; preflight and native OOXML audit passed')


def main():
    parser = argparse.ArgumentParser(); parser.add_argument('--representative', action='store_true')
    parser.add_argument('--out', type=Path, default=ROOT / 'catalog/editorial/samples')
    args = parser.parse_args(); entries = editorial_registry()
    chosen = [entries[k] for k in REPRESENTATIVE] if args.representative else list(entries.values())
    save_samples(chosen, args.out / ('representative.pptx' if args.representative else 'editorial.pptx'))


if __name__ == '__main__': main()
