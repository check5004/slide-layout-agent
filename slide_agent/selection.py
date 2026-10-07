"""Deterministic content-first variation within an already chosen semantic family.

Codex selects families and writes source-backed content. This module evaluates
only registered equivalents, never rewrites text or infers a story.
"""
from collections import Counter

from .catalog import editorial_registry
from .catalog_validate import validate_catalog


def candidates(slide):
    entries = editorial_registry()
    if slide.layout_id not in entries:
        return [], []
    current = entries[slide.layout_id]
    eligible, rejected = [], []
    for entry in entries.values():
        if entry['family'] != current['family']:
            continue
        reasons = []
        if entry.get('emphasis_item') != (slide.emphasis.item_id if slide.emphasis else None):
            reasons.append('semantic hierarchy differs: featured item requires explicit source-backed emphasis')
        if entry['semantic_signature'] != current['semantic_signature']:
            reasons.append('semantic slots or item count differ')
        if slide.allowed_layouts and entry['layout_id'] not in slide.allowed_layouts:
            reasons.append('excluded by explicit reading/content constraint')
        if not reasons:
            trial = slide.model_copy(update={'layout_id': entry['layout_id']})
            validate_catalog(trial, lambda c, m, s: reasons.append(f'{c}: {m}'))
        if reasons:
            rejected.append({'layout_id': entry['layout_id'], 'reasons': reasons})
        else:
            eligible.append(entry)
    return eligible, rejected


def selection_report(plan, add=None):
    entries = editorial_registry()
    report, signatures = [], []
    for slide in plan.slides:
        if slide.layout_id not in entries:
            signatures.append(slide.layout_id)
            continue
        entry = entries[slide.layout_id]
        fitting, rejected = candidates(slide)
        fitting_ids = [e['layout_id'] for e in fitting]
        allowed = slide.allowed_layouts
        if allowed:
            invalid = [key for key in allowed if key not in entries or entries[key]['family'] != entry['family']
                       or entries[key]['semantic_signature'] != entry['semantic_signature']]
            if invalid and add: add('VARIANT_CONSTRAINT', f'allowed_layouts contain non-equivalent layouts: {invalid}', slide.id)
            if slide.layout_id not in allowed and add: add('VARIANT_CONSTRAINT', 'selected layout is outside allowed_layouts', slide.id)
        signature = entry['visual_signature']
        repeated = len(signatures) >= 3 and signatures[-3:] == [signature] * 3
        alternatives = [e['layout_id'] for e in fitting if e['visual_signature'] != signature]
        if repeated and add:
            if not slide.repetition_reason:
                add('VARIANT_REPETITION', 'four consecutive slides share a visual structure; run select-variants or record a specific repetition_reason', slide.id)
            else:
                add('VARIANT_REPETITION_ACCEPTED', slide.repetition_reason, slide.id, 'warning')
        report.append({'slide_id': slide.id, 'family': entry['family'], 'item_count': entry['item_count'],
                       'layout_id': slide.layout_id, 'structural_variant': entry['structural_variant'],
                       'semantic_slot_mapping': entry['slot_mapping'], 'visual_signature': signature,
                       'family_reason': slide.rationale, 'content_fit_candidates': fitting_ids,
                       'emphasis': slide.emphasis.model_dump() if slide.emphasis else None,
                       'row_alignments': {key:slide.row_alignments.get(key,row['default_alignment']) for key,row in entry.get('rows',{}).items()},
                       'rejected_candidates': rejected, 'constraints_reason': slide.constraints_reason,
                       'fourth_consecutive_structure': repeated, 'fitting_other_structures': alternatives,
                       'repetition_reason': slide.repetition_reason,
                       'selection_reason': 'Equivalent semantic slots and per-slot capacities checked; source content and order retained.'})
        signatures.append(signature)
    return report


def select_variants(plan):
    """Return a new plan and transparent decisions. Never mutate the input."""
    result = plan.model_copy(deep=True)
    recent, used = [], Counter()
    decisions = []
    entries = editorial_registry()
    for slide in result.slides:
        if slide.layout_id not in entries:
            recent.append(slide.layout_id)
            continue
        fitting, rejected = candidates(slide)
        if not fitting:
            raise ValueError(f'{slide.id}: no equivalent variant fits; split/replan with retained evidence, never shrink')
        old = slide.layout_id
        # Actual content fit is a hard filter. Recency is only a tie-breaker.
        # Prefer a different structure from the preceding page, then rare use.
        chosen = min(fitting, key=lambda e: (
            sum(s == e['visual_signature'] for s in recent[-3:]),
            used[e['layout_id']], e['layout_id'] != old, e['layout_id']))
        sig = chosen['visual_signature']
        unavoidable = len(recent) >= 3 and recent[-3:] == [sig] * 3
        slide.layout_id = chosen['layout_id']
        if unavoidable and not slide.repetition_reason:
            slide.repetition_reason = ('Only this visual structure satisfies the registered semantic slots, '
                                       'capacities and explicit constraints. ' + (slide.constraints_reason or
                                       'Other structures were rejected by the per-slot fit checks.'))
        recent.append(sig); used[chosen['layout_id']] += 1
        decisions.append({'slide_id': slide.id, 'from': old, 'to': slide.layout_id,
                          'reason': 'Content-fit filter, then least recent structure and lowest usage; deterministic tie-break.',
                          'fitting': [e['layout_id'] for e in fitting], 'rejected': rejected})
    return result, decisions
