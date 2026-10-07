"""Trusted native row anchors and bounded body/evidence groups; no font fitting."""
import math

from pptx.enum.text import MSO_ANCHOR
from pptx.util import Pt

ANCHORS = {'top':MSO_ANCHOR.TOP, 'middle':MSO_ANCHOR.MIDDLE}


def anchor_target(shapes, member):
    shape = shapes[member['shape']]
    return shape.table.cell(*member['cell']) if 'cell' in member else shape.text_frame


def row_anchors(entry, overrides):
    """Expected native anchor for each registered row member."""
    flow_shapes = {entry['texts'][flow[key]]['shape']
                   for flow in entry.get('text_flow', []) for key in ('body','following')}
    for key, row in entry.get('rows', {}).items():
        alignment = overrides.get(key, row['default_alignment'])
        for member in row['members']:
            yield member, 'top' if member['shape'] in flow_shapes else alignment


def flow_positions(entry, contents, overrides):
    from .validate import units
    result = []
    for flow in entry.get('text_flow', []):
        spec = entry['texts'][flow['body']]
        following = entry['texts'][flow['following']]
        body_bounds, tail_bounds = spec['bounds_emu'], following['bounds_emu']
        text = contents.texts[flow['body']].text
        lines = sum(max(1, math.ceil(units(line)/spec['max_units_per_line'])) for line in text.split('\n'))
        height = min(body_bounds[3], Pt(spec['font_pt']*(1.33+1.25*(lines-1))+2.2))
        top = body_bounds[1]
        if flow.get('row'):
            row = entry['rows'][flow['row']]
            block_height = height + flow['gap_emu'] + tail_bounds[3]
            if block_height > row['height_emu']:
                raise ValueError('body plus evidence exceeds its registered row; split/replan, never shrink')
            free = row['height_emu']-block_height
            top = row['top_emu'] + (free//2 if overrides.get(flow['row'],row['default_alignment'])=='middle' else 0)
            tail_top = top+height+flow['gap_emu']
        else:
            tail_top = min(tail_bounds[1], top+height+flow['gap_emu'])
        result.append({'body':spec['shape'],'following':following['shape'],
                       'top':top,'height':height,'following_top':tail_top})
    return result


def apply_vertical(shapes, entry, contents, overrides):
    for member, alignment in row_anchors(entry, overrides):
        anchor_target(shapes, member).vertical_anchor = ANCHORS[alignment]
    for flow in flow_positions(entry, contents, overrides):
        body, following = shapes[flow['body']], shapes[flow['following']]
        body.top, body.height, following.top = flow['top'], flow['height'], flow['following_top']
