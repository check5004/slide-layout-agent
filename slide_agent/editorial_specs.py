"""Original research-sharing designs. Geometry variants, never color variants.

This small definition list is available before the generated manifest exists.
The build script is the only producer of template geometry.
"""
FAMILIES = {
    'cover': ('調査の入口', [('brief', 2), ('split', 2), ('band', 2)]),
    'takeaway': ('結論と根拠', [('left', 3), ('right', 3), ('top', 3), ('rail', 3)]),
    'bullets': ('論点と説明', [(v, n) for n in range(2, 7) for v in ('rows', 'columns')]),
    'cards': ('調査所見カード', [(v, n) for n in range(2, 7) for v in ('grid', 'rows')] +
              [(v, n) for n in (3, 4) for v in ('feature_left', 'feature_right')] +
              [(v, 2) for v in ('image_left', 'image_right')]),
    'diagram': ('仕組みと説明', [(v, 3) for v in ('left', 'right', 'wide_left', 'wide_right')]),
    'comparison': ('二案の比較', [(v, 2) for v in ('columns', 'swapped', 'rows', 'matrix')]),
    'steps': ('順序のある手順', [(v, n) for n in range(2, 7) for v in
                              (('horizontal', 'vertical') if n <= 4 else ('vertical',))]),
    'timeline': ('時間軸', [(v, n) for n in (3, 4) for v in ('horizontal', 'vertical')]),
    'metrics': ('数値と留保', [('row', n) for n in (2, 3, 4)] + [('stack', 3), ('breakdown_top', 3), ('breakdown_left', 3)]),
    'table': ('比較表', [('two_columns', n) for n in (3, 4, 6)] + [('three_columns', 4)]),
    'chart': ('グラフと所見', [(v, 4) for v in ('bar_left', 'bar_right', 'column_top', 'column_bottom')]),
    'quote': ('引用と解釈', [(v, 2) for v in ('left', 'right', 'top')]),
    'summary': ('まとめと次の確認', [(v, 3) for v in ('columns', 'rows', 'split')]),
    'agenda': ('調査範囲', [(v, 3) for v in ('rows', 'columns', 'rail')]),
}

SPECS = tuple(dict(layout_id=f'ed_{family}_{variant}_{n}', family=family,
                   structural_variant=variant, item_count=n, name=name)
              for family, (name, variants) in FAMILIES.items() for variant, n in variants)
LAYOUT_IDS = tuple(s['layout_id'] for s in SPECS)
