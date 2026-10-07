"""Finite relationship/communication grammar, separate from the 76 prose layouts."""
SPECS = []
for kind, counts in [('pair', [2]), ('fan', [3, 4, 5]), ('spoke', [4]), ('pair_fan', [5])]:
    for n in counts:
        for direction in ('lr', 'rl'):
            SPECS.append(dict(layout_id=f'rel_{kind}_{n}_{direction}', family='relationship',
                              kind=kind, item_count=n, direction=direction, max_events=0))
for n in (3, 4, 5):
    for maximum in (4, 8):
        for direction in ('lr', 'rl'):
            SPECS.append(dict(layout_id=f'rel_sequence_{n}_{maximum}_{direction}', family='communication',
                              kind='sequence', item_count=n, direction=direction, max_events=maximum))
SPECS.append(dict(layout_id='rel_compare_three', family='method_comparison', kind='compare',
                  item_count=3, direction='lr', max_events=0))
LAYOUT_IDS = tuple(s['layout_id'] for s in SPECS)
