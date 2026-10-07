"""Per-slot native template preflight. Preserves exact schema/cardinality."""
import math
import re
from decimal import DecimalException

from .catalog import all_registry as registry


def validate_catalog(slide, add):
    from .validate import units
    entry = registry()[slide.layout_id]
    c = slide.contents

    def capacity(value, spec, label):
        lines = sum(max(1, math.ceil(units(line)/spec["max_units_per_line"])) for line in value.split("\n"))
        if lines > spec["max_lines"]:
            add("TEMPLATE_OVERFLOW", f"{label}: {lines} lines exceed {spec['max_lines']}; replan/split, no shrink",slide.id)

    capacity(slide.title.text,entry["title"],"title")
    if entry.get('catalog') == 'editorial':
        if not slide.title.text.strip():
            add('TEMPLATE_EMPTY', 'title: whitespace is not content', slide.id)
        actual_emphasis = slide.emphasis.item_id if slide.emphasis else None
        if entry.get('emphasis_item') != actual_emphasis:
            add('EDITORIAL_EMPHASIS', 'this visual hierarchy requires matching source-backed emphasis; parallel items must use equal layouts', slide.id)
        expected_images, supplied_images = entry.get('images', {}), slide.contents.images
        if set(expected_images) != set(supplied_images):
            add('TEMPLATE_IMAGE_SLOTS', 'exact named image slots required; choose a template with the correct image count', slide.id)
    for label, supplied, expected in [("texts",c.texts,entry["texts"]),("charts",c.charts,entry["charts"]),("metrics",c.metrics,entry["metrics"]),("states",c.states,entry["states"])]:
        missing, unknown = set(expected)-set(supplied), set(supplied)-set(expected)
        if missing or unknown:
            add("TEMPLATE_SLOTS",f"{label}: missing={sorted(missing)}, unknown={sorted(unknown)}",slide.id)
    for key, value in c.texts.items():
        if key in entry["texts"]:
            if not value.text.strip(): add('TEMPLATE_EMPTY', f'{key}: whitespace is not content', slide.id)
            capacity(value.text,entry["texts"][key],key)
    for a,b in entry.get('comparison_axes', []):
        if a in c.texts and b in c.texts and c.texts[a].text != c.texts[b].text:
            add('EDITORIAL_COMPARISON_AXES','paired alternatives require the same explicit comparison axis',slide.id)
    for key,value in c.states.items():
        if key in entry["states"] and value.text not in entry["states"][key]["palette"]:
            add("TEMPLATE_STATE",f"{key}: choose {list(entry['states'][key]['palette'])}",slide.id)
    for key, chart in c.charts.items():
        if key not in entry["charts"]: continue
        spec = entry["charts"][key]
        if len(chart.categories)!=spec["category_count"] or len(chart.series)!=spec["series_count"]:
            add("TEMPLATE_CHART_DIMENSIONS",key,slide.id); continue
        for i, series in enumerate(chart.series):
            if len(series.values)!=spec["point_counts"][i] or (spec["kind"]=="xy" and len(series.x_values)!=len(series.values)) or (spec["kind"]!="xy" and series.x_values):
                add("TEMPLATE_CHART_DIMENSIONS",key,slide.id)
        for text in chart.categories+[s.name for s in chart.series]:
            if units(text.text)>spec["max_label_units"] or "\n" in text.text:
                add("TEMPLATE_CHART_LABEL",key,slide.id)
        domain=spec.get("domain")
        if domain:
            vals=[x.value for series in chart.series for x in series.values]
            if any(v<domain[0] or v>domain[1] for v in vals): add("TEMPLATE_CHART_DOMAIN",key,slide.id)
        if spec.get("x_domain"):
            low,high=spec["x_domain"]
            if any(not low<=x.value<=high for series in chart.series for x in series.x_values):
                add("TEMPLATE_CHART_X_DOMAIN",key,slide.id)
        if slide.layout_id=="ranked_bar_annotated":
            values=[x.value for x in chart.series[0].values]
            if values!=sorted(values,reverse=True):add("TEMPLATE_RANK_ORDER","ranked bars must be descending",slide.id)
        if slide.layout_id=="scenario_lines_cagr":
            if chart.elapsed_years is None:
                add("TEMPLATE_CAGR_PERIOD","explicit source-backed elapsed_years is required; never infer years from point count",slide.id)
            else:
                years=[re.fullmatch(r"([1-9]\d{3})(?:年(?:度)?)?",t.text) for t in chart.categories]
                if any(years) and not all(years):
                    add("TEMPLATE_CAGR_CALENDAR","mixed year/non-year labels are ambiguous; use consistent categories",slide.id)
                elif all(years):
                    year_values=[int(m[1]) for m in years]
                    steps=[b-a for a,b in zip(year_values,year_values[1:])]
                    if not steps or min(steps)<=0 or len(set(steps))!=1:
                        add("TEMPLATE_CAGR_CALENDAR","this category-axis template requires increasing, evenly spaced year labels",slide.id)
                    if not math.isclose(year_values[-1]-year_values[0],chart.elapsed_years.value,abs_tol=1e-8):
                        add("TEMPLATE_CAGR_PERIOD","elapsed_years must match the first/last calendar-year difference",slide.id)
            if any(s.values[0].value<=0 or s.values[-1].value<0 for s in chart.series):
                add("TEMPLATE_CAGR_DOMAIN","CAGR requires positive start and nonnegative end",slide.id)
        elif chart.elapsed_years is not None:
            add("TEMPLATE_CHART_PERIOD","elapsed_years is only supported by scenario_lines_cagr",slide.id)
    for key,value in c.metrics.items():
        if key not in entry["metrics"]: continue
        spec=entry["metrics"][key]
        if not spec["min"] <= value.value <= spec["max"] or (spec.get("integer") and not float(value.value).is_integer()):
            add("TEMPLATE_METRIC_RANGE",key,slide.id)
        if spec.get("label_capacity"):
            from .template_engine import metric_text
            capacity(metric_text(value.value,spec),spec["label_capacity"],key)
        if spec["binding"]=="gantt_start" and spec["end_key"] in c.metrics and c.metrics[spec["end_key"]].value<=value.value:
            add("TEMPLATE_GANTT_ORDER",key,slide.id)
    relation = entry.get('metric_total')
    if relation and {relation['total'], *relation['parts']} <= set(c.metrics):
        from .numeric_format import decimal_value
        if decimal_value(c.metrics[relation['total']].value) != sum(decimal_value(c.metrics[k].value) for k in relation['parts']):
            add('EDITORIAL_METRIC_TOTAL','total must exactly equal the mutually exclusive parts',slide.id)
    if all(k in c.charts for k in entry["charts"]):
        from .template_engine import alias_value
        for name,spec in entry.get("chart_aliases",{}).items():
            try:
                capacity(alias_value(spec,c),spec,name)
            except (ValueError,ZeroDivisionError,IndexError,TypeError,OverflowError,DecimalException):
                add("TEMPLATE_CHART_DERIVED",f"{name}: derived callout undefined",slide.id)
    if entry.get("bridge") and all(k in c.metrics for k in entry["bridge"]["keys"]):
        vals=[c.metrics[k].value for k in entry["bridge"]["keys"]]
        if vals[0]<=0 or vals[-1]<=0 or not math.isclose(sum(vals[:-1]),vals[-1],rel_tol=1e-8,abs_tol=1e-8):
            add("TEMPLATE_BRIDGE_TOTAL","start + deltas must equal end; positive start/end required",slide.id)
        total=vals[0]
        for delta in vals[1:-1]:
            total+=delta
            if total<0: add("TEMPLATE_BRIDGE_DOMAIN","bridge cannot cross below zero in this template",slide.id)
    if entry.get("calculation") and set(entry["metrics"])<=set(c.metrics):
        groups=entry["calculation"]["groups"]
        for a,b,result in zip(*groups):
            product=c.metrics[a["key"]].value*c.metrics[b["key"]].value
            parts=c.metrics[result["key"]].value+c.metrics[result["extra_key"]].value
            if not math.isclose(product,parts,rel_tol=1e-8,abs_tol=1e-8):
                add("TEMPLATE_CALCULATION","A×B must equal base+increment for each category",slide.id)
            from .numeric_format import decimal_value,exact_number
            exact_parts=decimal_value(c.metrics[result["key"]].value)+decimal_value(c.metrics[result["extra_key"]].value)
            if len(exact_number(exact_parts))>7: add("TEMPLATE_CALCULATION_LABEL","calculation total exceeds fixed label capacity",slide.id)
