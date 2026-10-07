# JSON contract v1

New research-sharing decks normally use the **76 Editorial templates** described in [editorial-design.md](editorial-design.md). `EditorialSlide` extends the existing saved-template contract with semantic family/variant metadata, optional `allowed_layouts` plus `constraints_reason`, and `repetition_reason`. `EditorialContents` also supports exact named `images` slots (`image_id`, explicit `fit/crop`) where registered. The old 62 `CatalogSlide` schemas and template files remain separate. No arbitrary paths, coordinates, font sizes or styles are accepted.

The `select-variants` command filters same-family candidates by exact semantic slots, item count, image presence and capacity, then chooses a recent-structure tie-break. It preserves all content and references. Normal validation includes `layout_selection` with candidates/rejections and stops unreasoned four-slide repetition (`VARIANT_REPETITION`). Color changes and mirror-only variants do not count as distinct visual structures. A documented exception remains visible as `VARIANT_REPETITION_ACCEPTED`.

Editorial rows also accept optional `row_alignments`, mapping only registered row IDs to `top` or `middle`. Omitted rows use the template's role-aware defaults. A body and its evidence stay together as one bounded group; short labels, numbers and table cells use native vertical anchors. Explicit choices survive variant selection, and candidates without those row IDs are rejected. Unknown IDs, other anchor values and overflowing groups are rejected. Font sizes, reading order and text capacities do not change. See the [same-content before/after review](../examples/editorial-rows/index.html).

`python -m slide_agent schema --out out/plan.schema.json` is authoritative for field types; `slide_agent/models.py` defines the discriminated union. Additional fields are rejected, including coordinates, code, URLs and font sizes in a plan. Source citations may contain URLs as inert text, never fetched.

Source bundle: `version`, `title`, `segments` (`id`, exact `text`, optional `citation`, `group`), `images` (`id`, relative `path`, `caption`). Hash the canonical JSON with `fingerprint`. Text mode `slides` gives each heading/body segment the same group.

Plan: `version`, `source_sha256`, `title`, `slides`, `omissions`. Each slide has `id`, `origin`, `layout_id`, `title`, optional `lead`, `rationale` and `contents`. Every displayed text object is `{text, refs:[{source_id, quote}], mode:"verbatim"}`. A rewritten title/body uses `mode:"paraphrase"` and a slide rationale. Quotes must be exact substrings of a real source segment. A verbatim text must occur in one of its quotes. Every data value also has `refs`.

The following table covers the eight legacy geometry layouts. The 62 saved native templates additionally use `CatalogSlide` with `variant`, `font_profile`, and exact per-layout `texts/charts/metrics/states` keys. Use `catalog/manifest.json`, `catalog/schemas/ID.schema.json` and the [catalog workflow](catalog/template-workflow.md). The runtime rejects a saved schema that differs from the current model/catalog contract; capacities and bindings are included in its contract digest.

| layout_id | contents | Limit / intent |
|---|---|---|
| title | items: Text[] | 0–3 brief introductory statements |
| bullets | items: Text[] | 1–5 items; CLI split supports pagination |
| text_image | items, image_id, image_mode | 1–4 items; fit or crop; existing PNG/JPEG |
| comparison | left_title, right_title, left, right | 1–3 statements on each side |
| process | steps: Text[] | 2–4 ordered steps, 48 characters each |
| table | headers: Text[], rows: Text[][] | 2–3 columns, 1–6 rows; split supports pagination |
| chart | categories: Text[], series: [{name:Text, values:[{value,refs}]}], note:Text | Native clustered columns, 1–6 categories, 1–2 series |
| closing | items: Text[] | 1–3 closing statements / next steps from source |

Theme is a separate trusted configuration. Default: no lead, 32 pt title, 24 pt body, 20 pt table, Meiryo, navy / teal / off-white. `lead_mode:true` allows a short 18 pt lead without enforcing any article's editorial doctrine. Long source IDs and excess slide count are rejected. Native chart axes can use generated tick values and slide numbers are generated metadata; neither is source content.

The default validation rejects newly introduced numeric tokens and missing original numeric tokens. Explicitly omitted sources are excluded only after review. It does not prove semantic entailment. Qualifiers, causal claims and citation accuracy must also be reviewed. Missing units or uncertainty must not be inferred away.

`ok`/`mechanical_validation` describe mechanical checks only. Numeric evidence covers its numeric tokens, not all surrounding quoted text. Unrepresented numeric conditions trigger `NUMERIC_CONTEXT_REVIEW` and `PARTIAL_SOURCE_REVIEW`. `source_coverage` and `semantic_review` are separate; semantic review remains `not_performed` until the operator performs it outside this validator.

Catalog input numeric labels preserve the stored numeric value without six-significant-digit rounding. Exact labels exceeding their one-line capacity are rejected. Derived CAGR (0 decimal places) and mean (1 decimal place) use explicit half-even policies in `chart_aliases.number_format`. `scenario_lines_cagr` requires a source-backed positive `elapsed_years` value; point count is not duration. Year labels must be increasing and evenly spaced, and their endpoint difference must agree with the explicit duration. Non-year labels still require the duration.

`split` supports bullets and table rows, retaining item/row order and all source references. Other overflows require plan revision and stop with a useful error. Continuation IDs use `-p2` etc.; `origin` retains the original group. No content is silently dropped.
