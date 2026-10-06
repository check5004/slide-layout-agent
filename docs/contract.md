# JSON contract v1

`python -m slide_agent schema --out out/plan.schema.json` is authoritative for field types; `slide_agent/models.py` defines the discriminated union. Additional fields are rejected, including coordinates, code, URLs and font sizes in a plan. Source citations may contain URLs as inert text, never fetched.

Source bundle: `version`, `title`, `segments` (`id`, exact `text`, optional `citation`, `group`), `images` (`id`, relative `path`, `caption`). Hash the canonical JSON with `fingerprint`. Text mode `slides` gives each heading/body segment the same group.

Plan: `version`, `source_sha256`, `title`, `slides`, `omissions`. Each slide has `id`, `origin`, `layout_id`, `title`, optional `lead`, `rationale` and `contents`. Every displayed text object is `{text, refs:[{source_id, quote}], mode:"verbatim"}`. A rewritten title/body uses `mode:"paraphrase"` and a slide rationale. Quotes must be exact substrings of a real source segment. A verbatim text must occur in one of its quotes. Every data value also has `refs`.

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

`split` supports bullets and table rows, retaining item/row order and all source references. Other overflows require plan revision and stop with a useful error. Continuation IDs use `-p2` etc.; `origin` retains the original group. No content is silently dropped.
