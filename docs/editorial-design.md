# 社内調査共有向け Editorial catalog

日常の調査共有を、要点・説明・根拠・留保の関係が読み取れる形にする新規デザインです。旧62型は reference catalog として保持。新作は14 family・76 layoutです。色違いは数に含めません。

通常は `catalog/editorial/index.html` の実描画と `python -m slide_agent catalog --layout ID` の容量を確認します。JSONに位置や文字サイズを書かず、保存済みの単一スライドPPTXのnamed slotsへ差し込みます。既存の配色指定 `variant: warm/cool` と構造variantは別物です。Editorialの配色は固定で、構造は `layout_id` で選びます。

## Familyと配置

| family | 用途 | 型数 | 構造・項目数 |
|---|---|---:|---|
| cover | 調査目的と入口 | 3 | brief / split / band |
| takeaway | 結論と根拠 | 4 | 左右 / 上段 / 要旨帯 |
| bullets | 論点と説明 | 10 | 2〜6項目、行 / 段組み |
| cards | 所見＋説明＋根拠・留保 | 16 | 2〜6項目、grid / rows、3・4項目の主従、画像付き2項目の左右 |
| diagram | 仕組みと文章 | 4 | 3段の図、左右と図文比率 |
| comparison | 二案の比較 | 4 | 2列 / 左右交換 / 2段 / マトリクス |
| steps | 順序のある手順 | 8 | 2〜4段の横 / 縦、5・6段の縦 |
| timeline | 時間上の段階 | 4 | 3・4段、横 / 縦。順序は反転しない |
| metrics | 数値と説明 | 6 | 2・3・4指標の横、3指標の縦、総数＋3内訳の2配置 |
| table | 確認事項の一覧 | 4 | 2列×3・4・6行、3列×4行 |
| chart | 数量比較と所見 | 4 | 4項目の横棒＋左右、縦棒＋上下 |
| quote | 引用と解釈 | 3 | 引用を左 / 右 / 上に置く |
| summary | まとめと次の確認 | 3 | 3項目の段組み / 行 / 主従 |
| agenda | 調査範囲 | 3 | 3項目の行 / 段組み / 目次帯 |

本文18〜21ptを中心に、項目見出し20〜23pt、ページ見出し28pt、短い補足12〜15pt。表は18pt、図中の短い補助ラベルは14ptです。文字は自動縮小しません。カード内外の間隔、左端、段落のまとまりを共通化し、見出しと説明を一つのレコードとして扱います。項目数が多い型は短い記述用です。長文を6カードへ押し込まず、少ない項目数または継続ページへ再構成します。

画像付きカードはsourceに登録されたPNG/JPEGのみを使います。`images.image` の `image_id` と `image_mode: fit/crop` を指定します。標準はfit。図解・本文・表・グラフはnativeで編集でき、写真の内容は画像のままです。見本の検索画面はプログラムで描いた架空UIです。

## 内容を保ったvariant選択

Codexが原文からfamily・項目数・意味上の関係を決めます。Pythonはその判断を置き換えず、同family・同じsemantic slotsの候補について、全slot容量・表/グラフの次元・画像の有無を検証します。

```sh
python -m slide_agent catalog --family cards
python -m slide_agent catalog --layout ed_cards_grid_3 --out work/layout.json
python -m slide_agent select-variants work/plan.json --source work/source.json --out work/selected.plan.json
python -m slide_agent review work/selected.plan.json --source work/source.json --out work/review.md
python -m slide_agent render work/selected.plan.json --source work/source.json --out out/result.pptx
```

`select-variants` は文字・数値・引用・origin・意味上のslot名を変更しません。内容が収まる候補から、直近3枚で少ない視覚構造、使用回数、元の選択、IDの順で決定します。乱数も強制ローテーションも使いません。収まる候補がなければ拒否します。

鏡像だけの左右variantは同じ視覚構造です。均等列カード・指標・横手順も共通構造として扱います。連続4枚が同じ構造になると通常のvalidate/renderが `VARIANT_REPETITION` で停止します。やむを得ない場合は `repetition_reason` を残し、検証レポートにも候補と棄却理由を表示します。`allowed_layouts` で候補を限定する場合は、読み順や比較条件を示す `constraints_reason` が必要です。色やIDの変更では反復検出を回避できません。

`*.selection.json` と通常のvalidation reportには、family選択理由、構造variant、項目数、semantic slot mapping、適合候補、容量等による棄却理由、反復理由を記録します。`rationale` はAIが判断したfamily選択理由のまま残します。数値の一致は意味の保証ではなく、従来どおり引用の条件と留保を別途確認します。

カードの `item_1_heading / item_1_body / item_1_evidence` は配置が変わっても同じ内容です。比較は `a_* / b_*` の識別を保ちます。図の `node_1 / node_1_detail`、手順・時間軸の番号は順序を維持します。familyをまたぐ自動変換はありません。

固定文字サイズ、全角換算の1行容量、行数、template SHA、slot mappingをmanifest/schemaに保存。契約の不一致、空文字、空白だけの文字、不明・不足slot、項目数違い、上限超過を拒否します。計算上の容量は保守的な目安で、実描画の確認も必要です。

並列項目に主従を作りません。featureカードとsummaryのsplitには `emphasis`（`item_id: item_1`、優先度を示す原文の `refs`、`reason`）が必要です。原文の正確な引用であることを検証し、実際に主従を裏付けるかは意味レビューへ残します。強調のない項目は対等なgrid/rows候補だけから選びます。

総数と相互排他の内訳は `metrics_breakdown_top_3 / metrics_breakdown_left_3` を使い、総数＝内訳合計を厳密に検証します。通常の指標カードは互いに重なり得る独立指標です。比較は左右で同じ軸名を要求します。

本文が短いときに留保だけがカードの下端へ離れないよう、登録済みの `text_flow` は本文と直後の留保を一つのまとまりにします。カード内の縦積みは上詰め、横並びの行ではまとまり全体を行中央へ配置します。本文と留保のテキスト自体は上寄せで、短い見出し・番号・時期・独立した本文欄はnativeのvertical anchorで中央寄せにします。幅・文字サイズ・順序・容量・shape名は変えず、監査で行範囲と注記間隔の逸脱を拒否します。

行形式では `row_alignments: {"item_1": "top", "item_2": "middle"}` をplanに指定できます。表の行は `row_0`（見出し）から始まります。省略時は各行の登録済みdefaultを使用します。指定できるキーは `catalog --layout ID` の `rows` にあるものだけで、schema/validatorが未登録の行や値を拒否します。この指定はvariant選択でも保持され、該当行のない配置への切替を防ぎます。全てを中央寄せするのではなく、長い説明を順に読む段組み・カード内・引用などは上寄せのままです。

[同じ内容による変更前後12組](../examples/editorial-rows/index.html)で、2〜6項目、1行/複数行、本文＋根拠、番号・日付・指標・表、明示的な上寄せと行別指定を確認できます。

## 開発と再現

```sh
python scripts/build_editorial_catalog.py
python scripts/make_editorial_samples.py --representative
python scripts/make_editorial_samples.py
python scripts/make_editorial_story.py
python -m unittest discover -s tests -v
```

buildは保守用です。通常生成は保存済みPPTXを読み込み、描き直しません。旧62テンプレートの実体は変更しません。実描画PNGは任意の開発用PowerPointスクリプトで生成した後、`scripts/publish_editorial_assets.py` がPPTXのSHA・ページ数・overflowを照合して一覧へ反映します。通常生成・構造検証・デモ再現にはOfficeを使いません。

16枚の架空デモは `examples/editorial-story/`。3〜6枚目が同じ3カードの原案から選ばれる例で、原案・最終plan・出典・選択理由を保存します。スクリプトはCodexによる手書き構成の再現用で、外部AIを呼ぶ自動調査機能ではありません。
