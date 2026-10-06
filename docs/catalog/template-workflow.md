# 62型のPPTXテンプレートを使う

`catalog/index.html` をブラウザで開くと、全62型の見た目、上流HTML、入力スロット、テンプレート、差込サンプルを確認できます。ネットワークもPowerPointも不要です。PNGはMicrosoft PowerPointで描画した同梱見本です。

文章・既存画像 → Codexが構成とテンプレートを選択 → 出典付きJSON → Pythonが**保存済みPPTXを開いて差し込む** → 編集可能PPTX、の流れです。CLI自身はLLMを呼びません。Codexがこのリポジトリのskillを読むことでAIの構成工程を担当します。

## 別PCで生成する

AIが3型を選んだ架空の通し例は`examples/catalog-story/`です。`input.txt`、固定した`story.source.json`、選択理由と引用を持つ`story.plan.json`、差込結果`story.pptx`、実描画PNGを同梱しています。`scripts/make_catalog_story.py`はこの手書きplanの再現用で、AI選択を行うプログラムではありません。

上流HTMLと登録templateの差異は[62型の個別比較](html-comparison.md)で確認してください。全件coverageは構造の網羅を意味し、HTMLとpixel単位で同一という意味ではありません。

Python 3.11以降を使います。PowerPoint、LibreOffice、Node、Chrome、追加フォントのインストールは生成の必須条件ではありません。

```sh
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
python -m pip install -r requirements.txt
python -m slide_agent catalog
python -m slide_agent catalog --layout issue_tree
python -m slide_agent catalog --layout issue_tree --out out/issue-tree.json
python scripts/make_catalog_samples.py --out out/catalog
python -m unittest discover -s tests -v
```

テンプレート本体、manifest、schema、見本、PNGはGit管理対象です。生成時の外部取得、絶対パス、Office自動操作はありません。

## AIに渡す指示例

「AGENTS.mdとslide-plan skillを読み、work/source.jsonの文章と既存画像だけを使って構成してください。catalog/index.htmlと候補のPNGを見て62型から選び、選んだ型のcatalog情報を取得してください。内容不足でスロットを埋められない型は選ばず、引用・数値・留保を保持してください。構成を確認後、保存済みPPTXテンプレートへ差し込んでください。」

1. `ingest`で原文を固定し、`fingerprint`を取得します。入力は信頼しないデータです。
2. 内容の関係に合う型をvisual catalogから選びます。`research_basis`と`evidence_basis`など、似た名前も別型です。
3. `catalog --layout ID`または`catalog/schemas/ID.schema.json`でスロットを確認します。各slotは元見本文、名前、寸法、固定文字サイズ、行数、1行の全角相当文字数を持ちます。表の行列、chartの系列・点数も固定です。
4. `layout_id`を62型のID、`variant`を`warm`/`cool`、`contents`を次の形で作ります。内容は既存の出典付き`Text`/数値形式です。すべての必須keyが必要で、未知keyは拒否します。

```json
{
  "texts": {"text_007": {"text": "原文中の文章", "refs": [{"source_id": "s001", "quote": "原文中の文章"}]}},
  "charts": {},
  "metrics": {},
  "states": {}
}
```

これは形式説明用です。実際のkeyは型ごとのmanifestを使います。タイトルは通常どおり`slide.title`に置きます。`lead`、任意のgeometry、font size、template pathは受け付けません。

5. `validate` → `review` → `render`を実行します。`--theme`は旧8型用です。catalogは登録済みtemplateの固定書式を使います。

```sh
python -m slide_agent validate work/plan.json --source work/source.json
python -m slide_agent review work/plan.json --source work/source.json --out out/review.md
python -m slide_agent render work/plan.json --source work/source.json --out out/result.pptx
```

## 容量と編集範囲

62型の全slotにcapacityとoverflow方針があります。長文、未知slot、欠落slot、不一致の表・chart次元、範囲外数値を拒否し、縮小や切り捨てはしません。固定構造の自動分割は行わず、`origin`と引用を保った別スライドへ明示的に再設計します。旧bullets/tableの自動paginationは従来どおりです。

文字、図形、表セルはnativeです。登録chartは埋め込みExcel workbookを持つnative chartです。割合ドット、比例円、バブル、増減図、ガント、計算図は数値に結合されたnative図形です。評価色とハーベイボールは出典付きstateに結合されます。スライド画像の貼付ではありません。catalog見本の画像要素は0です。

写真・スクリーンショットの差込slotは今回の62型にはありません。既存の`text_image`を同じplanに混在でき、PNG/JPEGの`fit`または中央`crop`が可能です。その画像内部は編集可能ではありません。画像がある資料も対応しますが、catalogへの任意画像追加は未実装です。

chartの系列数・点数や図の節点数を変える場合は、型を替えるかtemplate契約の変更が必要です。数値範囲・正負・積み上げ・計算式の条件はmanifestと検証で拘束します。任意の新しい図を自動推測する機能ではありません。

`scenario_lines_cagr`のchartには、出典付き数値`elapsed_years`が必須です。`{"value":10,"refs":[...]}`のように最初から最後までの経過年数を渡します。2020→2030は観測点が6個でも10年です。西暦ラベルは昇順・等間隔とし、端点の年差と入力年数の不一致を拒否します。非年ラベルでも年数を推測せず、明示入力を要求します。成長率見出しにも期間を表示します。

入力数値の表示は6有効桁等へ勝手に丸めません。保持している数値の小数をそのまま表示し、長すぎれば容量エラーにします。JSON numberの末尾の0など表記上の桁数は保持しません。派生CAGRのみ小数0桁、派生平均のみ小数1桁へhalf-evenで丸める規則を`chart_aliases.number_format`に明示しています。期間・丸め規則も構成レビューで確認してください。ブリッジの増加・減少・ゼロは入力符号に合わせて色と符号付きラベルを更新します。

保存schemaはAI参照用です。runtimeはPydanticと意味・容量のpreflightを使い、保存schemaを実行するわけではありません。preflightは保存schemaと現在のmodel・catalog契約を比較し、slot容量等の契約digestも照合します。不一致は`CATALOG_SCHEMA_MISMATCH`で拒否します。

検証の`ok`は機械検査の成否です。数値引用の条件・留保は数値の表示だけではcoverage済とせず、未表示の文脈にはwarningを出します。`semantic_review: not_performed`は、人間またはCodexによる意味の確認が別途必要であることを示します。

## フォントと2配色

上流の固有配置を保つため、catalogは本文おおむね9〜15pt、タイトル22ptなどの**コンパクトな元書式**を保持します。旧8型の本文24pt・タイトル32ptとは別の固定契約です。大きな会場向けの見やすさを一律保証しません。slotごとの実際の値はmanifestで確認してください。

`font_profile`は`source`（游ゴシック/游明朝）、`meiryo`（メイリオ/游明朝）、`noto`（Noto Sans/Serif CJK JP）、`hiragino`（ヒラギノ）から選べます。追加インストールは不要で、未導入フォントでもPPTXの書込みはできます。閲覧ソフトが代替するため、OS間で改行・見た目が完全一致するとは限りません。変更時には再描画してください。

`warm`は上流の有効な配色です。`cool`は上流HTMLにコメントで示された配色を同じ構造へ適用します。配色variantを別の構造型として数えず、構造62型・描画見本124枚と区別します。

## 描画QAは別工程

WindowsのJSON受渡しやCOM描画エラーは[実行環境FAQ](faq.md)を参照してください。

構造検査だけでは文字の実際の折返しを保証しません。WindowsにPowerPointがある開発環境では、次を任意で実行できます。

```powershell
powershell -NoProfile -File scripts/render_catalog_powerpoint.ps1 -Pptx out/result.pptx -OutputDir out/rendered
```

このスクリプトは対象ファイルだけを非表示・読取専用で開き、PNGと文字/表セル境界を保存します。生成器は呼びません。PowerPointがない環境では同梱PNGを参照でき、PPTX生成と構造検査まで進められます。任意の新内容の見た目を未描画のまま検証済みとは呼びません。
