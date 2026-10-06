# 棚卸しと再利用根拠

上流: https://github.com/carnot-tech/consulting-pptx-skill

固定commit: `e5dd04dc6f1acf010046d2fc7319bcf5242a849a`

`catalog/inventory.json`が母数、対象path、section、SHA、variant、非layoutの区分を持ちます。`manifest.json`が各HTML section → Python layout ID → template/slot/schema → sample/preview/検証の対応を持ちます。

| 区分 | 数 | 扱い |
|---|---:|---|
| templates/freeform_parts_16x9.html | 27 section | 全件対象、b01〜b27 |
| templates/freeform_parts_more_16x9.html | 35 section | 全件対象、m01〜m35 |
| tests/fixtures/good_deck.html | 6 section | 検査fixture。独立layoutではない |
| 独立共有partial | 0 | CSSは2つのHTML内に共通ルールとして存在 |
| SuperTemplate_62type.pptx | 70枚 | 型62枚＋カタログ案内8枚。型のnative構造を再利用 |

`warm`/`cool`は配色のvariantです。62型を124型と数えません。似た名前の型も除外していません。記事2の非公開コードは取得・使用していません。

## HTML converter評価

固定commitの`html_to_pptx.py`を基本27型へ実行しました。WindowsではPython UTF-8モードが必要でした。結果は27枚、native table 8個、native chart 0個、PNG画像6個です。b16・b17・b18・b22・b23・b24のSVGがPNGとして残り、内部の数値は編集できません。

このconverterは形と文字を取り出す補助には使えますが、型別slot契約や数値整合を提供しません。主経路には採用せず、上流付属native PPTXをtemplate化しました。付属native PPTXの画像は0で、4型がnative chartでした。追加の棒図3型（うち小図4個）をnative chartへ変換し、現catalogは合計10 chart/workbookを持ちます。

図の配置は同梱native版に沿います。HTMLとnative版はブラウザとOfficeの文字組み、余白、図表表現が同一ではありません。比較用HTMLのPNGを`catalog/source-previews/`に同梱し、型ごとの差異をmanifestに記録します。ピクセル完全一致を目標・達成とはしていません。

既知の意図した変更: カタログ用ヘッダーを出力metadataへ置換、入力に結合したchartの内部padding、数値に結合した図形、集計/CAGR/計算式の整合、上流の重複段落設定の修復、実測で幅不足だった演算子の拡幅。元の構造関係を保存し、62個のラベルを汎用gridへ割り当てて済ませていません。

## ライセンス

上流はMITです。`Copyright (c) 2026 Carnot AI Inc.`とpermission noticeの全文を`catalog/upstream/LICENSE`に保持します。再配布するtemplate、HTML、native元asset、QAヘルパーはこのライセンスの対象です。生成PPTXのcore propertiesとnotesにも帰属を記録します。

独自engineと検証コードはプロジェクト本体のLICENSEに従います。サンプルは公開上流見本と短い架空の差込文だけで、個人情報・企業内部資料・非公開研究は含みません。
