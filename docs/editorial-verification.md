# Editorial 検証記録

2026-10-07、Windows / Python 3.12.10 / python-pptx 1.0.2 / Microsoft PowerPoint 16.0で確認。基点は `5b40212afd8632c0d7258414c648db64926abd70`。旧62型のテンプレート・manifest・個別schema、および既存45テストは変更していません。

## 完成物

- [全76型の一覧](../catalog/editorial/index.html)、[全型PPTX](../catalog/editorial/samples/editorial.pptx)、[代表12型](../catalog/editorial/representative/index.html)
- [16枚の架空調査共有デモ](../examples/editorial-story/index.html)、[編集可能PPTX](../examples/editorial-story/story.pptx)
- [上下配置の同内容12組比較](../examples/editorial-rows/index.html)、[変更前PPTX](../examples/editorial-rows/rows.before.pptx)、[変更後PPTX](../examples/editorial-rows/rows.after.pptx)
- [全型の実描画対応表](../catalog/editorial/coverage.json)、[描画・文字境界の生レポート](../catalog/editorial/render-report.json)、[検証manifest](../catalog/editorial/qa/verification.json)

サンプルは全て架空です。実企業の研究資料や個人の資料は含みません。図・文字・表・グラフはnative要素、画像付き2型の画像は架空UIのPNGです。

## 検証範囲と結果

| 対象 | 枚数 | 実描画 | 文字境界のはみ出し |
|---|---:|---|---:|
| 全layout通常サンプル | 76 | PowerPoint PNG 1600×900 | 0 |
| 代表サンプル | 12 | 同上 | 0 |
| 連続した調査共有デモ | 16 | 同上 | 0 |
| 各text slotの行数・全角容量上限 | 76 | 同上 | 0 |
| 上下配置の変更前／変更後 | 12＋12 | 同上 | 0 |

計204枚。全76型とデモ16枚、上下配置の変更前後12組をコンタクトシートで一枚ずつ目視し、密な6項目・左右・主従・表・引用・グラフなどは原寸PNGも確認しました。縦型時系列の時期欄、結論型の小見出し、5・6項目の行型カードは、全型目視と独立レビューで見つかった語尾だけの折返しを直して再描画しています。

最終QA版はv9です。描画前に6種類のPPTXとplanを版別フォルダへコピーして固定し、その固定コピーを開いてPNGと計測レポートを作成しました。公開用PPTX・固定コピー・レポートのSHA一致を照合しています。以前の版の計測値を更新後のPPTXに適用しません。

上下配置は27型を更新しました。短い見出し・番号・表セルは行の中央、本文＋根拠は一つのまとまりとして配置し、カード内の縦積みは上寄せを維持します。登録済みの行には `row_alignments` で上寄せ／中央寄せを指定できます。2〜6項目、1行と複数行、根拠の有無、行ごとの指定を比較し、変更前後の本文・数値・参照・順序とnativeテキストの一致を検証しました。[容量比較](../examples/editorial-rows/capacity-comparison.json)は全773 title/text slotのフォントと行数・文字幅上限が不変であることを記録します。

文字枠の寸法に加えて上下位置も計測し、いずれもはみ出し0件でした。PowerPointの表セルの `BoundTop` は、この環境ではスライド上の絶対座標ではないため、表セルの上下位置判定には使いません。生の値とnative anchorを記録し、寸法計測・OOXML監査・実際のPNGを併用しています。この範囲は各レポートの `positionMeasurementScope` に明記しています。

境界サンプルは各slotの最大行数と、1行当たりの全角換算上限の整数部分を日本語文字で埋めます。明示改行を含む境界と、通常サンプルの長めの日本語・2行の表注記を分けて確認しています。全ての日本語文章やフォント置換を保証するものではありません。別の対象ビューアで利用する場合はそのビューアでも確認してください。

67テスト（既存45＋追加22）が成功。追加分では全76型のOfficeなし生成、native slot identity・schema・SHA、画像fit/crop、独立したグラフworkbook、項目数2/3/4/5/6、左右の意味対応、空・不足・不明slot・上限超過の拒否、数値の総数と内訳、比較軸、重点項目の根拠、内容不変のvariant選択、4枚反復と理由、注記位置の逸脱と欠落の検出、全773 title/text slotの境界fixture維持を確認しています。さらに行ごとの上寄せ／中央寄せ、本文＋根拠の配置、表セルと番号のnative anchor、未知行や不正anchorの拒否、選択時の指定保持、schemaの行情報改変検知を検証しました。PowerPoint呼出しは開発用QAのみです。

## 内容と選択のレビュー

デモ3〜6枚目は同じ3カードの原案から、`grid → rows → grid → rows` を選びます。入力文字・数値・引用・originを保持し、並列項目の重要度を変更しません。featureカードは優先を示す原文引用と明示的な `emphasis` がある場合だけ候補に入ります。左右を入れ替えただけの型は反復検知上同じ構造です。

合計120件と84/24/12件の内訳は独立指標から分け、合計関係を検証。比較は同じ軸で二案を並べ、引用は発言・解釈・一般化の留保を分離しました。デモに残るpartial-source等の警告もレビュー対象であり、数値一致や正確な引用だけで意味の正しさを認定しません。架空例の原文・plan・選択理由は同じディレクトリで比較できます。

デモの `PARTIAL_SOURCE_REVIEW` は `s009/s012/s013/s014` に残ります。同じ比較軸の反復、120と12のような部分一致、タイトルと表見出し・グラフ項目で繰り返す語句について、既存validatorのspan照合が最初の出現だけを計上するためです。原文・plan・実描画を照合し、比較条件、件数、架空例・効果未検証・同条件ではないという留保が残ることを確認しました。この例のために既存validatorの警告を抑制していません。

独立視覚レビューの指摘を受け、見出し幅、本文と留保の距離、数値と単位、比較軸、合計と内訳、図中説明、グラフ補助線を修正しました。最終的な利用者の好みや実データでの確認はdraft PRで継続できます。

## 再現

```sh
python -m unittest discover -s tests -v
python scripts/make_editorial_samples.py --representative
python scripts/make_editorial_samples.py
python scripts/make_editorial_story.py
python scripts/make_editorial_boundary_samples.py
python scripts/make_editorial_rows.py
```

任意のWindows QAでは `scripts/render_catalog_powerpoint.ps1` の `-Pptx` と空の `-OutputDir` を指定して実描画します。その後 `scripts/publish_editorial_assets.py --pptx ... --plan ... --rendered ... --out ...`（全型一覧には `--gallery`）でPPTXのSHA・枚数・overflowを照合して公開用画像へ反映します。HTMLによる代替描画は用いていません。

ライブラリ保存は正規helperの準備APIがこの実行環境に提供されていなかったため、保存前に停止しました。成果物はこの公開リポジトリのbranch上で確認できます。
