# 独立レビューの修正と再現

`4432d62`へのレビューで、サンプル以外の値を入れた際の計算・表示・coverageに問題が見つかりました。次を修正し、45テストの回帰に含めています。元レビュー作業ディレクトリは変更していません。

| 指摘 | 修正・確認 |
|---|---|
| CAGRが観測点数を年数とみなす | chart.elapsed_yearsを必須化。2020/2022/2024/2026/2028/2030と10年を入力し、35→114/96/81は13%/11%/9%。warm/coolで確認 |
| 増減符号を変えても色が残る | bridgeの色・ラベルを符号に結合。11,+3,-4,+3,13で、warmは増加5A3921/減少A22727、coolは増加1E5F8C/減少D94C68。累積geometryも確認 |
| metricが6有効桁に丸まる | 123456.7と123462.7をそのままnative文字へ。+1/+2/+3も表示。長い小数は1行容量を検査し、縮小や丸めで隠さず拒否 |
| 数値引用の留保をcoverage済扱いする | Datumは数値tokenだけをcoverageへ加算。32.0の引用にestimate only等の未表示文脈があればNUMERIC_CONTEXT_REVIEW/PARTIAL_SOURCE_REVIEWを出す |
| READMEが8型・template未実装のまま | README/design/usage/contractを62実体template＋従来8型へ更新。既存の利用手順を維持 |
| 保存schemaを変更しても検出しない | 保存schemaとruntime modelを比較。template SHA・slot・容量・binding等の契約digestも照合。不一致をCATALOG_SCHEMA_MISMATCHで拒否 |
| Windows cp932のcatalog出力が失敗 | `catalog --out PATH`はUTF-8直接保存。cp932/pipe標準出力はASCII JSON escapeにし、cp932 subprocess回帰で確認。Office QAのセッション制約も[FAQ](faq.md)に記載 |

機械検査の成功は意味の承認ではありません。`mechanical_validation`、`source_coverage`、`semantic_review`は別stateです。validatorの`semantic_review`は常に`not_performed`で、留保を保ったかは構成レビューで確認します。LLMによる意味判定を自動実行していません。

## 再現物

`catalog/qa/review-regressions/`にsource、plan、6枚PPTX、実出力の値・色・geometryを記録したresults.json、PowerPointの6枚PNGと文字境界計測を保存しています。3ケース×2配色をMicrosoft PowerPoint 16.0で描画し、全6枚のpixelsを確認、文字境界overflowは0です。

```sh
python -m unittest tests.test_catalog_review_regressions -v
python scripts/make_review_samples.py
```

標準カタログ124枚と架空ストーリー3枚も再生成・実描画しました。標準カタログの変更PNG6枚を確認し、残り118枚は前版の確認済PNGとSHA一致です。大きな総量に対して小さな増減を入力したprecision検査では、増減の棒が細く見えるのは比率を保持した結果です。適切な図の選択は構成レビューで確認します。

## 既存planの移行

scenario_lines_cagrは`charts.chart_013.elapsed_years`に引用付き期間が必要です。`texts.text_017`の年率見出しは入力欄から外し、期間を含む派生見出しになりました。非年ラベルも期間を省略できません。西暦ラベルは昇順・等間隔で、端点年差と期間が一致する必要があります。引用できる期間が不明なときは補足情報を確認してください。

Pythonが描く入力数値ラベルは保持している数値を丸めず表示します。JSON numberを受ける既存floatモデルのため、末尾0の数や元の数式などの字面は保持しません。派生CAGRは小数0桁、派生平均は小数1桁にhalf-evenで丸める規則をcatalogへ明示し、構成レビューで参照可能にしました。

保存schemaはAI参照用であり、runtimeが任意の保存schemaを実行する設計ではありません。runtimeはPydanticと追加検査を使い、AI参照とのずれをpreflightで検出します。
