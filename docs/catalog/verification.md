# 全62型の検証記録

対象は`carnot-tech/consulting-pptx-skill` commit `e5dd04dc6f1acf010046d2fc7319bcf5242a849a`です。`freeform_parts_16x9.html`の27節と`freeform_parts_more_16x9.html`の35節を別々の62型として登録しました。テストfixture 6節、native見本のナビゲーション8枚、共有CSSは構造型の母数に含めません。類似型の除外は0です。

## 配布物と検証範囲

| 内容 | 結果・証跡 |
|---|---|
| 実体template | `catalog/templates/`に1枚PPTX×62。読込時SHA照合 |
| coverage | `catalog/coverage.json`にsource file/section、layout、variant、sample slide、PNG hash、検証状態124件 |
| サンプル | warm/cool各62枚のdeckと、型別単独PPTX×62 |
| native要素 | 1配色deckで文字図形1105、表10、chart10、埋込workbook10、画像0 |
| 実描画 | Microsoft PowerPoint 16.0、1600×900、124/124枚。`catalog/qa/*-powerpoint.json` |
| 文字境界 | 全native文字と表セルをPowerPointで測定。boxを1pt超えてはみ出すもの0 |
| pixels | 全124枚と上流HTML62枚を目視比較。白紙なし。chartラベル・接続・色・図形配置を確認 |
| 最終差分 | 独立レビュー修正で各配色の24/34/35枚目が変化。ほか118枚は前版の確認済PNGとSHA一致。変化した6枚を再確認 |
| 異常系 | 全62型のoverflow、未知/欠落slot、任意geometry拒否。全metric範囲、state、chart次元、数値整合、固定geometry改変も検査 |
| 差込み残存 | 全62型の文字・数値を変え、slotの値、native chart cache、埋込workbookを検査。元のsampleを残さない回帰あり |
| 通し例 | `examples/catalog-story/`の架空3枚。型選択理由と引用→保存template読込→出力→PowerPoint実描画。文字境界overflow 0 |
| 帰属 | MIT全文を`catalog/upstream/LICENSE`と全template/生成catalog slideのnotesに保存 |

自動テストは`python -m unittest discover -s tests -v`で45件です。大量の型・slotの組合せはsubTestで同じテスト内から検証します。`scripts/make_demo.py`で既存8型の架空fixtureも再生成しました。独立レビューの[修正と再現](review-fixes.md)も参照してください。

## clean cloneの実行結果

mainのdocs更新`4860ab7`を取り込んだcode commit `931aed779fc76dea21003b5f133be7450d68e44e`を別ディレクトリへcloneし、新規venvに`requirements.txt`だけをインストールして検査しました。Python 3.12 / python-pptx 1.0.2 / pydantic 2.12.5 / Pillow 11.3.0で38テスト成功。全62型×2配色と型別62PPTXも再生成・構造監査に成功しました。

この再生成では`win32com`/`comtypes` importを無効化し、`subprocess.Popen`も例外で禁止しました。Office・Node・ブラウザを呼ばずに完了しています。これはWindowsホスト上でOffice連携を禁止した確認です。Linuxの実行は`.github/workflows/catalog.yml`のUbuntu / Python 3.11・3.12 CIで別途確認できます。CIの結果はPRのChecksを参照してください。

同梱galleryはネットワークを使わず、保存済みPNGを表示します。新規生成するPPTXは同梱の実templateから作るため、上流repoを別にcloneする必要もありません。

## 再実行

```sh
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python scripts/make_catalog_samples.py --out out/catalog --individual
python scripts/make_catalog_story.py
python -m slide_agent catalog --layout chart_insight
python -m slide_agent validate examples/catalog-story/story.plan.json --source examples/catalog-story/story.source.json
```

生成・構造検査はOfficeも外部LLMも使いません。CodexがAGENTS/skillとvisual catalogを読んで構成を考え、typed planを作る方式です。Python自身にLLM呼出し機能はありません。

## 限界

HTMLとnative見本の差異は[型別62件](html-comparison.md)に記録しています。構造を全件網羅していてもpixel完全一致ではありません。実描画の検査済みは同梱サンプルについてです。長さ・font・数値を変えた任意の新内容は再検証してください。

文字境界の自動測定はchart内部の文字を含まないため、chartは実画像も確認しました。許容値はtemplateのcompact fontと固定構造を前提にしています。missing fontの代替は閲覧環境に依存します。

写真・SVG等を貼り込んだスライドはこのcatalogにはありません。既存`text_image`に使う画像の内部は編集できません。master/themeの継承、固定footer、生成時の通常文字ページ番号の制限は[手動編集の説明](template-maintenance.md)を参照してください。
