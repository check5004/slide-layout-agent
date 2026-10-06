# 手動編集・企業テンプレート・ページ追加

完成したPPTXに手で加えた修正をplanへ逆取込みする機能はありません。手修正版は別ファイルとして保存します。再生成するときは原文とJSONを修正するのが基本です。

## スライドマスター・フッター・ページ番号

現実装は保存templateのnative図形、表、chartと個別の埋込workbookを、共通の空白layoutを持つ出力へ複製します。元PPTXのslide master/themeを汎用的に継承・統合するimporterではありません。元の直接指定font・色は残るため、PowerPointのマスターやthemeを変えただけでは一括変更されない要素があります。

上部資料名はplan.title、レイアウト記号は登録IDに対応したmetadataです。下部は現在`slide-layout-agent`、`native editable`、生成順のページ番号を通常のnative文字として差し込みます。ページ番号はPowerPointの自動番号fieldではないため、PowerPoint上でページを並べ替えた後は手修正または再生成が必要です。会社名・copyright・任意footerの公開入力項目はまだありません。元公開見本の会社表記は出力footerに引き継ぎませんが、MIT帰属はnotesと同梱LICENSEに保持します。

差込み文字は段落ごとの先頭run書式を使います。元templateの1段落内の混在色・太字範囲を入力文の意味に応じて再現する機能はありません。段落、表セル、図形、chartは編集できます。固定装飾・凡例を変える場合はtemplateの登録更新が必要です。

## 今の資料に1ページ追加する

原文segmentを追加してfingerprintを更新し、既存template IDと引用を持つslideをplanへ追加します。構成確認、validate、render、表示確認をやり直します。これは新しい再利用layoutを登録する作業とは別です。

## 再利用テンプレートを変更する

現在、自動importerはありません。ユーザーがPPTXと変更意図を渡し、エージェントが候補を別作業コピーで調整し、見本を確認してから登録する開発作業です。

生成時はtemplateのSHA-256を照合します。**PPTXを差し替えるだけでは停止します。** 位置・寸法・font変更でslot名を保持した場合も、`catalog/manifest.json`のbounds/font/capacity、template SHA、schema、preview、検証記録の更新が必要です。色だけの変更もSHAとpreviewと描画QAを更新します。

shape名の変更、項目の追加・削除、表の行列、chart系列・点数の変更には、slot定義と必要なbinding・異常系テストの修正が必要です。全型または影響型のgeometry、OOXML、数値整合、実描画を再検証します。勝手にslotを推測して登録しません。

`scripts/build_catalog.py`は固定した公開上流assetから現catalogを再生成する開発用スクリプトです。人が編集したPPTXのimporterではありません。直接編集したtemplateを再ビルドで上書きしないよう、元ファイルと候補を別管理してください。

## 企業テンプレート

企業ファイルをこのpublic repositoryへ入れません。private作業領域でmaster、layout、named shapes、既存chart/表、fontとライセンス、画像の編集範囲を確認し、別途契約と検証を作ります。今回の62型登録が任意の企業templateの無条件drop-in対応を意味するものではありません。

テンプレート候補・変更意図 → エージェントが構造とslotを調整 → 差込見本と制限を提示 → 本人が採用を判断、という手順です。未対応の要素があれば明記します。既存ファイルを上書きして手修正を失わせません。
