# 関係図・通信図の検証

2026-10-07。新しい25型、利用場面6枚、上限文字量25枚、既存Editorialの読者表示回帰76枚、計132枚をMicrosoft PowerPoint 16.0で描画しました。全ページを目視し、図形の文字と表セルの寸法超過0件、図形の文字位置超過0件です。登録文字サイズのままで検証し、自動縮小は使っていません。

## 成果物と再現用入力

- [全25型の実描画一覧](../catalog/relations/index.html)・[コンタクトシート](../catalog/relations/contact-sheet.png)・[編集可能PPTX](../examples/relations/relations.pptx?raw=true)
- [利用場面6枚](../examples/relations/usage/index.html)・[PPTX](../examples/relations/usage/relations.pptx?raw=true)：2主体1辺、利用者から内包された機能を経由する照会、独立4主体の8イベント、一対一／一対多、共有URL、3方式比較
- [4主体の単独plan](../examples/relations/sequence-four.plan.json)・[対応source](../examples/relations/sequence-four.source.json)
- [上限文字量PPTX](../catalog/relations/qa/boundary.pptx?raw=true)・[plan](../catalog/relations/qa/boundary.plan.json)・[source](../catalog/relations/qa/boundary.source.json)・[境界コンタクトシート1](../catalog/relations/qa/boundary-contact-01.png)、[2](../catalog/relations/qa/boundary-contact-02.png)、[3](../catalog/relations/qa/boundary-contact-03.png)、[4](../catalog/relations/qa/boundary-contact-04.png)、[5](../catalog/relations/qa/boundary-contact-05.png)
- [検証集計とSHA-256](../catalog/relations/qa/verification.json)・[80テスト結果](../catalog/relations/qa/tests.txt)・[安全な拒否6例](../examples/relations/rejection-fixtures.json)

上限文字量の入力は全角文字を容量まで詰めた計測用fixtureです。通常の説明文の見本は全型一覧・利用場面6枚を参照してください。いずれも架空で、実在の業務資料を含みません。

## 構造と表示

80テストには従来67テストと新規13テストを含みます。既存の数字精度、CAGR期間、増減符号色、引用留保、schema整合、cp932、行の上下寄せ、内容を変えないvariant選択の回帰を維持しています。

新規テストは2〜5主体、4/8イベント、左右反転、包含／同管理主体、共有参照、自己処理、等幅3方式と比較軸の対応を検証します。未知端点・重複ID・孤立主体・非対応の辺・不連続な囲み・空欄・上限超過・任意座標を拒否します。OOXMLの矢印、ラベル、囲み、欠落した辺を変更する負例では監査が失敗します。読者表示はクリック可能なHTTP(S)資料名・版・確認日とnotesを照合し、明示したQAモードでは内部表示を維持します。

nativeコネクタの始終点を確認し、左右反転でnode ID・辺の向き・イベント順・囲み所属を保持します。一対一／一対多のパネル見出しも対応する図と一緒に反転する回帰テストを入れました。目視ではこの対応、矢印方向、自己処理の折返し、囲み見出し、表外の注記、ラベルと線の間隔を確認しています。

独立レビューで8イベント型の最終自己ループが囲みの下端を越える指摘を受け、6型のイベント間隔を0.3975インチから0.39インチに調整しました。16ptの1行ラベルの枠高と自己ループ・矢印先端の形も調整し、途中のループが次のラベル背景に隠れないようにしています。フォント・文字数上限・主体位置・囲み・planの形式は変えていません。最終QA v4では通常25枚・上限文字量25枚・利用例6枚を再描画し、既存Editorial 76枚は同一PPTXのSHAを照合して前回の描画を引き継いでいます。

回帰テストでは3〜5主体、4/8イベント、左右、包含／同管理の全ての登録囲みと最終送受信先、計1,400条件を検査します。線幅と中サイズ矢印先端の余裕を含む範囲が図・内部の囲みを越える旧配置と、次のイベント行に重なる配置は`RELATION_GEOMETRY`で拒否します。修正後の最終自己ループは、この保守的な範囲でも囲み下端から0.05インチ（1600×900描画で6px相当）内側です。[修正前後の下端拡大](../catalog/relations/qa/self-loop-before-after.png)も保存しています。

既存reference 62型とEditorial 76型のtemplate PPTX・個別slot schema・manifestは、この追加の前後で同一です。新型は独立した`catalog/relations/`にあります。全体Plan schemaと通常出力の読者表示を更新しています。

## 実行方法と限界

```sh
python -m unittest discover -s tests -v
python scripts/make_relation_samples.py --out out/relations
python scripts/make_relation_samples.py --stress --out out/relations-boundary
python scripts/make_relation_samples.py --usage --out out/relations-usage
python -m slide_agent validate examples/relations/sequence-four.plan.json --source examples/relations/sequence-four.source.json
```

これらの生成・検証はOfficeなしで動きます。CIのUbuntu / Python 3.11・3.12でも通常・上限文字量生成を実行します。PowerPointは保存したPPTXを開く描画QAにだけ使用しています。

測定対象は計1,910箇所。表セルの`BoundTop`はこのPowerPoint環境でスライド絶対座標にならないため、表は幅・高さの計測、nativeセルanchorの監査、画像の目視を組み合わせています。別環境でフォントが置換された場合の同一改行は保証しません。

validatorは原文から因果・包含・所有の意味を推論しません。`RELATION_SEMANTIC_REVIEW`と`PARTIAL_SOURCE_REVIEW`を残し、計画者が原文と図の対応を確認します。有限の登録構造と最大1個の囲みが対象で、任意のグラフを自動配置する機能ではありません。独立レビューや個別資料の意味確認は、この実装者による架空fixtureのQAとは別です。
