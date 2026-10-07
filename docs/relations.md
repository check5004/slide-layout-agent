# Nativeの関係図・通信図

`catalog --collection relations` は、既存Editorial 76型と独立した25型です。保存済みPPTXの図形・線・文字・表を使います。自由座標、任意の辺追加、画像化、文字縮小は受け付けません。原文から方向・包含・時系列を判断するのは計画者で、validatorは意味の正しさを推定しません。

| ID | 主体・対応構造 | 上限 |
|---|---|---|
| `rel_pair_2_lr/rl` | 2主体の一方向または往復 | 各方向1辺、計2辺 |
| `rel_fan_3/4/5_lr/rl` | nodes[0]と各末端の一対多。逆向きなら複数主体から共有参照 | 3〜5主体、各方向1辺、最大8辺 |
| `rel_spoke_4_lr/rl` | nodes[0]が中央、[1]が左、[2]が右、[3]が上 | 中央と他主体の各方向1辺、最大6辺 |
| `rel_pair_fan_5_lr/rl` | 独立した2パネル。[0]↔[1]、[2]↔[3]/[4] | 各方向1辺、最大6辺 |
| `rel_sequence_3/4/5_4_lr/rl` | 3〜5主体、上から下への通信・自己処理 | 1〜4イベント、各ラベル最大2行 |
| `rel_sequence_3/4/5_8_lr/rl` | 同上、短いラベルの通信列 | 1〜8イベント、各ラベル1行 |
| `rel_compare_three` | 軸列＋等幅3方式、共通3比較軸 | 4列×4行のnative表、表外注記 |

IDの選択肢は `/` を展開します。たとえば `rel_sequence_4_8_lr`。`lr/rl` は空間配置だけを反転し、node ID・edge source/target・イベント配列順・囲みの所属・本文・参照は変えません。ミラーは反復検知上同じ見た目です。時系列を逆転したり、表の方式を優劣順に並べ直したりしません。

## Plan

全ての表示文字は通常の `Text = {text, refs:[{source_id, quote}], mode}` です。下の `Text` は型の略記で、実際のJSONではTextオブジェクトを入れます。単独で実行できる4主体の架空例は [sequence-four.plan.json](../examples/relations/sequence-four.plan.json) と [対応source](../examples/relations/sequence-four.source.json) です。[利用場面6枚の実描画](../examples/relations/usage/index.html)・[PPTX](../examples/relations/usage/relations.pptx?raw=true)・[検証記録](relations-verification.md)も参照してください。

```text
{
  "layout_id": "rel_sequence_4_8_lr",
  "contents": {
    "texts": {"context": Text, "caveat": Text},
    "network": {
      "nodes": [{"id": "client", "label": Text}, ...],
      "edges": [
        {"id": "request", "source": "client", "target": "server", "label": Text},
        {"id": "check", "source": "server", "target": "server", "label": Text}
      ],
      "boundaries": [
        {"id": "application", "kind": "contains", "label": Text, "members": ["client"]}
      ]
    }
  }
}
```

`nodes`の配列位置がテンプレート内の位置に対応します。表示名やURLではなく、IDをedgeから参照します。2主体が同じURLを名乗る図は、共有URLのnodeを一つにして両主体から辺を向けます。表示名の重複だけで同一主体と解釈しません。通信図だけは同じ主体間の繰返しイベントと自己処理を許可し、順番を配列で表します。関係図では同一方向の重複辺、非対応の辺、孤立したnodeを拒否します。

`boundaries`は0〜1個。`contains`は内包、`same_owner`は同管理主体で、矢印とは別に描きます。囲みはedgeの端点にできません。関係図では登録済みの単独nodeまたはfan末端群、通信図では連続するlaneの範囲だけを囲めます。2パネル型では囲みを受け付けません。中央への縦矢印があるspokeでは囲み見出しを下側に配置します。

`rel_pair_fan_5_*`のtextsには `left_caption` と `right_caption` も必須です。これらはlr版の左側の一対一、右側の一対多パネルを指す意味上の名前です。rl版ではパネルと見出しを一緒に反転するので値を交換しません。それ以外の新図型はcontextとcaveatが必須です。3方式比較はnetworkを持たず、次を使います。

```text
"comparison": {
  "axes": [Text, Text, Text],
  "methods": [
    {"id": "a", "label": Text, "values": [Text, Text, Text]},
    {"id": "b", "label": Text, "values": [Text, Text, Text]},
    {"id": "c", "label": Text, "values": [Text, Text, Text]}
  ]
}
```

各valuesの順序は共通axesと対応します。空セルを埋めるために事実を作らず、原文が未記載を明示する場合だけその記載を引用します。

## 容量・選択・監査

新図型のタイトルは28pt・最大2行。既存76型のタイトル容量は変えません。主体名18pt、通信ラベル16pt、関係図の短い辺ラベル15pt、比較本文18ptです。通信ラベルの横幅は実際の送受信lane間の幅、自己処理は固定の1.8インチです。4イベント型は2行、8イベント型は1行までです。長いラベルは4イベント型へ切り替えるか、出典を保持して明示的に分割します。

```sh
python -m slide_agent catalog --collection relations
python -m slide_agent catalog --layout rel_sequence_4_8_lr --out work/contract.json
python -m slide_agent select-variants work/plan.json --source work/source.json --out work/selected.json
python -m slide_agent render work/selected.json --source work/source.json --out out/deck.pptx
```

選択は同じ構造・主体数の候補だけを調べ、辺と囲みの適合、実際のラベル幅を検証します。内容は変更しません。定義外の主体数、辺、囲み、重複ID、不明端点、空文字、容量超過は拒否します。保存schemaはnode数・イベント上限と契約hashを含みます。OOXML監査はplanを保存済みtemplateに再適用し、図形identity・座標・arrowhead・反転・ラベル・表の軸対応・囲みstyleを照合します。notesを含むファイルそのものの意味の正しさや改ざん耐性を保証する機能ではありません。

通信線が中間laneを横切ることはシーケンス図の仕様です。そのlaneを送受信先とは解釈せず、矢印の始終点を読みます。同じイベント行に複数メッセージは置きません。ラベル背景でlifelineとの文字衝突を避けます。線と文字の関係は実描画でも確認します。

## 読者向けの出典表示

Planの `display` はデフォルトで `reader`。内部layout ID・`native editable`を消し、既存footer枠に資料名リンクと版・確認日を置きます。資料が未指定なら自動で名称を作りません。metadataは出典本文の代わりではなく、本文のrefsは引き続き必須です。QA表示は `{"mode":"qa"}` で残せます。

```json
"display": {
  "mode": "reader",
  "footer": "",
  "citations": [{
    "source_ids": ["s001", "s002"],
    "title": "架空設計書",
    "url": "https://example.com/fictional-design",
    "version": "v1",
    "accessed_on": "2026-10-07"
  }]
}
```

そのslideが引用するsource_idsに対応した資料だけを表示します。版や確認日は入力値を使い、推定しません。URLはHTTP(S)のクリック可能なhyperlinkとして保存し、生成中にアクセスしません。source_ids不明、日付不正、footer容量超過は拒否します。既存referenceの狭いfooterでは短い正式略称を指定するか別layoutを選びます。リンク・metadataと元のcitation文字列はnotesにも残ります。private資料の登録先は利用者の非公開作業領域であり、この公開fixtureに混ぜません。
