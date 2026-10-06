# スライド作成AIワークフロー

調査済みの文章・スライド別文章・手持ち画像から、Codex が構成とレイアウトを選び、Python / python-pptx が編集可能な PowerPoint を生成する小さな開発プロジェクトです。Web UI、調査機能、有料 API、API キーは不要です。

すぐ見る： [生成済みデモPPTX](examples/demo.pptx) · [全8レイアウトのプレビューと検証記録](docs/verification.md) · [設計JSON](examples/demo.plan.json) · [Codex用スキル](.agents/skills/slide-plan/SKILL.md)

![編集可能な比較レイアウトのPowerPoint実描画](examples/rendered/slide-04.png)

## 設計

```
原文 + 画像 → source.json（固定原文・ID）
           → Codex + skill（構成案・layout_id・原文引用）
           → plan.json（型付き中間表現）→ 人間が構成を確認
           → validate → 固定 geometry の python-pptx renderer
           → OOXML / geometry / 文字量検証 → 実描画レビュー → 手直し
```

AI は `layout_id` と内容を選びます。座標、任意 Python、外部 URL ダウンロードを JSON に渡せません。AI が作る `plan.json` と、独立した原文 `source.json` の SHA-256 を結び付けます。出典・主張・数値・留保を勝手に補完しません。引用参照、欠落、意図した省略、分割を検証レポートに残します。意味の同等性は完全自動検証できないため、人間の確認が必要です。

初版の AI 経路は **Codex がプロジェクトスキルを読む方式**です。CLI 自体に LLM/API 呼び出しはありません。デモ生成スクリプトは固定データの再現であり、自律 AI planner ではありません。

`examples/prose.plan.json` と `examples/slides.plan.json` は、このプロジェクト作成時に Codex が原文を読んで作った小さな構成例です。前者は未構造文章＋画像、後者はスライド境界を指定した文章を扱います。`demo.plan.json` は全8種の renderer を検証する固定フィクスチャです。

## 受入基準

- 日本語 16:9、8 種の layout、共通フォント・配色・余白。
- title / bullets / text_image / comparison / process / table / chart / closing。
- 本文 24 pt、タイトル 32 pt、表 20 pt が標準。本文 22 pt・タイトル 28 pt・表 18 pt 未満を拒否。出典・補助ラベルは 12–14 pt。
- 長文は分割または停止。自動縮小・切り捨て・はみ出し隠蔽はしない。
- 文字・図形・表・対応する縦棒チャートは native editable。写真・入力画像の内部画素は編集不可。
- 未構造文章とスライド別文章を読み込み、画像を相対パスで添付できる。
- 数値の捏造・欠落、不正 layout、欠落画像、範囲外パス、表の行超過を検出。
- 全スライドを画像化した PPTX は作らない。
- OOXML 検証と実描画確認を区別して報告する。未描画を描画済みとは呼ばない。

## 利用方法

Python 3.11 以降を用意し、このフォルダで実行します。依存は仮想環境だけに入れます。

```sh
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
python -m pip install -r requirements.txt
python scripts/make_demo.py
python -m slide_agent validate examples/demo.plan.json --source examples/demo.source.json
python -m slide_agent review examples/demo.plan.json --source examples/demo.source.json --out out/review.md
python -m slide_agent render examples/demo.plan.json --source examples/demo.source.json --out out/demo.pptx
python -m unittest discover -s tests -v
```

### 自分の文章から作る

1. `work/` に調査済み文章と画像を保存（Git 管理対象外）。
2. 未構造文章は段落で区切る。スライド別文章は Markdown の `##` をスライド境界にする。
3. 下記で原文 ID を固定する。添付画像は入力フォルダ内に置く。
4. Codex に「`AGENTS.md` と `.agents/skills/slide-plan/SKILL.md` を読み、work/source.json から構成案と plan.json を作成して」と依頼。
5. `review` の構成案と省略・言い換え警告を確認し、必要なら Codex に直してもらう。
6. `render` し、PowerPoint または LibreOffice で全ページを確認。PowerPoint で直接手直しするか、JSON を直して再生成する。

```sh
python -m slide_agent ingest work/input.txt --mode prose --out work/source.json
python -m slide_agent ingest work/slides.md --mode slides --image photo=work/photo.png --out work/source.json
python -m slide_agent schema --out out/plan.schema.json
python -m slide_agent review work/plan.json --source work/source.json --out work/review.md
python -m slide_agent render work/plan.json --source work/source.json --out work/result.pptx
```

`--image ID=PATH` は繰り返し指定可能。初版は PNG/JPEG のみ。画像ファイルは source.json のフォルダ以下に置きます。`render` は検証を必ず通し、明示的な省略がある場合は `--accept-omissions` が必要です。言い換えは警告となり、意味と留保を人間が確認します。既存出力を上書きする場合は `--force` を指定します。

標準は副題なしです。`--theme examples/theme.with-lead.json` で lead（短い要旨）を許可できます。AI は必須情報をタイトルや lead に移し替える際も出典を保持してください。

## 検証と制約

`validate` は型・原文 SHA・引用・数値・文字容量・画像・範囲を確認します。`render` は追加で PPTX の native 要素、geometry、文字サイズ、auto-shrink 不使用を検査し、`*.validation.json` を出力します。文字容量は日本語の文字幅を保守的に見積もる事前検査です。実フォントや Office の描画差まで証明するものではありません。

Windows + PowerPoint がある場合は、生成済みファイルを次のスクリプトで PNG に書き出せます。既存プレゼンテーションは操作しません。

```powershell
powershell -NoProfile -File scripts/render_powerpoint.ps1 -Pptx out/demo.pptx -OutputDir out/rendered
```

`out/rendered/render-report.json` に実描画と PowerPoint の文字境界測定を保存。native chart / table は別途目視確認してください。画像は fit または中央 crop を選択。最大 100 スライド、1画像 20 MB / 40メガピクセル、テキスト入力 2 MB、JSON 5 MB。表は 3 列まで・データ 6 行/ページ、chart は 6 分類・2系列まで。追加 layout と provider は将来の拡張点です。

参照はソース ID 単位・引用範囲単位で追跡します。数値以外の意味の変更、誤った関連付け、留保の弱まりは完全自動判定できません。部分未使用を警告し、全ソースを使用または明示省略する必要があります。

## 公開範囲

公開対象は独自コード・汎用文書・架空サンプルだけです。入力資料、生成した個別案件、認証情報は `work/` または `out/` に置き、コミットしません。API、ネットワーク、動的コード実行を製品に含めません。ライセンスは [MIT](LICENSE)。

着想の参考： [Jinba 記事](https://note.com/jinbaflow/n/nc8372b84e572)、[InsightEdge 記事](https://techblog.insightedge.jp/entry/powerpoint-auto-generation)、[Jinba 公開 repo](https://github.com/carnot-tech/consulting-pptx-skill)。第三者コード・記事本文は取り込んでいません。InsightEdge の非公開実装を取得したものでもありません。Jinba の現行 repo には PPTX 系の経路もあります。今回はユーザー指定により direct Python を選択しました。

実装 API の一次資料： [python-pptx](https://python-pptx.readthedocs.io/en/latest/)、[charts](https://python-pptx.readthedocs.io/en/latest/user/charts.html)、[tables](https://python-pptx.readthedocs.io/en/latest/user/table.html)。
