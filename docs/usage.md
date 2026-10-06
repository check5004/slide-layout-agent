# セットアップと操作コマンド

[README に戻る](../README.md)

Codex に任せる場合は、README の依頼文から始めてください。このページは、別 PC で環境を用意する場合や、手動で実行したい場合の手順です。

## 別 PC で始める

Git と Python 3.11 以降を用意し、リポジトリを取得します。

```sh
git clone https://github.com/check5004/slide-layout-agent.git
cd slide-layout-agent
```

既に取得済みなら clone は不要です。以降はこのフォルダで実行します。

### 何が必要か

| やりたいこと | 必要なもの |
|---|---|
| 設計 JSON から PPTX を生成 | Python と `requirements.txt` のライブラリ。PowerPoint 本体は不要。 |
| 文章から AI に構成を考えてもらう | 上記に加えて Codex。CLI 単体には AI planner がない。 |
| 見た目の確認・手直し | PowerPoint など、PPTX を開けるアプリと日本語フォント。 |
| 付属スクリプトで PNG 出力・文字境界計測 | Windows と Microsoft PowerPoint。生成だけなら不要。 |

自動の画像書き出しは任意の補助操作です。完成資料として使う前には、利用するアプリで全ページの見た目を確認してください。

### Windows / PowerShell

通常の Python をインストールした環境で実行します。

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

仮想環境の有効化は不要です。以下のコマンド例の先頭にある `python` を、Windows では `.\.venv\Scripts\python.exe` に置き換えてください。

### macOS / Linux

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

ライブラリ取得にはネット接続が必要です。準備後の Python CLI は外部 API を呼びません。OS 別の自動セットアップスクリプトは同梱していません。

### フォントについて

追加フォントの導入は PPTX 生成の必須条件ではありません。CLI はフォント名を記録し、閲覧するアプリが実際の文字を描画します。フォント自体は PPTX に埋め込みません。

Windows では標準設定の `Meiryo`、macOS では端末に入っているヒラギノ系など、日本語対応の既存フォントを候補にできます。利用状況は [Microsoft の Meiryo 情報](https://learn.microsoft.com/en-us/typography/font-list/meiryo) と [Apple のフォント一覧](https://support.apple.com/en-la/127491) も参照してください。Linux の標準フォントは配布環境によって異なります。

変更する場合はテーマの `font_family` を指定します。例：`{"font_family": "Hiragino Sans"}`。自動で OS に合わせて切り替わる実装ではありません。指定フォントがない端末では代替フォントになり、改行・文字幅が変わることがあるため、開いたときに確認してください。

### clone だけで引き継がれるもの

コード、Codex 用スキル、設計 JSON、架空サンプル画像はリポジトリに含まれます。開発時の専用 Python、個人用パス、未追跡のローカル画像は実行条件に含めていません。

自分の資料が入る `work/` と生成先 `out/` は Git 管理対象外なので、別 PC へは自分で移してください。Python・Codex・PowerPoint・フォント自体は clone ではインストールされません。

[GitHub Actions](https://github.com/check5004/slide-layout-agent/actions) は、新規チェックアウトから Python 3.11 / 3.12 で依存導入・自動テスト・サンプル生成を行います。Codex の判断、PowerPoint 本体による描画、フォントの完全一致までは CI の検証対象ではありません。

## サンプルを動かす

同梱の架空サンプルから試せます。

```sh
python -m slide_agent validate examples/demo.plan.json --source examples/demo.source.json
python -m slide_agent render examples/demo.plan.json --source examples/demo.source.json --out out/demo.pptx
```

`out/demo.pptx` を開いてください。生成済みの [demo.pptx](../examples/demo.pptx?raw=true) も同梱しています。

同じ出力先でやり直すときだけ `--force` を付けます。手直し済みのファイルには別の出力名を使ってください。

## 自分の入力を読み込む

`work/` を作り、UTF-8 の文章と PNG/JPEG 画像を保存します。

### 文章から構成を考えてもらう

文章を段落で区切り、`work/input.txt` に保存します。

```sh
python -m slide_agent ingest work/input.txt --mode prose --out work/source.json
```

### スライドごとの文章を渡す

`work/slides.md` に `## 見出し` を書き、その下に各ページの本文を置きます。[入力例](../examples/slides.md) を参照してください。

```sh
python -m slide_agent ingest work/slides.md --mode slides --out work/source.json
```

上の2つは、入力に合わせてどちらかを選びます。

### 画像も添付する

読み込み時に `--image ID=PATH` を追加します。`photo` は構成案から画像を指定するための名前です。

```sh
python -m slide_agent ingest work/input.txt --mode prose --image photo=work/photo.png --out work/source.json
```

複数画像は `--image` を繰り返して指定します。画像は `source.json` と同じフォルダか、その下に置いてください。

`source.json` には原文・参照 ID・画像情報が保存されます。構成案を作った後で原文や画像を変更した場合は、Codex に原文からの再確認を依頼してください。

## Codex に構成案を作ってもらう

CLI は文章から構成を考える機能を持ちません。ここは Codex の担当です。

> AGENTS.md と .agents/skills/slide-plan/SKILL.md を読み、work/source.json から work/plan.json を作ってください。
>
> 原文の数値・出典・留保を保ち、不明点や省略は示してください。
>
> 検証して、構成確認用の work/review.md も作ってください。

参考： [未構造文章の構成案](../examples/prose.plan.json) · [スライド別文章の構成案](../examples/slides.plan.json)

構成案ができたら、検証と確認用ファイルの作成ができます。

```sh
python -m slide_agent validate work/plan.json --source work/source.json
python -m slide_agent review work/plan.json --source work/source.json --out work/review.md
```

ページの順序、主張、数値、留保、省略、画像の扱いを確認します。修正は Codex に依頼するか、`plan.json` を編集して再検証します。

## 生成して手直しする

```sh
python -m slide_agent render work/plan.json --source work/source.json --out work/result.pptx
```

- `work/result.pptx`：編集可能な PowerPoint。
- `work/result.validation.json`：原文参照・文字量・PPTX 内部構造などの検証結果。

構造検証に合格しても、実際の見た目は別に確認が必要です。PowerPoint などで全ページを開き、文字の欠け、重なり、表・グラフの読みやすさを確認してください。

PowerPoint での手直しは `plan.json` に自動では戻りません。別名で保存し、再生成する場合は構成案の側にも変更を反映してください。

## 必要なときだけ使う操作

### 長文・行数超過で止まったとき

まず Codex にページの分割を依頼してください。`split` は、箇条書きの項目と表の行を、順序と原文参照を保って複数ページに分けます。

```sh
python -m slide_agent split work/plan.json --source work/source.json --out work/split.plan.json
```

成功したら、以降は `work/split.plan.json` を使います。1項目自体が長すぎる場合や他のレイアウトは、構成案を直す必要があります。

### 省略・上書き・テーマ

| オプション | 使うとき |
|---|---|
| `--accept-omissions` | 原文の省略を確認し、受け入れたうえで生成する。`render` 専用。 |
| `--force` | 指定した既存出力を上書きする。 |
| `--theme PATH` | `validate` / `review` / `split` / `render` で同じテーマを使う。 |

標準では副題を置きません。短い要旨を許可する設定例は [theme.with-lead.json](../examples/theme.with-lead.json) です。[テーマの項目](design.md#見た目のルール) も参照してください。

### 原文ハッシュと JSON スキーマ

構成案を作る Codex や、機能を拡張する開発者向けです。

```sh
python -m slide_agent fingerprint work/source.json
python -m slide_agent schema --out out/plan.schema.json
```

[JSON の仕様](contract.md) · [同梱スキーマ](plan.schema.json)

### PowerPoint で PNG に書き出す

Windows と Microsoft PowerPoint がある場合に使えます。保存先は空のフォルダを指定します。

```powershell
powershell -NoProfile -File scripts/render_powerpoint.ps1 -Pptx work/result.pptx -OutputDir out/rendered
```

PNG と文字境界の計測結果 `out/rendered/render-report.json` が出力されます。表やグラフ内部の文字は自動測定の対象外なので、画像でも確認してください。

### 開発時のテストとデモ再生成

```sh
python -m unittest discover -s tests -v
python scripts/make_demo.py
```

`make_demo.py` は固定の架空データを再作成します。AI が構成を考える処理ではありません。既存のデモ JSON と画像を上書きするため、通常の利用では実行不要です。
