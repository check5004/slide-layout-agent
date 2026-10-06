# 実行環境のFAQ

## Windowsでcatalogの文字化け・UnicodeEncodeErrorが出る

AIや別プログラムへ渡すときは、シェルの`>`ではなく直接ファイル出力を使います。

```sh
python -B -m slide_agent catalog --layout executive_summary --out out/executive-summary.json
python -B -m slide_agent catalog --out out/catalog.json
```

`--out`は常にUTF-8で保存し、既存ファイルは`--force`を指定しない限り上書きしません。標準出力は対話的なUTF-8端末なら日本語をそのまま、cp932等またはpipe/redirectionではASCIIのJSON escapeで出します。どちらもJSONとして同じ内容です。

古いPowerShellはnative commandのstdoutを独自の文字コードで再解釈したり、リダイレクト時に再保存したりします。`-X utf8`だけでその後の変換まで制御できるとは限りません。`--out`ならこの経路を使いません。Python側で読む場合は`Path(...).read_text(encoding='utf-8')`を使います。

## Office描画だけがCOMエラーになる

PPTX生成、schema確認、自動テストはOfficeを呼びません。`render_catalog_powerpoint.ps1`は任意の表示QAです。COMエラー`0x80070520`などは、sandboxや非対話セッションからPowerPointのユーザーセッションへ接続できない環境で起こる場合があります。

既存のPowerPointが使える同一ユーザーの通常Windowsセッションで、このQAスクリプトを実行してください。sandbox外実行が必要な場合は実行環境の承認手順に従います。OS/Dockerのセキュリティ設定や認証を変更して解決しようとしません。描画できない環境では、生成・構造検査までの結果と未描画であることを区別して報告します。
