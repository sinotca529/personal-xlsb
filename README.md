# personal-xlsx

Excel の個人用マクロブック (`PERSONAL.XLSB`) を、VBA ソースからビルド・デプロイするためのリポジトリ。
`src/` のソースを [pyopenvba](https://pypi.org/project/pyopenvba/) で
`%APPDATA%\Microsoft\Excel\XLSTART\PERSONAL.XLSB` に書き込む。

## ビルド

```sh
uv run build.py
```

- `build/PERSONAL.xlsb` を新規生成し、`XLSTART` の `PERSONAL.XLSB` を置き換える。Excel を終了させてから実行すること。
- Windows / WSL のどちらからでも実行できる。

## 使い方

`Ctrl+L` でランチャー ([参考](https://zenn.dev/sinotca/articles/6bf254d056af28)) を開き、「マクロ名 引数...」を入力して `Enter` で実行する (例: `bg amber500`, `tb 2`)。

- 入力中は補完候補と説明を表示する。`Tab` / `Shift+Tab` / 上下キーで選択、`Esc` で閉じる。
- 引数は空白区切り。`"..."` で囲むと空白を含められる (例: `fg red "a b" c`)。
- `!!` で直前のコマンドを繰り返す。

## コマンド

各コマンドの詳しい使い方は、ソースの先頭コメント (NAME / SYNOPSIS / DESCRIPTION) に書いてある。

| コマンド                            | 概要                                                                                              | ソース                               |
| ----------------------------------- | ------------------------------------------------------------------------------------------------- | ------------------------------------ |
| `bg` / `fg` / `bo`                  | 塗り色 / 文字色 / 図形の枠線色を取得・コピー・貼り付け・設定                                      | [`mColor.bas`](src/mColor.bas)       |
| `ln`                                | 図形の線の書式 (太さ・種類・色・矢印) を取得・設定・コピー・貼り付け                              | [`mLine.bas`](src/mLine.bas)         |
| `llp`                               | 軽量リンク画像の挿入・更新・移行・整理 ([参考](https://zenn.dev/sinotca/articles/18638886a004d4)) | [`mLlp.bas`](src/mLlp.bas)           |
| `mg` / `mt`                         | Excel 方眼紙の作成 / テーブル風の罫線                                                             | [`mMisc.bas`](src/mMisc.bas)         |
| `tb`                                | 表の罫線                                                                                          | [`mUtil.bas`](src/mUtil.bas)         |
| `db`                                | 選択範囲の各セルから指定辺の罫線を消す                                                            | [`mBorder.bas`](src/mBorder.bas)     |
| `ditto`                             | すぐ上のセルと同じ値のセルの文字色をグレーにする (同じ値を並べた表を見やすく)                     | [`mDitto.bas`](src/mDitto.bas)       |
| `fmt` / `pbp`                       | 全シートの表示をリセット / 改ページプレビューに                                                   | [`mView.bas`](src/mView.bas)         |
| `polygon` / `box`                   | 数字セルを結んだ多角形 / 選択範囲を囲む矩形を作成                                                 | [`mPolygon.bas`](src/mPolygon.bas)   |
| `path`                              | ブックのフルパス・ファイル名・フォルダ・選択範囲の場所をコピー                                    | [`mPath.bas`](src/mPath.bas)         |
| `palette`                           | Tailwind 配色の一覧シートを表示                                                                   | [`mPalette.bas`](src/mPalette.bas)   |
| `find` / `fn` / `fp` / `fq`         | セルと図形内のテキストをブック横断で検索し、`n` / `N` で順に移動                                  | [`mFind.bas`](src/mFind.bas)         |
| `ff`                                | シート名のあいまい検索でシートへ移動 (候補を良い順に表示)                                         | [`mSheet.bas`](src/mSheet.bas)       |
| `s`                                 | 選択範囲のセルと図形内のテキストを正規表現で置換 (`s/PAT/REP/g`)                                  | [`mSubst.bas`](src/mSubst.bas)       |
| `lint`                              | 数式のエラーと外部ブックへの参照を調べ、`n` / `N` で順に移動                                      | [`mLint.bas`](src/mLint.bas)         |
| `md`                                | 選択範囲を Markdown の表としてクリップボードにコピー                                              | [`mMarkdown.bas`](src/mMarkdown.bas) |
| `html`                              | 選択範囲を罫線・書式つきの HTML の表としてクリップボードにコピー                                  | [`mHtml.bas`](src/mHtml.bas)         |
| `sort` / `uniq` / `hsort` / `huniq` | 選択範囲の行 (h: 列) を並べ替え / 重複を取り除く                                                  | [`mSort.bas`](src/mSort.bas)         |

## テスト

Excel を使わずに、ヘッドレスの LibreOffice Calc 上で `src/` の VBA をコンパイル・実行して検証できる。push / PR では GitHub Actions で自動実行する。

```sh
sudo apt-get install -y libreoffice-calc python3-uno   # 1回だけ
python3 tests/run.py                  # コンパイル確認 + 全テスト
python3 tests/run.py --compile        # コンパイル確認だけ
python3 tests/run.py --bisect mHtml   # コンパイルに失敗する箇所を絞り込む
```

LibreOffice の VBA 互換は完全ではない (罫線の線種を読み戻せない、`step` / `base` を識別子に使えない、など)。仕組み・テストの書き方・担保できる範囲は [docs/TESTING.md](docs/TESTING.md) を参照。

## 開発メモ

- **文字コード**: `src/` は UTF-8 / CRLF で保存する。VBA プロジェクトには cp932 で埋め込まれるため、使える文字は cp932 の範囲内に限られる (範囲外はビルドエラー)。VBE から書き出した cp932 のファイルもそのまま読める。
- **補完表**: ランチャーの補完候補は、`build.py` が各モジュールの先頭コメントの `COMPLETIONS` 節から生成して埋め込む。書式は `build.py` を参照。コマンドを追加したら、同じモジュールの `COMPLETIONS` 節にも書く。候補が入力のたびに変わるもの (シート名など) は、候補の代わりに `@モジュール.関数` と書き、その関数が候補を作る (例: `mSheet.bas` の `ff`)。
- **フォーム**: `Launcher` フォームのコードは `src/Launcher.frm`、コントロールの配置は `build.py` の `FORMS` で定義する (`.frx` は使わない)。
