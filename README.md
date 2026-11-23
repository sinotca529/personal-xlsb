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
| `fmt` / `pbp`                       | 全シートの表示をリセット / 改ページプレビューに                                                   | [`mView.bas`](src/mView.bas)         |
| `polygon` / `box`                   | 数字セルを結んだ多角形 / 選択範囲を囲む矩形を作成                                                 | [`mPolygon.bas`](src/mPolygon.bas)   |
| `palette`                           | Tailwind 配色の一覧シートを表示                                                                   | [`mPalette.bas`](src/mPalette.bas)   |
| `find` / `fn` / `fp` / `fq`         | セルと図形内のテキストをブック横断で検索し、`n` / `N` で順に移動                                  | [`mFind.bas`](src/mFind.bas)         |
| `md`                                | 選択範囲を Markdown の表としてクリップボードにコピー                                              | [`mMarkdown.bas`](src/mMarkdown.bas) |
| `sort` / `uniq` / `hsort` / `huniq` | 選択範囲の行 (h: 列) を並べ替え / 重複を取り除く                                                  | [`mSort.bas`](src/mSort.bas)         |

## 開発メモ

- **文字コード**: `src/` は UTF-8 / CRLF で保存する。VBA プロジェクトには cp932 で埋め込まれるため、使える文字は cp932 の範囲内に限られる (範囲外はビルドエラー)。VBE から書き出した cp932 のファイルもそのまま読める。
- **補完表**: ランチャーの補完候補は、`build.py` が各モジュールの先頭コメントの `COMPLETIONS` 節から生成して埋め込む。書式は `build.py` を参照。コマンドを追加したら、同じモジュールの `COMPLETIONS` 節にも書く。
- **フォーム**: `Launcher` フォームのコードは `src/Launcher.frm`、コントロールの配置は `build.py` の `FORMS` で定義する (`.frx` は使わない)。
