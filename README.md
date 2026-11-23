# personal-xlsx

Excel の個人用マクロブック (`PERSONAL.XLSB`) を、VBA ソースからビルド・デプロイするためのリポジトリ。
`src/*.bas` / `*.cls` を [pyopenvba](https://pypi.org/project/pyopenvba/) 経由で
`%APPDATA%\Microsoft\Excel\XLSTART\PERSONAL.XLSB` に書き込む。

## セットアップ

```sh
./build.py
```

- `src/` 内の `.bas` / `.cls` をすべて読み込み、`build/PERSONAL.xlsb` を新規生成。
- 生成後、`XLSTART` 配下の `PERSONAL.XLSB` を置き換える (Excel を終了させてから実行すること)。

### エンコーディング

- `src/*.bas` / `*.cls` は **UTF-8** で保存する (このリポジトリでの正)。通常のエディタでそのまま編集できる。
- VBAプロジェクト自体のコードページ (VBE 上での日本語表示に使われる) は `build.py` の `SRC_CODEPAGE` (cp932) で別途指定しており、ソースファイルのディスク上のエンコーディングとは独立している。
- ただし、コンパイル後のVBAプロジェクトへ埋め込む際は最終的に cp932 へエンコードされるため、**ソースコード中で使う文字は cp932(Shift-JIS) の範囲内に収めること** (絵文字や一部の拡張漢字・記号は不可)。範囲外の文字が含まれる場合、`build.py` はビルド時にエラーで検出する。
- VBE 上で直接コードを編集して `Export File...` した場合、そのファイルは Shift-JIS で書き出される。`build.py` は UTF-8 として読めない場合に自動で cp932 にフォールバックするため、そのまま `src/` に置いてもビルドは壊れない (ただし正はUTF-8なので、余裕があれば `iconv -f CP932 -t UTF-8` 等で正規化しておくと良い)。

## 機能一覧

マクロは、スペース区切りで「マクロ名 引数...」を入力して実行するランチャー ([参考](https://zenn.dev/sinotca/articles/6bf254d056af28)) から呼び出す想定 (例: `bg amber500`, `tb 2`)。多くはステータスバーに結果・エラーを表示する。

| モジュール | マクロ | 概要 |
| --- | --- | --- |
| `mColor` | `bg` | セル/図形の**塗り色**を取得・コピー・貼り付け・設定 |
| `mColor` | `fg` | セル/図形の**文字色**を取得・コピー・貼り付け・設定 |
| `mColor` | `bo` | 図形の**枠線色**を取得・コピー・貼り付け・設定 (セル罫線は対象外) |
| `mLlp` | `llp` | 数式で参照範囲を保持する「軽量リンク画像」の挿入・更新・移行・整理 |
| `mMisc` | `mg` | 選択列幅・行高をそろえて Excel 方眼紙を作成 |
| `mMisc` | `mt` | 選択範囲に縞模様の罫線 (テーブル風の書式) を適用 |
| `mUtil` | `tb` | 選択範囲にヘッダー行区切り・列区切りの罫線を適用 |
| `mBorder` | `db` | 選択範囲の各セルから指定辺の罫線を消す |
| `mView` | `fmt` | 全シートの表示位置を A1・ズーム 100% にリセット |
| `mView` | `pbp` | 全シートを改ページプレビュー表示にし、目盛線を非表示に |
| `mPolygon` | `polygon` | 数字セルを番号昇順に結んだ多角形の図形を作成 |
| `mPolygon` | `box` | 選択範囲に外接する矩形の図形を作成 |
| `mPalette` | `palette` | Tailwind配色の全一覧をシートに表示するガイド |

### `bg` / `fg` / `bo` — 色のコピー&ペースト (`mColor.bas`)

セル選択時はセルへ、図形選択時は図形へ自動的にスコープが切り替わる。

```
bg              現在の色を #rrggbb 形式でステータスバーに表示
bg copy | c     色をコピー用バッファへ保存
bg paste | p    バッファの色を選択範囲に適用
bg set COLOR    色を直接指定して設定
bg COLOR        set の短縮形
```

`fg` は文字色、`bo` は図形の枠線色 (セルの罫線は対象外) に対して同様に動作する。コピー用バッファは `bg` / `fg` / `bo` でそれぞれ独立している。

**COLOR の指定形式** (`mColorUtil.bas`)

- `#rrggbb` / `rrggbb` — 16進カラーコード (例: `#ffcc00`)
- `HUE[STEP]` — [Tailwind CSS v4](https://tailwindcss.com/docs/colors) の配色名 + 段階。`STEP` は `50,100,200,300,400,500,600,700,800,900,950`。`50` が最も明るく `950` が最も暗い (例: `blue50`〜`blue950`)。省略時は既定段階 `500`
- 使用可能な `HUE`: `red, orange, amber, yellow, lime, green, emerald, teal, cyan, sky, blue, indigo, violet, purple, fuchsia, pink, rose, slate, gray, zinc, neutral, stone, black, white`
- 色の一覧は `palette` マクロを実行するとシートで確認できる (下記参照)

### `llp` — Lightweight Linked Picture (`mLlp.bas`)

リンク画像を「名前定義 + `IF(LLP_LINK, 参照式, "")`」で管理し、参照範囲の追跡・再リンクを容易にする仕組み。

```
llp insert   | i   リンク画像を挿入 (事前に参照範囲をコピーしておく)
llp update   | u   すべての LLP 画像の描画を更新
llp migrate  | m   既存のリンク画像をすべて LLP 管理下に置き換え
llp clean    | c   使われなくなった LLP 関連の名前定義を削除
llp goto     | g   選択中の LLP 画像から参照範囲へジャンプ
llp localize | l   LLP 参照式から外部ブック名の修飾を削除
```

### `mg` — Excel 方眼紙作成 (`mMisc.bas`)

```
mg [SizePt]   全列・全行を SizePt (既定 20pt) 四方のマス目に揃える
```

### `fmt` / `pbp` — 表示切替 (`mView.bas`)

```
fmt           全シートを A1 選択・ズーム 100% にし、先頭シートをアクティブ化
pbp           全シートを改ページプレビュー表示にし、目盛線を非表示にする
```

### `mt` / `tb` — 表の罫線整形 (`mMisc.bas` / `mUtil.bas`)

選択範囲に対し、ヘッダー下端の実線、データ行の破線区切り、データ列の実線区切り、外枠の実線を自動適用する。空欄行・空欄列は区切り線の対象から除外される。

```
mt [headerRows]   ヘッダー行数を指定 (既定 1)。偶奇グループごとに罫線を一括適用する版
tb headerRows     ヘッダー行数を指定 (必須)。単純な区切り検出版
```

### `db` — 罫線を消す (`mBorder.bas`)

選択範囲の各セルについて、指定した辺の罫線を個別に消す (選択範囲全体の外周ではなく、選択した全セルそれぞれの指定辺が対象)。

```
db SIDES   SIDES で指定した辺の罫線を選択範囲内の各セルから消す
```

- `SIDES` は `oprl` の部分文字列 (順不同・重複可)。Excel リボンの罫線メニュー ([Alt] H B の後のアクセスキー) に合わせて `o` = 下辺、`p` = 上辺、`r` = 右辺、`l` = 左辺。
- 例: `db o` (下辺のみ)、`db pl` (上辺・左辺)、`db oprl` (全4辺)。
- `oprl` 以外の文字は無視する。有効な文字が1つもない場合は何もしない。

### `polygon` / `box` — 選択範囲から図形を作成 (`mPolygon.bas`)

いずれも線幅 0.5pt・線色は黒・塗りつぶしは白の図形を作成する。

```
polygon   選択範囲内の数値セルを昇順に結んだ多角形 (フリーフォーム図形) を作成
box       選択範囲のセル境界に外接する矩形を作成
```

`polygon`:
1. セルに数字 (文字列としての数字でも可) を書き込む。
2. 範囲選択して `polygon` を実行する。
3. 番号の昇順でセル左上の頂点を結んだ多角形を作成する。

- 数値以外 (空欄・非数値文字列) のセルは無視する。
- 対象セルが3個未満の場合、または複数領域選択の場合は何もしない。

`box`:
- 選択範囲に複数領域が含まれる場合は何もしない。

### `palette` — 配色ガイド (`mPalette.bas`)

`HUE[STEP]` トークンの一覧を、実際の色で塗ったセルとして PERSONAL.XLSB 内の `Palette` シートに表示する (既に存在すれば作り直す)。セルには `amber500` のようなトークン文字列がそのまま入っているので、`bg copy` (またはショートカット `bg c`) でセルの色をコピーし、対象セルで `bg paste` (`bg p`) すれば、色名を覚えていなくてもそのまま貼り付けられる。

```
palette   配色一覧シートを表示 (22色相 × 11段階)
```

- データは Tailwind CSS v4 (MIT) が公開する `oklch()` 値を、[coloraide](https://facelessuser.github.io/coloraide/) (`minde-chroma` 法 = CSS Color 4相当のガマットマッピング) で sRGB に変換した固定値。VBA側での計算は行わない。

## リポジトリ構成

```
build.py       ビルド & デプロイスクリプト
src/           VBA ソース (.bas / .cls、UTF-8 / CRLF。文字種は cp932 の範囲内)
  cBooster.cls   画面更新・イベント抑制を自動 ON/OFF するヘルパークラス
  mColor.bas     bg / fg / bo コマンド
  mColorUtil.bas 色パースユーティリティ (HUE[STEP] トークンの解決)
  mPalette.bas   Tailwind CSS 配色テーブル・palette コマンド
  mLlp.bas       llp コマンド
  mMisc.bas      mg / mt コマンド
  mUtil.bas      tb コマンド
  mBorder.bas    db コマンド
  mView.bas      fmt / pbp コマンド
  mPolygon.bas   polygon / box コマンド
build/         ビルド成果物 (gitignore 対象)
```
