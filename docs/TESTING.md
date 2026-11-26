# Excel なしで VBA をテストする

このリポジトリの VBA は、**Excel も GUI も使わずに** Linux 上で実行・検証できる。

```sh
sudo apt-get install -y libreoffice-calc python3-uno   # 1回だけ
python3 tests/run.py            # コンパイル確認 + 全テスト
python3 tests/run.py --compile  # コンパイル確認だけ (速い)
python3 tests/run.py --bisect mColorUtil   # コンパイル失敗箇所の絞り込み
```

## 仕組み

LibreOffice Calc のドキュメントに `VBACompatibilityMode` を立てると、そのドキュメントの
Basic は **Excel の VBA オブジェクトモデルと Excel 定数** を持つ。

```python
doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, ())
doc.setPropertyValue("VBACompatibilityMode", True)
```

この状態で `Option VBASupport 1` を付けたモジュールを流し込むと、
`Application` / `ActiveSheet` / `Selection` / `ActiveCell` / `Range` / `Interior` /
`Font` / `Borders` / `xlEdgeTop` / `xlNone` などがそのまま動く。
`RGB()` の並びも VBA と同じ (BGR の Long) になる。

テストランナーは次の順に動く。

1. `src/*.bas` `*.cls` を読み、VBE 固有のヘッダ (`Attribute VB_*` など) を落として
   `Option VBASupport 1` (クラスは `Option ClassModule` も) を付ける。
   `Declare` 文の `PtrSafe` も取り除く (後述)。**それ以外のコード本文には手を入れない。**
2. ヘッドレスの `soffice` を起動し、UNO ソケット越しに接続する。
3. VBA 互換モードの Calc ドキュメントを開き、`src` と `tests/vba` の全モジュールを入れる。
4. **モジュールごとにコンパイル確認**する (後述)。
5. `tests/vba/Test_*.bas` の `Run` を順に呼び、結果を集計する。

## テストの書き方

`tests/vba/Test_<モジュール名>.bas` に `Public Sub Run()` を 1 つ書く。

```vb
Public Sub Run()
    Dim c As Long
    TestLib.IsTrue "16進を受け付ける", mColorUtil.TryParseHexColor("#ff6900", c)
    TestLib.Eq "往復する", mColorUtil.ColorLongToHex(c), "#FF6900"
End Sub
```

`TestLib.Eq` / `IsTrue` / `IsFalse` / `Failed` が使える。ファイルを置くだけで自動的に拾われる。
Excel オブジェクトを使う結合テストも同じ形式で書ける (`Test_mColor.bas` が例)。
各モジュールの `Run` は個別にエラートラップされるので、1 つが落ちても他の結果は残る。

## ハマりどころ (実測で確認済み)

### 1. LibreOffice Basic の予約語は VBA より広い

`step` と `base` を変数名・引数名に使うと、**エラーも出さずにモジュールのコンパイルが
失敗**する。しかも壊れたモジュールが 1 つあるとライブラリ全体が道連れになり、
無関係なモジュールの関数まで何も返さなくなる。

このため、ランナーは全モジュールに `ZZ_CompileProbe` を差し込み、
**1 モジュールずつ叩いて** コンパイル成功を確認する。壊れている場合は
`--bisect <モジュール名>` で原因の手続きまで絞り込める。

VBA では合法なので、Excel 上では動いてしまう。テスト可能に保つため
`src/` 側でこれらの識別子を避けている (`step` → `lvl`, `base` → `baseRng`)。

### 2. `Hidden=True` でドキュメントを開くと `ActiveCell` が落ちる

ViewShell が作られず `Object variable not set (91)` になる。`--headless` 起動なら
ウィンドウは出ないので、`Hidden` は指定しない。

### 3. 同じドキュメントにライブラリを入れ直さない

再インストールしても古いコンパイル結果が残り、テスト結果が汚染される。
1 回のテスト実行につき新しいドキュメントを開く。

### 4. コンパイルできない構文 (識別子のほか)

1 と同じく、エラーも出さずにモジュールのコンパイルが失敗する。

| 構文 | 対処 |
| --- | --- |
| `Declare PtrSafe` (`#If VBA7` などで分岐させても不可) | ランナーが読み込み時に `PtrSafe` を取り除く (`tests/harness/source.py`)。`Declare` と `LongPtr` はそのまま通る。Win32 API は Linux では呼べないので、`Declare` した関数を呼ぶ経路はテストしない |
| 戻り値が `Byte()` の `Function` | `src/` 側で `ByRef data() As Byte` の引数で受け渡す (`mClipboard.BuildCfHtml`) |

### 5. 結果の文字列に改行やタブを含めてよい

`TestLib` は結果を 1 件 1 行のタブ区切りで Python 側に渡す。改行・タブ・`\` は
`\n` `\t` `\\` などに置き換えて渡すので、`TestLib.Eq` に複数行の文字列をそのまま渡せる。

## 再現度の限界 — ここが緑でも Excel で緑とは限らない

実測して分かっている差分:

| 項目 | LibreOffice の挙動 |
| --- | --- |
| `Borders(...).LineStyle` の**読み出し** | 常に `1` (xlContinuous) を返す。線種の検証には使えない |
| 罫線の**書き込み** | `xlDash` や Union 範囲への一括適用が反映されないことがある |
| `Application.StatusBar` の**読み出し** | 常に `False`。メッセージの検証には使えない |
| `Shape` 系 (`BuildFreeform` など) | 未検証。`mPolygon` / `mLlp` のテストは未整備 |
| `step` / `base` などの識別子 | VBA では合法、LibreOffice では壊れる |
| `Range.DisplayFormat` | 存在しない (エラー 423)。`mHtml` は取れなければセル自身の書式を使う |
| `Range.HorizontalAlignment` の**読み出し** | `Nothing` を返す。横位置の検証には使えない (`mHtml` は標準として扱う) |
| `Font.Superscript` / `Font.Subscript` の**読み出し** | エラーになる (`mHtml` は False として扱う) |
| `Interior.Pattern` の**読み出し** | 塗りがあっても `xlPatternNone`。塗りの有無は `Interior.ColorIndex` (塗りなしは `xlColorIndexNone`) なら正しく取れる |
| `Font.Color` (自動) | `-1` を返す (Excel は `0`) |
| セル内で書式が混在するときの `Range.Font.Bold` など | `Null` にならない。`Characters(i, 1).Font` は1文字ずつ正しく読めるが、混在を前提にした経路 (`mHtml` のセル内の一部だけの書式) は通らない |
| `LenB(StrConv(s, vbFromUnicode))` | Excel と結果が違う。全角を2文字ぶんと数える列幅の揃え (`md`) は検証できない |

そのため、罫線の検証は VBA からではなく **Python 側の UNO API**
(`Cell.TopBorder.LineWidth` など) で行っている (`tests/run.py` の `check_tb_borders`)。

**このテストが担保するのは主に「そもそも動くか」**である。
構文エラー、タイプミス、未定義の変数・手続き、引数の食い違い、境界条件での実行時エラー —
AI が書いた VBA が動かない原因のほとんどはここで落ちる。
一方、Excel 固有の描画結果の細部は最終的に実機で確認する必要がある。

## 実機 (Excel) での最終確認

```sh
./build.py     # Windows 上で PERSONAL.XLSB を生成・配置
```

Excel を再起動し、ランチャーから各マクロを 1 回ずつ叩く。
LibreOffice で緑になっていれば、ここで出るのは「Excel 固有の見た目の差」だけになる。

## このテスト基盤が実際に見つけた不具合

- `mColorUtil.TryParseHueToken`: VBA の `And` は短絡評価しないため、
  `Do While i >= 1 And Mid$(token, i, 1) Like "[0-9]"` は数字だけのトークン
  (`bg 500`) や空文字で `Mid$(token, 0, 1)` を評価してしまい、実行時エラー 5 になっていた。
- 同関数の修正時に書いた `If Not Mid$(...) Like "[0-9]"` は、`Not` が `Like` より
  結合が強いため型不一致 (エラー 13) になる。括弧が要る。
