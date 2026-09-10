Attribute VB_Name = "Test_mColorUtil"
Option Explicit
' mColorUtil の色パースは Excel オブジェクトに一切依存しないため、
' ここは純粋な単体テストになる。

Public Sub Run()
    Dim c As Long
    Dim ok As Boolean

    ' ---- TryParseHexColor ----
    ok = mColorUtil.TryParseHexColor("#ff6900", c)
    TestLib.IsTrue "TryParseHexColor: '#ff6900' を受け付ける", ok
    TestLib.Eq "TryParseHexColor: VBA の RGB 並び (BGR) で Long 化する", c, 27135
    TestLib.Eq "ColorLongToHex: '#ff6900' に往復する", mColorUtil.ColorLongToHex(c), "#FF6900"

    ok = mColorUtil.TryParseHexColor("ff6900", c)
    TestLib.IsTrue "TryParseHexColor: 先頭の '#' は省略できる", ok

    ok = mColorUtil.TryParseHexColor("FF6900", c)
    TestLib.IsTrue "TryParseHexColor: 大文字の16進も受け付ける", ok
    TestLib.Eq "ColorLongToHex: 出力は Hex$ 由来の大文字", mColorUtil.ColorLongToHex(c), "#FF6900"

    TestLib.IsFalse "TryParseHexColor: 非16進を弾く", mColorUtil.TryParseHexColor("#zzzzzz", c)
    TestLib.IsFalse "TryParseHexColor: 3桁省略記法は弾く", mColorUtil.TryParseHexColor("#fff", c)
    TestLib.IsFalse "TryParseHexColor: 7桁は弾く", mColorUtil.TryParseHexColor("#ff69000", c)
    TestLib.IsFalse "TryParseHexColor: 空文字を弾く", mColorUtil.TryParseHexColor("", c)

    TestLib.Eq "ColorLongToHex: 黒", mColorUtil.ColorLongToHex(0), "#000000"
    TestLib.Eq "ColorLongToHex: 白", mColorUtil.ColorLongToHex(RGB(255, 255, 255)), "#FFFFFF"

    ' ---- TryParseHueToken ----
    ok = mColorUtil.TryParseHueToken("amber500", c)
    TestLib.IsTrue "TryParseHueToken: 'amber500' を解決する", ok
    TestLib.Eq "TryParseHueToken: amber500 = #FD9A00", mColorUtil.ColorLongToHex(c), "#FD9A00"

    ok = mColorUtil.TryParseHueToken("blue", c)
    TestLib.IsTrue "TryParseHueToken: STEP 省略を受け付ける", ok
    TestLib.Eq "TryParseHueToken: STEP 省略時は 500", _
               mColorUtil.ColorLongToHex(c), UCase$("#" & mPalette.LookupHex("blue", "500"))

    TestLib.IsTrue "TryParseHueToken: 大文字小文字を区別しない", mColorUtil.TryParseHueToken("BLUE500", c)
    TestLib.IsTrue "TryParseHueToken: 前後の空白を無視する", mColorUtil.TryParseHueToken("  amber500  ", c)

    ok = mColorUtil.TryParseHueToken("black", c)
    TestLib.Eq "TryParseHueToken: black", mColorUtil.ColorLongToHex(c), "#000000"
    ok = mColorUtil.TryParseHueToken("white", c)
    TestLib.Eq "TryParseHueToken: white", mColorUtil.ColorLongToHex(c), "#FFFFFF"

    TestLib.IsFalse "TryParseHueToken: 未知の色相は弾く", mColorUtil.TryParseHueToken("nosuchhue700", c)
    TestLib.IsFalse "TryParseHueToken: 未定義の段階は弾く", mColorUtil.TryParseHueToken("blue123", c)
    TestLib.IsFalse "TryParseHueToken: 空文字を弾く", mColorUtil.TryParseHueToken("", c)
    ' 数字だけのトークンは、段階を切り出すループが i = 0 まで進む境界ケース。
    ' VBA の And は短絡評価しないため、素朴に書くと Mid$(token, 0, 1) で実行時エラー5になる。
    TestLib.IsFalse "TryParseHueToken: 数字だけのトークンを弾く (エラーにしない)", _
                    mColorUtil.TryParseHueToken("500", c)

    ' 50 が最も明るく 950 が最も暗い、という README の約束を確認する。
    Dim light As Long, dark As Long
    ok = mColorUtil.TryParseHueToken("blue50", light)
    ok = mColorUtil.TryParseHueToken("blue950", dark)
    TestLib.IsTrue "配色: blue50 は blue950 より明るい", Luma(light) > Luma(dark)
End Sub

Private Function Luma(ByVal c As Long) As Double
    Dim r As Long, g As Long, b As Long
    r = c And &HFF&
    g = (c \ &H100&) And &HFF&
    b = (c \ &H10000) And &HFF&
    Luma = 0.299 * r + 0.587 * g + 0.114 * b
End Function
