Attribute VB_Name = "Test_mPalette"
Option Explicit
' mPalette.LookupHex は固定の配色テーブル。Excel には依存しない。

Public Sub Run()
    TestLib.Eq "LookupHex: red-500", mPalette.LookupHex("red", "500"), "fb2c36"
    TestLib.Eq "LookupHex: orange-50", mPalette.LookupHex("orange", "50"), "fff7ed"
    TestLib.Eq "LookupHex: 未知の色相は空文字", Len(mPalette.LookupHex("nope", "500")), 0
    TestLib.Eq "LookupHex: 未知の段階は空文字", Len(mPalette.LookupHex("red", "123")), 0

    ' README が謳う 22 色相 x 11 段階がすべて引けること。
    Dim hues As Variant, steps As Variant
    Dim i As Long, j As Long
    Dim missing As Long, badLen As Long, dup As Long
    Dim seen As String
    hues = Array("red", "orange", "amber", "yellow", "lime", "green", "emerald", "teal", _
                 "cyan", "sky", "blue", "indigo", "violet", "purple", "fuchsia", "pink", _
                 "rose", "slate", "gray", "zinc", "neutral", "stone")
    steps = Array("50", "100", "200", "300", "400", "500", "600", "700", "800", "900", "950")

    Dim hexStr As String
    For i = 0 To UBound(hues)
        For j = 0 To UBound(steps)
            hexStr = mPalette.LookupHex(CStr(hues(i)), CStr(steps(j)))
            If Len(hexStr) = 0 Then
                missing = missing + 1
            ElseIf Len(hexStr) <> 6 Then
                badLen = badLen + 1
            Else
                If InStr(seen, "|" & hexStr) > 0 Then dup = dup + 1
                seen = seen & "|" & hexStr
            End If
        Next j
        seen = ""   ' 重複判定は色相ごと (色相をまたぐ一致は Tailwind 側に実在する)
    Next i

    TestLib.Eq "配色表: 22色相 x 11段階がすべて存在する", missing, 0
    TestLib.Eq "配色表: すべて6桁の16進", badLen, 0
    TestLib.Eq "配色表: 同一色相の11段階に重複がない", dup, 0
    TestLib.Eq "配色表: 件数は 22 x 11", (UBound(hues) + 1) * (UBound(steps) + 1), 242
End Sub
