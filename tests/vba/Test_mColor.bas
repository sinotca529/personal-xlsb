Attribute VB_Name = "Test_mColor"
Option Explicit
'===============================================================================
' bg / fg の結合テスト。
'
' Excel オブジェクト (Selection / ActiveCell / Interior / Font) を実際に触る
' コードパスを、LibreOffice の VBA 互換オブジェクトモデル上で動かして検証する。
'===============================================================================

Public Sub Run()
    Dim ws As Object
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Activate

    ' ---- bg COLOR (set の短縮形) ----
    ws.Range("B2").Select
    mColor.bg "amber500"
    TestLib.Eq "bg amber500: セルの塗りが amber500 になる", _
               mColorUtil.ColorLongToHex(CLng(ws.Range("B2").Interior.Color)), "#FD9A00"

    ws.Range("C3").Select
    mColor.bg "#ff0000"
    TestLib.Eq "bg #ff0000: 16進指定でセルの塗りが変わる", _
               mColorUtil.ColorLongToHex(CLng(ws.Range("C3").Interior.Color)), "#FF0000"

    ws.Range("C4").Select
    mColor.bg "set", "blue700"
    TestLib.Eq "bg set blue700: set 形式でも設定できる", _
               mColorUtil.ColorLongToHex(CLng(ws.Range("C4").Interior.Color)), _
               UCase$("#" & mPalette.LookupHex("blue", "700"))

    ' ---- fg COLOR ----
    ws.Range("D4").Select
    mColor.fg "blue700"
    TestLib.Eq "fg blue700: 文字色が変わる", _
               mColorUtil.ColorLongToHex(CLng(ws.Range("D4").Font.Color)), _
               UCase$("#" & mPalette.LookupHex("blue", "700"))

    ' ---- copy / paste ----
    ws.Range("B2").Select
    mColor.bg "copy"
    ws.Range("E5").Select
    mColor.bg "paste"
    TestLib.Eq "bg copy/paste: コピー元と同じ塗りが貼られる", _
               CLng(ws.Range("E5").Interior.Color), CLng(ws.Range("B2").Interior.Color)

    ws.Range("E6").Select
    mColor.bg "c"
    ws.Range("E7").Select
    mColor.bg "p"
    TestLib.Eq "bg c/p: 1文字の別名でも動く", _
               CLng(ws.Range("E7").Interior.Color), CLng(ws.Range("E6").Interior.Color)

    ' ---- 不正な色トークン ----
    ' 不正な入力でも例外で落ちず、対象セルの色を変えないこと。
    ws.Range("F8").Select
    ws.Range("F8").Interior.Color = RGB(1, 2, 3)
    mColor.bg "nosuchcolor999"
    TestLib.Eq "bg 不正トークン: 色を変更しない", _
               CLng(ws.Range("F8").Interior.Color), CLng(RGB(1, 2, 3))
End Sub
