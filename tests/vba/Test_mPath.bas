Attribute VB_Name = "Test_mPath"
Option Explicit
'===============================================================================
' path (ブックのパス・ファイル名などのコピー) のテスト。
' クリップボードは偽物 (tests/fakes/mClipboard.bas) に差し替わっている。
'===============================================================================

Public Sub Run()
    Dim wb As Object
    Dim ws As Object

    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets.Add
    ws.Name = "設計"

    ' ---- 保存していないブック ----
    TestLib.Eq "前提: 保存していないブックはパスが空", wb.Path, ""

    mClipboard.Reset
    mPath.path
    TestLib.Eq "保存前: path は何もコピーしない", mClipboard.CallCount(), 0

    mClipboard.Reset
    mPath.path "dir"
    TestLib.Eq "保存前: path dir は何もコピーしない", mClipboard.CallCount(), 0

    mClipboard.Reset
    mPath.path "name"
    TestLib.Eq "保存前: path name は仮の名前をコピーする", mClipboard.LastText(), wb.Name

    ' ---- path cell ----
    ws.Range("B3:F10").Select
    mClipboard.Reset
    mPath.path "cell"
    TestLib.Eq "path cell: ブック名 シート名!範囲", mClipboard.LastText(), wb.Name & " 設計!B3:F10"

    ws.Range("A1").Select
    mClipboard.Reset
    mPath.path "CELL"
    TestLib.Eq "path cell: 1セル、種類の大文字・小文字は区別しない", mClipboard.LastText(), wb.Name & " 設計!A1"

    TestLib.Eq "CellLocation: 複数の範囲はつなげる", _
        mPath.CellLocation(wb, Union(ws.Range("A1:B2"), ws.Range("D4"))), wb.Name & " 設計!A1:B2,D4"

    ' ---- 不明な種類 ----
    mClipboard.Reset
    mPath.path "nosuch"
    TestLib.Eq "不明な種類では何もコピーしない", mClipboard.CallCount(), 0

    ' ---- 保存済みのブック ----
    Dim dir As String, file As String
    dir = Environ$("TMPDIR")
    If Len(dir) = 0 Then dir = "/tmp"
    file = dir & "/path_test_" & CStr(CLng(Timer)) & ".xlsx"
    On Error Resume Next
    wb.SaveAs file
    If Err.Number <> 0 Then
        TestLib.Failed "前提: ブックを保存できる", "E" & Err.Number & ": " & Err.Description
        Exit Sub
    End If
    On Error GoTo 0

    ' LibreOffice の FullName はフォルダを含まないファイル名だけを返すので、フルパスに
    ' なっていることは確かめられない (Excel の FullName をそのまま使っていることだけを見る)
    mClipboard.Reset
    mPath.path
    TestLib.Eq "保存後: path は FullName", mClipboard.LastText(), wb.FullName
    mClipboard.Reset
    mPath.path "dir"
    TestLib.Eq "保存後: path dir はフォルダのパス", mClipboard.LastText(), wb.Path
    mClipboard.Reset
    mPath.path "name"
    TestLib.Eq "保存後: path name はファイル名", mClipboard.LastText(), wb.Name

    ' 保存したファイルを消す (開いたままで消せなくても、テストの結果には影響しない)
    On Error Resume Next
    Kill file
    On Error GoTo 0
End Sub
