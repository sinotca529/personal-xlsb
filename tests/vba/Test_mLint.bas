Attribute VB_Name = "Test_mLint"
Option Explicit
'===============================================================================
' lint の外部リンクの調査のテスト。
'
' LibreOffice では次のものが使えないので、ここでは確かめない
' (docs/TESTING.md の「再現度の限界」)。
'   - VBScript.RegExp (外部参照の判定に使う)。外部参照の数式・入力規則・グラフの系列・
'     リンクされた画像の判定は確かめられない
'   - 入力規則のあるセルの検索 (SpecialCells(xlCellTypeAllValidation))
'   - 図形の OnAction / TopLeftCell / HasChart、グラフシート (Workbook.Charts)
' そのため、正規表現を使わない経路 (ハイパーリンク) と、図形やグラフがあっても
' 調査が止まらないことを確かめる。数式のセルを置かなければ、正規表現は使われない。
'===============================================================================

Public Sub Run()
    Dim ws As Object
    Dim hits As Collection

    ' ---- IsFileLink ----
    TestLib.IsTrue "ファイルのパスはファイルへのリンク", mLint.IsFileLink("C:\docs\spec.xlsx")
    TestLib.IsTrue "相対パスもファイルへのリンク", mLint.IsFileLink("..\spec.xlsx")
    TestLib.IsTrue "file: の URL もファイルへのリンク", mLint.IsFileLink("file:///C:/docs/spec.xlsx")
    TestLib.IsFalse "空 (ブック内のリンク) は対象外", mLint.IsFileLink("")
    TestLib.IsFalse "https は対象外", mLint.IsFileLink("https://example.com/")
    TestLib.IsFalse "http は大文字でも対象外", mLint.IsFileLink("HTTP://example.com/")
    TestLib.IsFalse "mailto は対象外", mLint.IsFileLink("mailto:someone@example.com")

    ' ---- ScanSheet: ハイパーリンクを位置の順に ----
    Set ws = ActiveWorkbook.Worksheets.Add
    ' 足す順番を位置の順 (C3 → A5) と逆にして、並べ替えを確かめる
    ws.Hyperlinks.Add ws.Range("A5"), "..\other.xlsx"
    ws.Hyperlinks.Add ws.Range("B2"), "https://example.com/"
    ws.Hyperlinks.Add ws.Range("C3"), "C:\docs\spec.xlsx"
    ws.Range("D4").Value = "問題なし"

    Set hits = mLint.ScanSheet(ws, Nothing, New Collection)
    TestLib.Eq "ファイルへのハイパーリンクの 2 件 (Web のリンクは除く)", hits.Count, 2
    If hits.Count = 2 Then
        TestLib.Eq "位置の順: 1 件目は C3", hits(1)(mHitList.H_TARGET).Address(False, False), "C3"
        TestLib.Eq "1 件目の説明はリンク先のファイル", hits(1)(mHitList.H_NOTE), "リンク先のファイル: C:\docs\spec.xlsx"
        TestLib.Eq "位置の順: 2 件目は A5", hits(2)(mHitList.H_TARGET).Address(False, False), "A5"
    End If

    ' ---- 図形やグラフがあっても止まらない ----
    ws.Shapes.AddShape 1, 100, 100, 50, 30
    ws.ChartObjects.Add 200, 200, 200, 100
    ' (On Error Resume Next の後の Err.Number は、中で握りつぶしたエラーが LibreOffice では
    '  残っていることがあるので使わない。例外が飛んできたかはエラー処理で受けて調べる)
    Dim raised As String
    Set hits = ScanOrError(ws, raised)
    TestLib.Eq "図形やグラフがあっても例外にならない", raised, ""
    If Not hits Is Nothing Then TestLib.Eq "図形やグラフがあっても結果は同じ", hits.Count, 2

    ' ---- 問題のないシート ----
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Range("A1").Value = 1
    ws.Hyperlinks.Add ws.Range("A2"), "https://example.com/"
    TestLib.Eq "問題のないシートは 0 件", mLint.ScanSheet(ws, Nothing, New Collection).Count, 0
End Sub

' ScanSheet を呼び、例外が飛んできたら raised にその内容を入れて Nothing を返す
Private Function ScanOrError(ByVal ws As Object, ByRef raised As String) As Collection
    raised = ""
    On Error GoTo FAILED
    Set ScanOrError = mLint.ScanSheet(ws, Nothing, New Collection)
    Exit Function
FAILED:
    raised = "E" & Err.Number & " " & Err.Description
End Function
