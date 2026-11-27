Attribute VB_Name = "Test_mSheet"
Option Explicit
'===============================================================================
' ff (シート名のあいまい検索) のテスト。
' ランチャーの候補一覧 (UserForm) は LibreOffice では動かないので、候補を作る関数
' (SheetCandidates) と、移動する ff 本体を確かめる。
'===============================================================================

Public Sub Run()
    Dim wb As Object
    Dim c As Variant

    ' ---- FuzzyScore ----
    TestLib.IsTrue "文字を飛ばして一致する", mSheet.FuzzyScore("設計書", "設書") >= 0
    TestLib.Eq "順番が違えば一致しない", mSheet.FuzzyScore("設計書", "書設"), -1
    TestLib.Eq "含まない文字があれば一致しない", mSheet.FuzzyScore("設計書", "設図"), -1
    TestLib.IsTrue "大文字・小文字を区別しない", mSheet.FuzzyScore("DB設計", "db") >= 0
    TestLib.IsTrue "全角・半角を区別しない", mSheet.FuzzyScore("ＤＢ設計", "db") >= 0
    TestLib.IsTrue "検索語の空白は無視する", mSheet.FuzzyScore("設計書", "設計 書") >= 0
    TestLib.IsTrue "連続した一致のほうが高い", _
        mSheet.FuzzyScore("設計書", "設計") > mSheet.FuzzyScore("設a計", "設計")
    TestLib.IsTrue "先頭での一致のほうが高い", _
        mSheet.FuzzyScore("画面設計", "画面") > mSheet.FuzzyScore("旧画面設計", "画面")
    TestLib.IsTrue "区切りの直後の一致のほうが高い", _
        mSheet.FuzzyScore("旧_画面", "画面") > mSheet.FuzzyScore("旧x画面", "画面")
    TestLib.IsTrue "完全一致は最高点", _
        mSheet.FuzzyScore("設計", "設計") > mSheet.FuzzyScore("設計2", "設計")
    TestLib.Eq "空の検索語はすべてに一致", mSheet.FuzzyScore("何でも", ""), 0

    ' ---- 並び順と候補 ----
    Set wb = ActiveWorkbook
    AddSheet wb, "旧画面設計"
    AddSheet wb, "画面設計 一覧"
    AddSheet wb, "画面設計"
    AddSheet wb, "DB設計"
    AddSheet(wb, "画面設計_非表示").Visible = xlSheetHidden

    c = mSheet.SheetCandidates("画面設計")
    TestLib.Eq "候補: 一致するシートの数 (非表示を除く)", UBound(c) - LBound(c) + 1, 3
    TestLib.Eq "候補: 完全一致が先頭", c(0)(0), "画面設計"
    TestLib.Eq "候補: 次は先頭で一致する名前", c(1)(0), "画面設計 一覧"
    TestLib.Eq "候補: 先頭以外での一致は後ろ", c(2)(0), "旧画面設計"
    TestLib.Eq "候補: 説明は空", c(0)(1), ""
    TestLib.Eq "候補: 一致しなければ空の配列", UBound(mSheet.SheetCandidates("存在しない")), -1

    c = mSheet.SheetCandidates("")
    TestLib.Eq "候補: 空の検索語は表示中の全シート", UBound(c) - LBound(c) + 1, CountVisible(wb)

    ' ---- ff 本体 ----
    wb.Worksheets("DB設計").Activate
    mSheet.ff "画面設計"
    TestLib.Eq "ff: 完全一致のシートへ移動する", ActiveSheet.Name, "画面設計"
    mSheet.ff "一覧"
    TestLib.Eq "ff: いちばん良く一致するシートへ移動する", ActiveSheet.Name, "画面設計 一覧"
    mSheet.ff "db"
    TestLib.Eq "ff: 大文字・小文字を区別しない", ActiveSheet.Name, "DB設計"
    mSheet.ff "存在しない"
    TestLib.Eq "ff: 一致しなければ移動しない", ActiveSheet.Name, "DB設計"
    mSheet.ff "非表示"
    TestLib.Eq "ff: 非表示のシートへは移動しない", ActiveSheet.Name, "DB設計"
    mSheet.ff
    TestLib.Eq "ff: 引数なしでは移動しない", ActiveSheet.Name, "DB設計"
End Sub

' タブの末尾にシートを足す
Private Function AddSheet(ByVal wb As Object, ByVal sheetName As String) As Object
    Dim ws As Object
    Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = sheetName
    Set AddSheet = ws
End Function

Private Function CountVisible(ByVal wb As Object) As Long
    Dim sht As Object
    For Each sht In wb.Sheets
        If sht.Visible = xlSheetVisible Then CountVisible = CountVisible + 1
    Next sht
End Function
