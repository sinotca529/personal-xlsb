Attribute VB_Name = "Test_mUtil"
Option Explicit
'===============================================================================
' tb (罫線整形) の結合テスト。
'
' NOTE: LibreOffice の Borders(...).LineStyle は getter が常に xlContinuous を
'       返すため、VBA 側から線種を読み戻す検証はできない。実際に引かれた罫線の
'       検証は Python 側 (UNO の TopBorder / LeftBorder) で行う。
'       ここでは「例外を投げずに走り切ること」と「ガード条件」を検証する。
'===============================================================================

Public Const TB_SHEET As String = "TbCase"

Public Sub Run()
    Dim ws As Object
    Set ws = ActiveWorkbook.Worksheets.Add
    ws.Name = TB_SHEET
    ws.Activate

    ws.Range("A1").Value = "H1"
    ws.Range("B1").Value = "H2"
    ws.Range("C1").Value = "H3"
    ws.Range("A2").Value = "r1"
    ' A3 は空欄のまま (行区切りの対象外になることを期待する)
    ws.Range("A4").Value = "r3"

    ws.Range("A1:C4").Select
    TestLib.Eq "tb 1: 例外を投げずに完了する", CallTb(1), 0

    ' ---- ガード条件 ----
    ' headerRows が範囲外のときは何もせずに抜ける (例外も出さない)。
    ws.Range("A10:C13").Select
    TestLib.Eq "tb 0: headerRows < 1 でも例外にならない", CallTb(0), 0
    TestLib.Eq "tb 99: headerRows > 行数 でも例外にならない", CallTb(99), 0
End Sub

' tb を呼び、呼び出し元まで伝播した実行時エラー番号を返す (0 = 正常)。
Private Function CallTb(ByVal headerRows As Long) As Long
    On Error GoTo Trapped
    mUtil.tb headerRows
    CallTb = 0
    Exit Function
Trapped:
    CallTb = Err.Number
End Function
