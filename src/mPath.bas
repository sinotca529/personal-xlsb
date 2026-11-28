Attribute VB_Name = "mPath"
Option Explicit
'===============================================================================
' NAME
'     path - アクティブなブックのパス・ファイル名などをクリップボードにコピーする
'
' SYNOPSIS
'     path
'     path name
'     path dir
'     path cell
'
' DESCRIPTION
'     アクティブなブックについて、次のものをクリップボードにコピーする。
'         path        フルパス             例: C:\work\設計書.xlsx
'         path name   ファイル名           例: 設計書.xlsx
'         path dir    フォルダのパス       例: C:\work
'         path cell   選択範囲の場所       例: 設計書.xlsx シート1!B3:F10
'     path cell は人が読むための形 (数式の参照の書き方ではない)。複数の範囲を
'     選んでいれば A1:B2,D4 のようにつなげる。
'
' NOTES
'     - 保存していないブックにはパスがないので、path / path dir は何もコピーしない。
'       path name と path cell は仮の名前 (Book1 など) を使う。
'     - OneDrive や SharePoint 上のブックは、Excel が返すとおり URL のままコピーする。
'     - path cell は、セルを選んでいないとき (図形の選択中など) は何もコピーしない。
'
' COMPLETIONS
'     path            ブックのフルパスをコピー
'     path name       ファイル名をコピー
'     path dir        フォルダのパスをコピー
'     path cell       ブック名・シート名・選択範囲をコピー
'===============================================================================

Public Sub path(Optional ByVal kind As String)
    On Error GoTo FINALLY

    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        Application.StatusBar = "[path] No workbook is open."
        Exit Sub
    End If

    Dim text As String
    Select Case LCase$(kind)
        Case ""
            If Not IsSaved(wb) Then Exit Sub
            text = wb.FullName
        Case "name"
            text = wb.Name
        Case "dir"
            If Not IsSaved(wb) Then Exit Sub
            text = wb.Path
        Case "cell"
            If TypeName(Selection) <> "Range" Then
                Application.StatusBar = "[path] Select cells for ""path cell""."
                Exit Sub
            End If
            text = CellLocation(wb, Selection)
        Case Else
            Application.StatusBar = "[path] Unknown kind: " & kind & " (name / dir / cell)"
            Exit Sub
    End Select

    If mClipboard.SetClipboardText(text) Then
        Application.StatusBar = "[path] Copied: " & text
    Else
        Application.StatusBar = "[path] Failed to write to the clipboard."
    End If
    Exit Sub

FINALLY:
    Application.StatusBar = "[path] E" & Err.Number & ": " & Err.Description
End Sub

' 人が読むための選択範囲の場所 (例: 設計書.xlsx シート1!B3:F10)
Public Function CellLocation(ByVal wb As Workbook, ByVal rng As Range) As String
    CellLocation = wb.Name & " " & rng.Worksheet.Name & "!" & rng.Address(False, False)
End Function

' 保存済みか。保存していなければステータスバーに出して False
Private Function IsSaved(ByVal wb As Workbook) As Boolean
    IsSaved = (Len(wb.Path) > 0)
    If Not IsSaved Then Application.StatusBar = "[path] The workbook has not been saved yet."
End Function
