Attribute VB_Name = "mView"
Option Explicit
'===============================================================================
' NAME
'     fmt, pbp - 全シートの表示状態を一括で変更する
'
' SYNOPSIS
'     fmt
'     pbp
'
' DESCRIPTION
'     fmt は全シートを A1 選択・ズーム 100% にリセットする。
'     pbp は全シートを改ページプレビュー表示にし、目盛線を非表示にする。
'     いずれも実行後、先頭シートをアクティブにする。
'
' COMPLETIONS
'     fmt             全シートを A1・ズーム 100% にリセット
'     pbp             全シートを改ページプレビュー表示に
'===============================================================================

' 全シート A1 を選択し、縮尺100% に。
Public Sub fmt()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Type = xlWorksheet Then
            ws.Activate
            ActiveWindow.Zoom = 100
            Application.Goto ws.Range("A1"), True
        End If
    Next ws

    ActiveWorkbook.Worksheets(1).Activate
    Exit Sub

FINALLY:
    Application.StatusBar = "[fmt] E" & Err.Number & ": " & Err.Description
End Sub

' 全シートを改ページプレビュー表示にし、目盛線を非表示に。
Public Sub pbp()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Type = xlWorksheet Then
            ws.Activate
            ActiveWindow.View = xlPageBreakPreview
            ActiveWindow.DisplayGridlines = False
        End If
    Next ws

    ActiveWorkbook.Worksheets(1).Activate
    Exit Sub

FINALLY:
    Application.StatusBar = "[pbp] E" & Err.Number & ": " & Err.Description
End Sub
