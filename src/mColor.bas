Attribute VB_Name = "mColor"
Option Explicit
'===============================================================================
' NAME
'     bg, fg, ol - Excel セル/図形の色を取得・設定・コピー・貼り付けする
'
' SYNOPSIS
'     bg [ copy | c | paste | p | set COLOR | COLOR ]
'     fg [ copy | c | paste | p | set COLOR | COLOR ]
'     bo [ copy | c | paste | p | set COLOR | COLOR ]
'
' DESCRIPTION
'     bg はセル/図形の塗り (Range.Interior / Shape.Fill) を操作する。
'     fg はセル/図形の文字色 (Range.Font / Shape.TextFrame2 など) を操作する。
'     bo は図形の枠線色 (Shape.Line) を操作する。セルの罫線は対象外。
'     引数なしで実行すると、現在の色を #rrggbb 形式でステータスバーに表示する。
'
' COMMANDS
'     copy, c    アクティブセル (または選択中の先頭図形) の色をコピー用バッファに保存。
'     paste, p   コピー済みの色を選択範囲の全セル/全図形に適用。
'     set COLOR  指定色をアクティブセル/先頭図形に設定。
'     COLOR      set の短縮形。COLOR の形式は ARGUMENTS を参照。
'
' ARGUMENTS
'     COLOR は以下の形式が使用可能:
'         #rrggbb        - 16進指定 (例: #ffcc00)
'         rrggbb         - 16進指定 (先頭の # は省略可)
'         HUE[STEP]      - Tailwind CSS v4 の配色 (例: blue50〜blue950, amber500)
'                          STEP は 50,100,200,300,400,500,600,700,800,900,950。
'                          50 が最も明るく、950 が最も暗い (濃い)
'         HUE            - STEP 省略時は既定段階 500 にフォールバック
'
'     HUE は以下のいずれか:
'         red, orange, amber, yellow, lime, green, emerald, teal, cyan, sky,
'         blue, indigo, violet, purple, fuchsia, pink, rose,
'         slate, gray, zinc, neutral, stone, black, white
'
'     色の一覧は palette マクロ (mPalette.bas) を実行するとシートで確認できる。
'
' NOTES
'     - スコープは auto。セル選択時はセルへ、図形選択時は図形へ操作する。
'     - コピー用バッファは bg / fg / bo で独立 (Variant 番兵: Empty)。
'     - fg は Range.Font / 図形テキストのフォント色を扱う。セル内の部分的な色は考慮しない。
'     - ol は図形の枠線色のみ。セルの罫線は対象外。
'     - すべてのメッセージやエラーはステータスバーに表示される。
'===============================================================================

'========================
' バッファ (Variant 番兵: Empty)
'========================
Private gBgColor As Variant
Private gFgColor As Variant
Private gBoColor As Variant

' target ("bg"/"fg"/"bo") に対応するコピー用バッファの値を返す
Private Function GetCopiedColor(ByVal target As String) As Variant
    Select Case target
        Case "bg": GetCopiedColor = gBgColor
        Case "fg": GetCopiedColor = gFgColor
        Case "bo": GetCopiedColor = gBoColor
    End Select
End Function

'========================
' エントリポイント
'========================
Public Sub bg(Optional a As String, Optional b As String): Dispatch "bg", a, b: End Sub
Public Sub fg(Optional a As String, Optional b As String): Dispatch "fg", a, b: End Sub
Public Sub bo(Optional a As String, Optional b As String): Dispatch "bo", a, b: End Sub

'========================
' 分岐専用ディスパッチャ
'========================
Private Sub Dispatch(ByVal target As String, Optional arg1 As String, Optional arg2 As String)
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim cmd As String
    cmd = LCase$(Trim$(arg1))
    Select Case cmd
        Case "": ShowCurrentColor target
        Case "copy", "c": DoCopy target
        Case "paste", "p": DoPaste target
        Case "set", "s": DoSet target, arg2
        Case Else: DoSet target, cmd
    End Select
    Exit Sub

FINALLY:
    ' エラーでも Sub を正常終了させ、booster の Class_Terminate で
    ' ScreenUpdating / EnableEvents を確実に復元する。
    Application.StatusBar = "[" & target & "] E" & Err.Number & ": " & Err.Description
End Sub

Private Sub DoCopy(ByVal target As String)
    Dim c As Long
    Dim s As shape
    Dim shapeIsSelected As String

    On Error Resume Next
    Set s = Selection.ShapeRange(1)
    On Error GoTo 0

    shapeIsSelected = CStr(Abs(Not s Is Nothing))
    Select Case target & "," & shapeIsSelected
        Case "bg,0": gBgColor = CLng(ActiveCell.Interior.Color)
        Case "bg,1": gBgColor = CLng(s.Fill.ForeColor.RGB)

        Case "fg,0": gFgColor = CLng(ActiveCell.Font.Color)
        Case "fg,1":
            If GetShapeTextColor(s, c) Then
                gFgColor = CLng(c)
            Else
                Application.StatusBar = "[fg] No text in the selected shape."
                Exit Sub
            End If

        Case "bo,1": gBoColor = CLng(s.line.ForeColor.RGB)

        Case Else
            Application.StatusBar = "[" & target & "] No valid source to copy."
            Exit Sub
    End Select

    Application.StatusBar = "[" & target & "] copied: " & ColorLongToHex(CLng(GetCopiedColor(target)))
End Sub

Private Sub DoPaste(ByVal target As String)
    Dim sr As ShapeRange, i As Long

    On Error Resume Next: Set sr = Selection.ShapeRange: On Error GoTo 0

    If (target = "bg" And IsEmpty(gBgColor)) _
    Or (target = "fg" And IsEmpty(gFgColor)) _
    Or (target = "bo" And IsEmpty(gBoColor)) Then
        Application.StatusBar = "[" & target & "] No color copied yet.": Exit Sub
    End If

    Select Case target & "," & CStr(Abs(Not sr Is Nothing))
        Case "bg,0":
            Selection.Interior.Color = gBgColor
        Case "bg,1":
            For i = 1 To sr.Count
                sr(i).Fill.Visible = True
                sr(i).Fill.ForeColor.RGB = gBgColor
            Next
        Case "fg,0":
            Selection.Font.Color = gFgColor
        Case "fg,1":
            For i = 1 To sr.Count
                SetShapeTextColor sr(i), gFgColor
            Next
        Case "bo,0":
            Application.StatusBar = "[bo] No shapes in selection."
            Exit Sub
        Case "bo,1":
            For i = 1 To sr.Count
                sr(i).line.Visible = True
                sr(i).line.ForeColor.RGB = gBoColor
            Next
        Case Else:
            Application.StatusBar = "[" & target & "] Unsupported context."
            Exit Sub
    End Select

    Application.StatusBar = "[" & target & "] pasted: " & ColorLongToHex(CLng(GetCopiedColor(target)))
End Sub

Private Sub DoSet(ByVal target As String, ByVal token As String)
    Dim sr As ShapeRange, c As Long, ok As Boolean, i As Long, applied As Long

    token = LCase$(Trim$(token))
    If token = "" Then Application.StatusBar = "[" & target & "] Please specify a color.": Exit Sub

    ok = TryParseHexColor(token, c): If Not ok Then ok = mColorUtil.TryParseHueToken(token, c)
    If Not ok Then Application.StatusBar = "[" & target & "] Invalid color: " & token: Exit Sub

    On Error Resume Next: Set sr = Selection.ShapeRange: On Error GoTo 0

    Select Case target & "," & CStr(Abs(Not sr Is Nothing))
        Case "bg,0"
            Selection.Interior.Color = c

        Case "bg,1"
            For i = 1 To sr.Count
                sr(i).Fill.Visible = True
                sr(i).Fill.ForeColor.RGB = c
            Next

        Case "fg,0"
            Selection.Font.Color = c

        Case "fg,1"
            For i = 1 To sr.Count
                If SetShapeTextColor(sr(i), c) Then applied = applied + 1
            Next
            If applied = 0 Then Application.StatusBar = "[fg] No shapes with text.": Exit Sub

        Case "bo,0"
            Application.StatusBar = "[bo] No shape selected.": Exit Sub

        Case "bo,1"
            For i = 1 To sr.Count
                sr(i).line.Visible = True
                sr(i).line.ForeColor.RGB = c
            Next

        Case Else
            Application.StatusBar = "[" & target & "] Set: unsupported context.": Exit Sub
    End Select

    Application.StatusBar = "[" & target & "] set: " & ColorLongToHex(c)
End Sub

Private Sub ShowCurrentColor(ByVal target As String)
    Dim c As Long
    Dim s As shape
    Dim shapeIsSelected As String
    Dim colorStr As String

    On Error Resume Next
    Set s = Selection.ShapeRange(1)
    On Error GoTo 0

    shapeIsSelected = CStr(Abs(Not s Is Nothing))
    Select Case target & "," & shapeIsSelected
        Case "bg,0"
            If ActiveCell.Interior.ColorIndex = xlColorIndexNone Then
                colorStr = "none"
            Else
                colorStr = ColorLongToHex(ActiveCell.Interior.Color)
            End If

        Case "bg,1"
            If s.Fill.Visible Then
                colorStr = ColorLongToHex(s.Fill.ForeColor.RGB)
            Else
                colorStr = "none"
            End If

        Case "fg,0"
            colorStr = ColorLongToHex(ActiveCell.Font.Color)

        Case "fg,1"
            If GetShapeTextColor(s, c) Then
                colorStr = ColorLongToHex(c)
            Else
                colorStr = "(no text)"
            End If

        Case "bo,0"
            colorStr = "(n/a for cells)"

        Case "bo,1"
            If s.line.Visible Then
                colorStr = ColorLongToHex(s.line.ForeColor.RGB)
            Else
                colorStr = "none"
            End If
        Case Else
            colorStr = "(unsupported)"
    End Select

    Application.StatusBar = "[" & target & "] current :" & colorStr
End Sub

Private Function GetShapeTextColor(ByVal shp As shape, ByRef outColor As Long) As Boolean
    If Not shp.TextFrame2.HasText Then Exit Function
    outColor = shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    GetShapeTextColor = True
End Function

Private Function SetShapeTextColor(ByVal shp As shape, ByVal c As Long) As Boolean
    If Not shp.TextFrame2.HasText Then Exit Function
    shp.TextFrame2.TextRange.Font.Fill.Visible = True
    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = c
End Function
