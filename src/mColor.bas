Attribute VB_Name = "mColor"
Option Explicit
'===============================================================================
' NAME
'     bg, fg, bo - Excel セル/図形の色を取得・設定・コピー・貼り付けする
'
' SYNOPSIS
'     bg [ copy | c | paste | p | set COLOR | COLOR ]
'     fg [ copy | c | paste | p | set COLOR | COLOR ] [START [END]]
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
'     fg のみ、COLOR の後に START END を続けると、アクティブセル内の
'     START～END 区間 (両端の断片を含む) だけ文字色を変える (部分着色)。
'         fg red500 [ ]        [ から ] までを赤にする
'         fg set red500 [ ]    同上
'     区間が複数あればすべて変更する。START が複数ある場合は END に最も近い
'     ものを採用 (最短一致)。大文字・小文字は区別する。END を省略すると
'     InputBox で入力する (空白を含む断片はこちらで指定)。
'     数式セル・文字列以外のセル・図形は対象外。
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
'     - fg は Range.Font / 図形テキストのフォント色を扱う。現在色の表示・copy では
'       セル内の部分的な色は考慮しない。
'     - bo は図形の枠線色のみ。セルの罫線は対象外。
'     - すべてのメッセージやエラーはステータスバーに表示される。
'
' COMPLETIONS
'     bg              セル/図形の塗り色
'     bg copy         色をコピー
'     bg paste        色を貼り付け
'     bg set          色を設定
'     fg              セル/図形の文字色
'     fg copy         色をコピー
'     fg paste        色を貼り付け
'     fg set          色を設定
'     bo              図形の枠線色
'     bo copy         色をコピー
'     bo paste        色を貼り付け
'     bo set          色を設定
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
Public Sub fg(Optional a As String, Optional b As String, Optional c As String, Optional d As String): Dispatch "fg", a, b, c, d: End Sub
Public Sub bo(Optional a As String, Optional b As String): Dispatch "bo", a, b: End Sub

'========================
' 分岐専用ディスパッチャ
'========================
Private Sub Dispatch(ByVal target As String, Optional arg1 As String, Optional arg2 As String, _
                     Optional arg3 As String, Optional arg4 As String)
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim cmd As String
    cmd = LCase$(Trim$(arg1))
    Select Case cmd
        Case "": ShowCurrentColor target
        Case "copy", "c": DoCopy target
        Case "paste", "p": DoPaste target
        Case "set", "s": DoSet target, arg2, arg3, arg4
        Case Else: DoSet target, cmd, arg2, arg3
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

Private Sub DoSet(ByVal target As String, ByVal token As String, _
                  Optional ByVal startFrag As String, Optional ByVal endFrag As String)
    Dim sr As ShapeRange, c As Long, ok As Boolean, i As Long, applied As Long

    token = LCase$(Trim$(token))
    If token = "" Then Application.StatusBar = "[" & target & "] Please specify a color.": Exit Sub

    ok = TryParseHexColor(token, c): If Not ok Then ok = mColorUtil.TryParseHueToken(token, c)
    If Not ok Then Application.StatusBar = "[" & target & "] Invalid color: " & token: Exit Sub

    On Error Resume Next: Set sr = Selection.ShapeRange: On Error GoTo 0

    If Len(startFrag) > 0 Then
        If target <> "fg" Or Not sr Is Nothing Then
            Application.StatusBar = "[" & target & "] Fragments are supported only for fg on cells.": Exit Sub
        End If
        DoSetFragment c, startFrag, endFrag
        Exit Sub
    End If

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

' アクティブセル内の startFrag～endFrag 区間の文字色を c にする (fg の部分着色)
Private Sub DoSetFragment(ByVal c As Long, ByVal startFrag As String, ByVal endFrag As String)
    Dim cell As Range
    Dim n As Long

    Set cell = ActiveCell
    If cell Is Nothing Then Application.StatusBar = "[fg] No active cell.": Exit Sub
    If cell.HasFormula Or VarType(cell.Value) <> vbString Then
        Application.StatusBar = "[fg] Active cell is not a text constant.": Exit Sub
    End If

    If Len(endFrag) = 0 Then endFrag = InputBox("終了の文字断片", "fg")
    If Len(endFrag) = 0 Then Exit Sub

    n = ColorBetween(cell, startFrag, endFrag, c)
    Application.StatusBar = "[fg] " & n & " span(s) set: " & ColorLongToHex(c)
End Sub

' cell の文字列中の startFrag～endFrag 区間 (最短一致) を fgColor で塗り、塗った区間数を返す
Private Function ColorBetween(ByVal cell As Range, ByVal startFrag As String, ByVal endFrag As String, ByVal fgColor As Long) As Long
    Dim text As String
    Dim pos As Long
    Dim s As Long
    Dim e As Long

    text = CStr(cell.Value)
    pos = 1

    Do
        ' まず開始断片を探す
        s = InStr(pos, text, startFrag, vbBinaryCompare)
        If s = 0 Then Exit Do

        ' その後ろで最初に現れる終了断片を探す
        e = InStr(s + Len(startFrag), text, endFrag, vbBinaryCompare)
        If e = 0 Then Exit Do

        ' 終了断片の直前にある最後の開始断片まで寄せる (最短一致)
        s = InStrRev(Left$(text, e - 1), startFrag, -1, vbBinaryCompare)

        cell.Characters(s, e + Len(endFrag) - s).Font.Color = fgColor
        ColorBetween = ColorBetween + 1

        pos = e + Len(endFrag)
    Loop
End Function

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
