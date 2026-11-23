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
'         none           - bg / bo のみ。塗りつぶしなし / 枠線なし (透明)
'         auto           - fg のみ。文字色を「自動」に戻す (セルのみ)
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
'       塗りなし (none)・自動 (auto) もそのままコピー・貼り付けできる。
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
'     bg none         塗りつぶしなし
'     fg              セル/図形の文字色
'     fg copy         色をコピー
'     fg paste        色を貼り付け
'     fg set          色を設定
'     fg auto         文字色を自動に戻す
'     bo              図形の枠線色
'     bo copy         色をコピー
'     bo paste        色を貼り付け
'     bo set          色を設定
'     bo none         枠線なし
'===============================================================================

'========================
' バッファ (Variant 番兵: Empty)
'========================
' 色の値は RGB の Long。RGB は 0 以上なので、負の値を「なし」「自動」の番兵に使う
Private Const COLOR_NONE As Long = -1 ' 塗り・枠線なし (bg / bo)
Private Const COLOR_AUTO As Long = -2 ' 文字色の自動 (fg、セルのみ)

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
        Case "bg,0"
            If ActiveCell.Interior.ColorIndex = xlColorIndexNone Then
                gBgColor = COLOR_NONE
            Else
                gBgColor = CLng(ActiveCell.Interior.Color)
            End If
        Case "bg,1"
            If s.Fill.Visible = msoFalse Then
                gBgColor = COLOR_NONE
            Else
                gBgColor = CLng(s.Fill.ForeColor.RGB)
            End If

        Case "fg,0"
            If ActiveCell.Font.ColorIndex = xlColorIndexAutomatic Then
                gFgColor = COLOR_AUTO
            Else
                gFgColor = CLng(ActiveCell.Font.Color)
            End If
        Case "fg,1"
            If GetShapeTextColor(s, c) Then
                gFgColor = CLng(c)
            Else
                Application.StatusBar = "[fg] No text in the selected shape."
                Exit Sub
            End If

        Case "bo,1"
            If s.line.Visible = msoFalse Then
                gBoColor = COLOR_NONE
            Else
                gBoColor = CLng(s.line.ForeColor.RGB)
            End If

        Case Else
            Application.StatusBar = "[" & target & "] No valid source to copy."
            Exit Sub
    End Select

    Application.StatusBar = "[" & target & "] copied: " & ColorText(CLng(GetCopiedColor(target)))
End Sub

Private Sub DoPaste(ByVal target As String)
    If IsEmpty(GetCopiedColor(target)) Then
        Application.StatusBar = "[" & target & "] No color copied yet.": Exit Sub
    End If
    If ApplyColor(target, CLng(GetCopiedColor(target))) Then
        Application.StatusBar = "[" & target & "] pasted: " & ColorText(CLng(GetCopiedColor(target)))
    End If
End Sub

Private Sub DoSet(ByVal target As String, ByVal token As String, _
                  Optional ByVal startFrag As String, Optional ByVal endFrag As String)
    Dim sr As ShapeRange, c As Long

    token = LCase$(Trim$(token))
    If token = "" Then Application.StatusBar = "[" & target & "] Please specify a color.": Exit Sub
    If Not TryParseColorFor(target, token, c) Then
        Application.StatusBar = "[" & target & "] Invalid color: " & token: Exit Sub
    End If

    If Len(startFrag) > 0 Then
        On Error Resume Next: Set sr = Selection.ShapeRange: On Error GoTo 0
        If target <> "fg" Or Not sr Is Nothing Then
            Application.StatusBar = "[" & target & "] Fragments are supported only for fg on cells.": Exit Sub
        End If
        DoSetFragment c, startFrag, endFrag
        Exit Sub
    End If

    If ApplyColor(target, c) Then
        Application.StatusBar = "[" & target & "] set: " & ColorText(c)
    End If
End Sub

' 色 c (COLOR_NONE / COLOR_AUTO を含む) を選択中のセル/図形に適用する。
' 適用できなければステータスバーに理由を出して False を返す
Private Function ApplyColor(ByVal target As String, ByVal c As Long) As Boolean
    Dim sr As ShapeRange, i As Long, applied As Long

    On Error Resume Next: Set sr = Selection.ShapeRange: On Error GoTo 0

    Select Case target & "," & CStr(Abs(Not sr Is Nothing))
        Case "bg,0"
            If c = COLOR_NONE Then
                Selection.Interior.ColorIndex = xlColorIndexNone
            Else
                Selection.Interior.Color = c
            End If

        Case "bg,1"
            For i = 1 To sr.Count
                If c = COLOR_NONE Then
                    sr(i).Fill.Visible = msoFalse
                Else
                    sr(i).Fill.Visible = msoTrue
                    sr(i).Fill.ForeColor.RGB = c
                End If
            Next

        Case "fg,0"
            If c = COLOR_AUTO Then
                Selection.Font.ColorIndex = xlColorIndexAutomatic
            Else
                Selection.Font.Color = c
            End If

        Case "fg,1"
            If c = COLOR_AUTO Then
                Application.StatusBar = "[fg] auto is supported only for cells.": Exit Function
            End If
            For i = 1 To sr.Count
                If SetShapeTextColor(sr(i), c) Then applied = applied + 1
            Next
            If applied = 0 Then Application.StatusBar = "[fg] No shapes with text.": Exit Function

        Case "bo,0"
            Application.StatusBar = "[bo] No shape selected.": Exit Function

        Case "bo,1"
            For i = 1 To sr.Count
                If c = COLOR_NONE Then
                    sr(i).line.Visible = msoFalse
                Else
                    sr(i).line.Visible = msoTrue
                    sr(i).line.ForeColor.RGB = c
                End If
            Next

        Case Else
            Application.StatusBar = "[" & target & "] Unsupported context.": Exit Function
    End Select

    ApplyColor = True
End Function

' COLOR に加え、bg / bo では none (塗り・枠線なし)、fg では auto (自動) を受け付ける
Private Function TryParseColorFor(ByVal target As String, ByVal token As String, ByRef outColor As Long) As Boolean
    Select Case True
        Case token = "none" And (target = "bg" Or target = "bo")
            outColor = COLOR_NONE
            TryParseColorFor = True
        Case token = "auto" And target = "fg"
            outColor = COLOR_AUTO
            TryParseColorFor = True
        Case Else
            TryParseColorFor = TryParseHexColor(token, outColor)
            If Not TryParseColorFor Then TryParseColorFor = mColorUtil.TryParseHueToken(token, outColor)
    End Select
End Function

' 表示用の色文字列 (#rrggbb / none / auto)
Private Function ColorText(ByVal c As Long) As String
    Select Case c
        Case COLOR_NONE: ColorText = "none"
        Case COLOR_AUTO: ColorText = "auto"
        Case Else: ColorText = ColorLongToHex(c)
    End Select
End Function

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
    Application.StatusBar = "[fg] " & n & " span(s) set: " & ColorText(c)
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

        With cell.Characters(s, e + Len(endFrag) - s).Font
            If fgColor = COLOR_AUTO Then .ColorIndex = xlColorIndexAutomatic Else .Color = fgColor
        End With
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
            If ActiveCell.Font.ColorIndex = xlColorIndexAutomatic Then
                colorStr = "auto"
            Else
                colorStr = ColorLongToHex(ActiveCell.Font.Color)
            End If

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
