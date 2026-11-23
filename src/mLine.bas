Attribute VB_Name = "mLine"
Option Explicit
'===============================================================================
' NAME
'     ln - 図形の線の書式 (太さ・種類・色・矢印) を取得・設定・コピー・貼り付けする
'
' SYNOPSIS
'     ln
'     ln TOKEN...
'     ln { copy | c | paste | p }
'
' DESCRIPTION
'     選択中の図形の線 (枠線) の書式を操作する。引数なしで実行すると、先頭の図形の
'     線の書式をステータスバーに表示する (例: [ln] 0.75pt dash #ff0000 ->)。
'
'     TOKEN は空白区切りで複数並べられ、順不同。それぞれ次のように判別する。
'         数値 / 数値pt   太さ (pt)。例: 2, 0.75, 1.5pt
'         COLOR           色。#rrggbb / rrggbb / HUE[STEP] (bg と同じ。mColor.bas を参照)
'         線の種類        solid (実線)      dot (丸点線)        sdot (角点線)
'                         dash (破線)       dashdot (一点鎖線)
'                         ldash (長破線)    ldashdot (長一点鎖線)
'         矢印            -> (終点)  <- (始点)  <-> (両端)  - (なし)
'         none            線を消す (ほかのトークンとは併用できない)
'     none 以外を指定すると、線が非表示なら表示に戻す。
'
'     例: ln 2              太さ 2pt
'         ln 2 dash red     赤い 2pt の破線
'         ln ->             終点に矢印
'
' COMMANDS
'     copy, c    先頭の図形の線の書式 (表示/非表示・太さ・種類・色・矢印) をコピー。
'     paste, p   コピーした線の書式を選択中の全図形に適用。
'
' NOTES
'     - 対象は図形のみ (セルの罫線は対象外)。
'     - 6桁の16進数 (例: 123456) は太さではなく色とみなす。
'     - 色だけを扱う bo (mColor.bas) とコピー用バッファは独立。
'
' COMPLETIONS
'     ln              図形の線の書式 (太さ・種類・色・矢印)
'     ln copy         線の書式をコピー
'     ln paste        線の書式を貼り付け
'     ln none         線を消す
'     ln 0.25         太さ 0.25pt
'     ln 0.5          太さ 0.5pt
'     ln 0.75         太さ 0.75pt
'     ln 1            太さ 1pt
'     ln 1.5          太さ 1.5pt
'     ln 2.25         太さ 2.25pt
'     ln 3            太さ 3pt
'     ln 4.5          太さ 4.5pt
'     ln 6            太さ 6pt
'     ln solid        実線
'     ln dot          丸点線
'     ln sdot         角点線
'     ln dash         破線
'     ln dashdot      一点鎖線
'     ln ldash        長破線
'     ln ldashdot     長一点鎖線
'     ln ->           終点に矢印
'     ln <-           始点に矢印
'     ln <->          両端に矢印
'     ln -            矢印なし
'===============================================================================

' 線の書式。各項目は Empty なら「指定なし (変更しない)」
Private Type LineSpec
    Visible As Variant              ' Boolean
    Weight As Variant               ' Double (pt)
    DashStyle As Variant            ' MsoLineDashStyle
    Color As Variant                ' Long (RGB)
    BeginArrowheadStyle As Variant  ' MsoArrowheadStyle
    EndArrowheadStyle As Variant    ' MsoArrowheadStyle
End Type

Private gCopied As LineSpec
Private gHasCopy As Boolean

Public Sub ln(Optional a As String, Optional b As String, Optional c As String, _
              Optional d As String, Optional e As String, Optional f As String)
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim sr As ShapeRange
    Dim tokens As Variant
    Dim spec As LineSpec
    Dim errToken As String

    On Error Resume Next: Set sr = Selection.ShapeRange: On Error GoTo FINALLY
    If sr Is Nothing Then
        Application.StatusBar = "[ln] No shape selected."
        Exit Sub
    End If

    tokens = Array(a, b, c, d, e, f)
    Select Case LCase$(Trim$(a))
        Case ""
            Application.StatusBar = "[ln] current: " & DescribeLine(sr(1).line)
            Exit Sub
        Case "copy", "c"
            gCopied = ReadLine(sr(1).line)
            gHasCopy = True
            Application.StatusBar = "[ln] copied: " & DescribeSpec(gCopied)
            Exit Sub
        Case "paste", "p"
            If Not gHasCopy Then
                Application.StatusBar = "[ln] Nothing copied yet."
                Exit Sub
            End If
            spec = gCopied
        Case Else
            If Not ParseTokens(tokens, spec, errToken) Then
                Application.StatusBar = "[ln] Invalid token: " & errToken
                Exit Sub
            End If
    End Select

    Dim i As Long
    For i = 1 To sr.Count
        ApplySpec sr(i).line, spec
    Next i
    Application.StatusBar = "[ln] set: " & DescribeLine(sr(1).line)
    Exit Sub

FINALLY:
    Application.StatusBar = "[ln] E" & Err.Number & ": " & Err.Description
End Sub

'========================
' 解析
'========================
Private Function ParseTokens(ByVal tokens As Variant, ByRef spec As LineSpec, ByRef errToken As String) As Boolean
    Dim t As Variant
    Dim tok As String
    Dim c As Long
    Dim hasNone As Boolean, hasOther As Boolean

    For Each t In tokens
        tok = LCase$(Trim$(t))
        If Len(tok) = 0 Then GoTo NEXT_TOKEN
        errToken = tok

        If tok = "none" Then
            spec.Visible = False
            hasNone = True
        ElseIf TryParseDash(tok, c) Then
            spec.DashStyle = c
            hasOther = True
        ElseIf TryParseArrow(tok, spec) Then
            hasOther = True
        ElseIf IsHexColorToken(tok) Then
            If Not TryParseHexColor(tok, c) Then Exit Function
            spec.Color = c
            hasOther = True
        ElseIf TryParseWeight(tok, spec) Then
            hasOther = True
        ElseIf mColorUtil.TryParseHueToken(tok, c) Then
            spec.Color = c
            hasOther = True
        Else
            Exit Function
        End If
NEXT_TOKEN:
    Next t

    If hasNone And hasOther Then errToken = "none (cannot be combined)": Exit Function
    If hasOther Then spec.Visible = True
    ParseTokens = hasNone Or hasOther
End Function

' # 付き、または 6 桁の 16 進数は色とみなす (123456 を太さと誤認しないため)
Private Function IsHexColorToken(ByVal tok As String) As Boolean
    If Left$(tok, 1) = "#" Then IsHexColorToken = True: Exit Function
    IsHexColorToken = (Len(tok) = 6 And Not tok Like "*[!0-9a-f]*")
End Function

Private Function TryParseWeight(ByVal tok As String, ByRef spec As LineSpec) As Boolean
    If Right$(tok, 2) = "pt" Then tok = Left$(tok, Len(tok) - 2)
    If Len(tok) = 0 Or tok Like "*[!0-9.]*" Then Exit Function
    If Len(tok) - Len(Replace(tok, ".", "")) > 1 Then Exit Function
    If Val(tok) <= 0 Then Exit Function
    spec.Weight = Val(tok)
    TryParseWeight = True
End Function

Private Function TryParseDash(ByVal tok As String, ByRef outStyle As Long) As Boolean
    TryParseDash = True
    Select Case tok
        Case "solid": outStyle = msoLineSolid
        Case "dot": outStyle = msoLineRoundDot
        Case "sdot": outStyle = msoLineSquareDot
        Case "dash": outStyle = msoLineDash
        Case "dashdot": outStyle = msoLineDashDot
        Case "ldash": outStyle = msoLineLongDash
        Case "ldashdot": outStyle = msoLineLongDashDot
        Case Else: TryParseDash = False
    End Select
End Function

Private Function TryParseArrow(ByVal tok As String, ByRef spec As LineSpec) As Boolean
    TryParseArrow = True
    Select Case tok
        Case "->":  spec.BeginArrowheadStyle = msoArrowheadNone:     spec.EndArrowheadStyle = msoArrowheadTriangle
        Case "<-":  spec.BeginArrowheadStyle = msoArrowheadTriangle: spec.EndArrowheadStyle = msoArrowheadNone
        Case "<->": spec.BeginArrowheadStyle = msoArrowheadTriangle: spec.EndArrowheadStyle = msoArrowheadTriangle
        Case "-":   spec.BeginArrowheadStyle = msoArrowheadNone:     spec.EndArrowheadStyle = msoArrowheadNone
        Case Else: TryParseArrow = False
    End Select
End Function

'========================
' 読み書き
'========================
Private Function ReadLine(ByVal lf As Object) As LineSpec
    Dim spec As LineSpec
    spec.Visible = (lf.Visible <> msoFalse)
    spec.Weight = lf.Weight
    spec.DashStyle = lf.DashStyle
    spec.Color = lf.ForeColor.RGB
    spec.BeginArrowheadStyle = lf.BeginArrowheadStyle
    spec.EndArrowheadStyle = lf.EndArrowheadStyle
    ReadLine = spec
End Function

' 指定のある項目だけを適用する。矢印は線 (コネクタ) 以外の図形では例外になるので無視する
Private Sub ApplySpec(ByVal lf As Object, ByRef spec As LineSpec)
    If IsHidden(spec) Then lf.Visible = msoFalse: Exit Sub
    If Not IsEmpty(spec.Visible) Then lf.Visible = msoTrue
    If Not IsEmpty(spec.Weight) Then lf.Weight = spec.Weight
    If Not IsEmpty(spec.DashStyle) Then lf.DashStyle = spec.DashStyle
    If Not IsEmpty(spec.Color) Then lf.ForeColor.RGB = spec.Color
    On Error Resume Next
    If Not IsEmpty(spec.BeginArrowheadStyle) Then lf.BeginArrowheadStyle = spec.BeginArrowheadStyle
    If Not IsEmpty(spec.EndArrowheadStyle) Then lf.EndArrowheadStyle = spec.EndArrowheadStyle
End Sub

'========================
' 表示
'========================
' 線を消す指定か。Empty (指定なし) は False と等しく比較されてしまうため先に除く
Private Function IsHidden(ByRef spec As LineSpec) As Boolean
    If IsEmpty(spec.Visible) Then Exit Function
    IsHidden = Not spec.Visible
End Function

Private Function DescribeLine(ByVal lf As Object) As String
    DescribeLine = DescribeSpec(ReadLine(lf))
End Function

Private Function DescribeSpec(ByRef spec As LineSpec) As String
    Dim s As String
    If IsHidden(spec) Then DescribeSpec = "none": Exit Function

    s = CStr(Round(spec.Weight, 2)) & "pt " & DashName(spec.DashStyle) & " " & ColorLongToHex(spec.Color)
    Select Case True
        Case spec.BeginArrowheadStyle <> msoArrowheadNone And spec.EndArrowheadStyle <> msoArrowheadNone: s = s & " <->"
        Case spec.BeginArrowheadStyle <> msoArrowheadNone: s = s & " <-"
        Case spec.EndArrowheadStyle <> msoArrowheadNone: s = s & " ->"
    End Select
    DescribeSpec = s
End Function

Private Function DashName(ByVal style As Long) As String
    Select Case style
        Case msoLineSolid: DashName = "solid"
        Case msoLineRoundDot: DashName = "dot"
        Case msoLineSquareDot: DashName = "sdot"
        Case msoLineDash: DashName = "dash"
        Case msoLineDashDot: DashName = "dashdot"
        Case msoLineLongDash: DashName = "ldash"
        Case msoLineLongDashDot: DashName = "ldashdot"
        Case Else: DashName = "dash(" & style & ")"
    End Select
End Function
