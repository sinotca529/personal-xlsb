Attribute VB_Name = "mColorUtil"
' == mColorUtil

Public Function TryParseHexColor(ByVal s As String, ByRef outColor As Long) As Boolean
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    If Len(s) <> 6 Then Exit Function
    If Not IsHexStr(s) Then Exit Function

    Dim r As Long, g As Long, b As Long
    On Error GoTo Fail
    r = CLng("&H" & Mid$(s, 1, 2))
    g = CLng("&H" & Mid$(s, 3, 2))
    b = CLng("&H" & Mid$(s, 5, 2))
    outColor = RGB(r, g, b)
    TryParseHexColor = True
    Exit Function
Fail:
    TryParseHexColor = False
End Function

Public Function ColorLongToHex(ByVal c As Long) As String
    Dim r As Long, g As Long, b As Long
    r = (c And &HFF&)
    g = (c \ &H100&) And &HFF&
    b = (c \ &H10000) And &HFF&
    ColorLongToHex = "#" & TwoHex(r) & TwoHex(g) & TwoHex(b)
End Function

' =====================

Private Function IsHexStr(ByVal s As String) As Boolean
    Dim i As Long, ch As Integer
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        Select Case ch
            Case 48 To 57, 65 To 70, 97 To 102 ' 0-9 A-F a-f
            Case Else: Exit Function
        End Select
    Next
    IsHexStr = True
End Function

' HUE[STEP] トークン (例: amber500, blue50) を Tailwind CSS の配色テーブル (mPalette.bas)
' で解決する。STEP を省略した場合は既定段階 500 にフォールバックする。
Public Function TryParseHueToken(ByVal token As String, ByRef outColor As Long) As Boolean
    Dim name As String
    Dim lvl As String
    Dim hexStr As String
    Dim i As Long

    token = LCase$(Trim$(token))

    Select Case token
        Case "black"
            outColor = RGB(0, 0, 0)
            TryParseHueToken = True
            Exit Function
        Case "white"
            outColor = RGB(255, 255, 255)
            TryParseHueToken = True
            Exit Function
    End Select

    ' 末尾の数字 (段階) を切り出す。
    ' VBA の And は短絡評価しないため、i >= 1 と Mid$ を1つの条件に
    ' まとめると i = 0 のときに Mid$(token, 0, 1) が実行時エラー5になる。
    i = Len(token)
    Do While i >= 1
        ' Not は Like より結合が強いため、括弧が無いと型不一致 (エラー13) になる。
        If Not (Mid$(token, i, 1) Like "[0-9]") Then Exit Do
        i = i - 1
    Loop
    name = Left$(token, i)
    lvl = Mid$(token, i + 1)
    If Len(lvl) = 0 Then lvl = "500"

    hexStr = mPalette.LookupHex(name, lvl)
    If Len(hexStr) = 0 Then Exit Function

    TryParseHueToken = TryParseHexColor(hexStr, outColor)
End Function

Private Function TwoHex(ByVal v As Long) As String
    TwoHex = Right$("0" & Hex$(v), 2)
End Function
