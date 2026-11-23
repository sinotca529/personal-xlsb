Attribute VB_Name = "mPalette"
Option Explicit
'===============================================================================
' NAME
'     palette - Tailwind CSS の配色を実物の色見本として一覧表示する
'
' SYNOPSIS
'     palette
'
' DESCRIPTION
'     PERSONAL.XLSB 内に Palette シートを作成 (既にあれば作り直す) して表示する。
'     Tailwind の22色相 × 11段階に加え、black / white も別枠で表示する。
'     各セルには amber500 のような HUE[STEP] トークンがそのまま入っているので、
'     bg copy (bg c) でセルの色をコピーし、対象セルで bg paste (bg p) すれば、
'     色名を覚えていなくても貼り付けられる。
'
' NOTES
'     - LookupHex は Tailwind CSS v4 (MIT) の公開値 (oklch()) を、coloraide
'       (MINDE-chroma = CSS Color 4 相当のガマットマッピング) で sRGB に変換した
'       固定値。VBA 側での色計算は行わない。
'
' COMPLETIONS
'     palette         Tailwind 配色の一覧シートを表示
'===============================================================================

Private Const SHEET_NAME As String = "Palette"

Public Sub palette()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim ws As Worksheet
    Set ws = GetOrCreatePaletteSheet()
    FillPaletteSheet ws

    ThisWorkbook.Windows(1).Visible = True
    ws.Activate
    Exit Sub

FINALLY:
    Application.StatusBar = "[palette] E" & Err.Number & ": " & Err.Description
End Sub

Private Function GetOrCreatePaletteSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = SHEET_NAME
    End If

    ws.Cells.Clear
    Set GetOrCreatePaletteSheet = ws
End Function

Private Sub FillPaletteSheet(ByVal ws As Worksheet)
    Dim hueNames As Variant, steps As Variant
    hueNames = Array("red", "orange", "amber", "yellow", "lime", "green", "emerald", "teal", "cyan", "sky", "blue", "indigo", "violet", "purple", "fuchsia", "pink", "rose", "slate", "gray", "zinc", "neutral", "stone")
    steps = Array("50", "100", "200", "300", "400", "500", "600", "700", "800", "900", "950")

    Dim r As Long, c As Long
    Dim hexStr As String
    Dim clr As Long

    ' ヘッダー行 (段階)
    ws.Cells(1, 1).Value = "HUE"
    For c = 0 To UBound(steps)
        ws.Cells(1, c + 2).Value = steps(c)
        ws.Cells(1, c + 2).Font.Bold = True
    Next c

    For r = 0 To UBound(hueNames)
        ws.Cells(r + 2, 1).Value = hueNames(r)
        ws.Cells(r + 2, 1).Font.Bold = True

        For c = 0 To UBound(steps)
            hexStr = LookupHex(CStr(hueNames(r)), CStr(steps(c)))
            If Len(hexStr) > 0 Then
                If mColorUtil.TryParseHexColor(hexStr, clr) Then
                    With ws.Cells(r + 2, c + 2)
                        .Value = hueNames(r) & steps(c)
                        .Interior.Color = clr
                        .Font.Color = ContrastFontColor(clr)
                        .HorizontalAlignment = xlCenter
                    End With
                End If
            End If
        Next c
    Next r

    ' black / white (段階を持たない特別なキーワード) は1行空けて別枠で表示する
    Dim lastRow As Long
    lastRow = UBound(hueNames) + 2
    AddSpecialSwatch ws, lastRow + 2, "black", RGB(0, 0, 0)
    AddSpecialSwatch ws, lastRow + 3, "white", RGB(255, 255, 255)

    ws.Columns("A:L").ColumnWidth = 9
    ws.Rows(1).RowHeight = 16
End Sub

Private Sub AddSpecialSwatch(ByVal ws As Worksheet, ByVal row As Long, ByVal token As String, ByVal clr As Long)
    ws.Cells(row, 1).Value = token
    ws.Cells(row, 1).Font.Bold = True

    With ws.Cells(row, 2)
        .Value = token
        .Interior.Color = clr
        .Font.Color = ContrastFontColor(clr)
        .HorizontalAlignment = xlCenter
    End With
End Sub

' 背景色に対して読みやすい文字色 (黒/白) を返す
Private Function ContrastFontColor(ByVal bgColor As Long) As Long
    Dim r As Long, g As Long, b As Long
    Dim luminance As Double
    r = bgColor And &HFF&
    g = (bgColor \ &H100&) And &HFF&
    b = (bgColor \ &H10000) And &HFF&
    luminance = (r * 299 + g * 587 + b * 114) / 1000
    If luminance >= 150 Then
        ContrastFontColor = RGB(0, 0, 0)
    Else
        ContrastFontColor = RGB(255, 255, 255)
    End If
End Function

' =====================
' 配色テーブル (Tailwind CSS v4 より変換)
' =====================

Public Function LookupHex(ByVal name As String, ByVal step As String) As String
    Select Case name & "-" & step
        ' red
        Case "red-50": LookupHex = "fef2f2"
        Case "red-100": LookupHex = "ffe2e2"
        Case "red-200": LookupHex = "ffc9c9"
        Case "red-300": LookupHex = "ffa2a2"
        Case "red-400": LookupHex = "ff6467"
        Case "red-500": LookupHex = "fb2c36"
        Case "red-600": LookupHex = "e7000b"
        Case "red-700": LookupHex = "c10007"
        Case "red-800": LookupHex = "9f0712"
        Case "red-900": LookupHex = "82181a"
        Case "red-950": LookupHex = "460809"
        ' orange
        Case "orange-50": LookupHex = "fff7ed"
        Case "orange-100": LookupHex = "ffedd4"
        Case "orange-200": LookupHex = "ffd6a7"
        Case "orange-300": LookupHex = "ffb86a"
        Case "orange-400": LookupHex = "ff8904"
        Case "orange-500": LookupHex = "ff6900"
        Case "orange-600": LookupHex = "f54900"
        Case "orange-700": LookupHex = "ca3500"
        Case "orange-800": LookupHex = "9f2d00"
        Case "orange-900": LookupHex = "7e2a0c"
        Case "orange-950": LookupHex = "441306"
        ' amber
        Case "amber-50": LookupHex = "fffbeb"
        Case "amber-100": LookupHex = "fef3c6"
        Case "amber-200": LookupHex = "fee685"
        Case "amber-300": LookupHex = "ffd230"
        Case "amber-400": LookupHex = "ffba00"
        Case "amber-500": LookupHex = "fd9a00"
        Case "amber-600": LookupHex = "e17100"
        Case "amber-700": LookupHex = "bb4d00"
        Case "amber-800": LookupHex = "973c00"
        Case "amber-900": LookupHex = "7b3306"
        Case "amber-950": LookupHex = "461901"
        ' yellow
        Case "yellow-50": LookupHex = "fefce8"
        Case "yellow-100": LookupHex = "fef9c2"
        Case "yellow-200": LookupHex = "fff085"
        Case "yellow-300": LookupHex = "ffdf20"
        Case "yellow-400": LookupHex = "fcc800"
        Case "yellow-500": LookupHex = "efb100"
        Case "yellow-600": LookupHex = "d08700"
        Case "yellow-700": LookupHex = "a65f00"
        Case "yellow-800": LookupHex = "894b00"
        Case "yellow-900": LookupHex = "733e0a"
        Case "yellow-950": LookupHex = "432004"
        ' lime
        Case "lime-50": LookupHex = "f7fee7"
        Case "lime-100": LookupHex = "ecfcca"
        Case "lime-200": LookupHex = "d8f999"
        Case "lime-300": LookupHex = "bbf451"
        Case "lime-400": LookupHex = "9ae600"
        Case "lime-500": LookupHex = "7ccf00"
        Case "lime-600": LookupHex = "5ea500"
        Case "lime-700": LookupHex = "497d00"
        Case "lime-800": LookupHex = "3c6300"
        Case "lime-900": LookupHex = "35530e"
        Case "lime-950": LookupHex = "192e03"
        ' green
        Case "green-50": LookupHex = "f0fdf4"
        Case "green-100": LookupHex = "dcfce7"
        Case "green-200": LookupHex = "b9f8cf"
        Case "green-300": LookupHex = "7bf1a8"
        Case "green-400": LookupHex = "05df72"
        Case "green-500": LookupHex = "00c950"
        Case "green-600": LookupHex = "00a63e"
        Case "green-700": LookupHex = "008236"
        Case "green-800": LookupHex = "016630"
        Case "green-900": LookupHex = "0d542b"
        Case "green-950": LookupHex = "032e15"
        ' emerald
        Case "emerald-50": LookupHex = "ecfdf5"
        Case "emerald-100": LookupHex = "d0fae5"
        Case "emerald-200": LookupHex = "a4f4cf"
        Case "emerald-300": LookupHex = "5ee9b5"
        Case "emerald-400": LookupHex = "00d492"
        Case "emerald-500": LookupHex = "00bc7d"
        Case "emerald-600": LookupHex = "009966"
        Case "emerald-700": LookupHex = "007a55"
        Case "emerald-800": LookupHex = "006045"
        Case "emerald-900": LookupHex = "004f3b"
        Case "emerald-950": LookupHex = "002c22"
        ' teal
        Case "teal-50": LookupHex = "f0fdfa"
        Case "teal-100": LookupHex = "cbfbf1"
        Case "teal-200": LookupHex = "96f7e4"
        Case "teal-300": LookupHex = "46ecd5"
        Case "teal-400": LookupHex = "00d5be"
        Case "teal-500": LookupHex = "00bba7"
        Case "teal-600": LookupHex = "009689"
        Case "teal-700": LookupHex = "00786f"
        Case "teal-800": LookupHex = "005f5a"
        Case "teal-900": LookupHex = "0b4f4a"
        Case "teal-950": LookupHex = "022f2e"
        ' cyan
        Case "cyan-50": LookupHex = "ecfeff"
        Case "cyan-100": LookupHex = "cefafe"
        Case "cyan-200": LookupHex = "a2f4fd"
        Case "cyan-300": LookupHex = "53eafd"
        Case "cyan-400": LookupHex = "00d3f2"
        Case "cyan-500": LookupHex = "00b8db"
        Case "cyan-600": LookupHex = "0092b8"
        Case "cyan-700": LookupHex = "007595"
        Case "cyan-800": LookupHex = "005f78"
        Case "cyan-900": LookupHex = "104e64"
        Case "cyan-950": LookupHex = "053345"
        ' sky
        Case "sky-50": LookupHex = "f0f9ff"
        Case "sky-100": LookupHex = "dff2fe"
        Case "sky-200": LookupHex = "b8e6fe"
        Case "sky-300": LookupHex = "74d4ff"
        Case "sky-400": LookupHex = "00bcff"
        Case "sky-500": LookupHex = "00a6f4"
        Case "sky-600": LookupHex = "0084d1"
        Case "sky-700": LookupHex = "0069a8"
        Case "sky-800": LookupHex = "00598a"
        Case "sky-900": LookupHex = "024a70"
        Case "sky-950": LookupHex = "052f4a"
        ' blue
        Case "blue-50": LookupHex = "eff6ff"
        Case "blue-100": LookupHex = "dbeafe"
        Case "blue-200": LookupHex = "bedbff"
        Case "blue-300": LookupHex = "8ec5ff"
        Case "blue-400": LookupHex = "51a2ff"
        Case "blue-500": LookupHex = "2b7fff"
        Case "blue-600": LookupHex = "155dfc"
        Case "blue-700": LookupHex = "1447e6"
        Case "blue-800": LookupHex = "193cb8"
        Case "blue-900": LookupHex = "1c398e"
        Case "blue-950": LookupHex = "162456"
        ' indigo
        Case "indigo-50": LookupHex = "eef2ff"
        Case "indigo-100": LookupHex = "e0e7ff"
        Case "indigo-200": LookupHex = "c6d2ff"
        Case "indigo-300": LookupHex = "a3b3ff"
        Case "indigo-400": LookupHex = "7c86ff"
        Case "indigo-500": LookupHex = "615fff"
        Case "indigo-600": LookupHex = "4f39f6"
        Case "indigo-700": LookupHex = "432dd7"
        Case "indigo-800": LookupHex = "372aac"
        Case "indigo-900": LookupHex = "312c85"
        Case "indigo-950": LookupHex = "1e1a4d"
        ' violet
        Case "violet-50": LookupHex = "f5f3ff"
        Case "violet-100": LookupHex = "ede9fe"
        Case "violet-200": LookupHex = "ddd6ff"
        Case "violet-300": LookupHex = "c4b4ff"
        Case "violet-400": LookupHex = "a684ff"
        Case "violet-500": LookupHex = "8e51ff"
        Case "violet-600": LookupHex = "7f22fe"
        Case "violet-700": LookupHex = "7008e7"
        Case "violet-800": LookupHex = "5d0ec0"
        Case "violet-900": LookupHex = "4d179a"
        Case "violet-950": LookupHex = "2f0d68"
        ' purple
        Case "purple-50": LookupHex = "faf5ff"
        Case "purple-100": LookupHex = "f3e8ff"
        Case "purple-200": LookupHex = "e9d4ff"
        Case "purple-300": LookupHex = "dab2ff"
        Case "purple-400": LookupHex = "c27aff"
        Case "purple-500": LookupHex = "ad46ff"
        Case "purple-600": LookupHex = "9810fa"
        Case "purple-700": LookupHex = "8200db"
        Case "purple-800": LookupHex = "6e11b0"
        Case "purple-900": LookupHex = "59168b"
        Case "purple-950": LookupHex = "3c0366"
        ' fuchsia
        Case "fuchsia-50": LookupHex = "fdf4ff"
        Case "fuchsia-100": LookupHex = "fae8ff"
        Case "fuchsia-200": LookupHex = "f6cfff"
        Case "fuchsia-300": LookupHex = "f4a8ff"
        Case "fuchsia-400": LookupHex = "ed6aff"
        Case "fuchsia-500": LookupHex = "e12afb"
        Case "fuchsia-600": LookupHex = "c800de"
        Case "fuchsia-700": LookupHex = "a800b7"
        Case "fuchsia-800": LookupHex = "8a0194"
        Case "fuchsia-900": LookupHex = "721378"
        Case "fuchsia-950": LookupHex = "4b004f"
        ' pink
        Case "pink-50": LookupHex = "fdf2f8"
        Case "pink-100": LookupHex = "fce7f3"
        Case "pink-200": LookupHex = "fccee8"
        Case "pink-300": LookupHex = "fda5d5"
        Case "pink-400": LookupHex = "fb64b6"
        Case "pink-500": LookupHex = "f6339a"
        Case "pink-600": LookupHex = "e60076"
        Case "pink-700": LookupHex = "c6005c"
        Case "pink-800": LookupHex = "a3004c"
        Case "pink-900": LookupHex = "861043"
        Case "pink-950": LookupHex = "510424"
        ' rose
        Case "rose-50": LookupHex = "fff1f2"
        Case "rose-100": LookupHex = "ffe4e6"
        Case "rose-200": LookupHex = "ffccd3"
        Case "rose-300": LookupHex = "ffa1ad"
        Case "rose-400": LookupHex = "ff637e"
        Case "rose-500": LookupHex = "ff2056"
        Case "rose-600": LookupHex = "ec003f"
        Case "rose-700": LookupHex = "c70036"
        Case "rose-800": LookupHex = "a50036"
        Case "rose-900": LookupHex = "8b0836"
        Case "rose-950": LookupHex = "4d0218"
        ' slate
        Case "slate-50": LookupHex = "f8fafc"
        Case "slate-100": LookupHex = "f1f5f9"
        Case "slate-200": LookupHex = "e2e8f0"
        Case "slate-300": LookupHex = "cad5e2"
        Case "slate-400": LookupHex = "90a1b9"
        Case "slate-500": LookupHex = "62748e"
        Case "slate-600": LookupHex = "45556c"
        Case "slate-700": LookupHex = "314158"
        Case "slate-800": LookupHex = "1d293d"
        Case "slate-900": LookupHex = "0f172b"
        Case "slate-950": LookupHex = "020618"
        ' gray
        Case "gray-50": LookupHex = "f9fafb"
        Case "gray-100": LookupHex = "f3f4f6"
        Case "gray-200": LookupHex = "e5e7eb"
        Case "gray-300": LookupHex = "d1d5dc"
        Case "gray-400": LookupHex = "99a1af"
        Case "gray-500": LookupHex = "6a7282"
        Case "gray-600": LookupHex = "4a5565"
        Case "gray-700": LookupHex = "364153"
        Case "gray-800": LookupHex = "1e2939"
        Case "gray-900": LookupHex = "101828"
        Case "gray-950": LookupHex = "030712"
        ' zinc
        Case "zinc-50": LookupHex = "fafafa"
        Case "zinc-100": LookupHex = "f4f4f5"
        Case "zinc-200": LookupHex = "e4e4e7"
        Case "zinc-300": LookupHex = "d4d4d8"
        Case "zinc-400": LookupHex = "9f9fa9"
        Case "zinc-500": LookupHex = "71717b"
        Case "zinc-600": LookupHex = "52525c"
        Case "zinc-700": LookupHex = "3f3f46"
        Case "zinc-800": LookupHex = "27272a"
        Case "zinc-900": LookupHex = "18181b"
        Case "zinc-950": LookupHex = "09090b"
        ' neutral
        Case "neutral-50": LookupHex = "fafafa"
        Case "neutral-100": LookupHex = "f5f5f5"
        Case "neutral-200": LookupHex = "e5e5e5"
        Case "neutral-300": LookupHex = "d4d4d4"
        Case "neutral-400": LookupHex = "a1a1a1"
        Case "neutral-500": LookupHex = "737373"
        Case "neutral-600": LookupHex = "525252"
        Case "neutral-700": LookupHex = "404040"
        Case "neutral-800": LookupHex = "262626"
        Case "neutral-900": LookupHex = "171717"
        Case "neutral-950": LookupHex = "0a0a0a"
        ' stone
        Case "stone-50": LookupHex = "fafaf9"
        Case "stone-100": LookupHex = "f5f5f4"
        Case "stone-200": LookupHex = "e7e5e4"
        Case "stone-300": LookupHex = "d6d3d1"
        Case "stone-400": LookupHex = "a6a09b"
        Case "stone-500": LookupHex = "79716b"
        Case "stone-600": LookupHex = "57534d"
        Case "stone-700": LookupHex = "44403b"
        Case "stone-800": LookupHex = "292524"
        Case "stone-900": LookupHex = "1c1917"
        Case "stone-950": LookupHex = "0c0a09"
    End Select
End Function
