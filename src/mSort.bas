Attribute VB_Name = "mSort"
Option Explicit
'===============================================================================
' NAME
'     sort, uniq, hsort, huniq - 選択範囲を並べ替える・重複を取り除く
'
' SYNOPSIS
'     sort [COL][r]
'     uniq
'     hsort [ROW][r]
'     huniq
'
' DESCRIPTION
'     sort は選択範囲の各行を1件のレコードとして、COL 列目 (選択範囲の左端を 1 と
'     数える。既定 1) をキーに昇順で並べ替える。r を付けると降順。
'     数値・文字列の順序や大文字・小文字を区別しない点は Excel の並べ替えに従う。
'
'     uniq は選択範囲の行のうち、全列の値が一致する行を重複とみなし、最初の1行だけを
'     残して取り除く。残った行は選択範囲内で上に詰める。Unix の uniq と違い、
'     隣接していない重複も取り除く (Excel の「重複の削除」と同じ)。
'
'     例: sort       1列目で昇順
'         sort r     1列目で降順
'         sort 2     2列目で昇順
'         sort 2r    2列目で降順
'
'     hsort / huniq は横向き版。各列を1件のレコードとして扱う。
'     hsort は ROW 行目 (選択範囲の上端を 1 と数える。既定 1) をキーに左から右へ
'     並べ替える。huniq は全行の値が一致する列を重複とみなし、最初の1列だけを残して
'     左に詰め、空いた右側の列は値を消す (書式は残る)。huniq の比較は大文字・小文字を
'     区別しない。
'
' NOTES
'     - 対象は選択範囲内のセルのみ。範囲外のセルは動かない。
'     - 列全体などを選択した場合は、使用範囲 (UsedRange) と重なる部分だけを対象にする。
'     - 見出し行は扱わない。見出しを除いて選択すること。
'     - 複数領域の選択 (Ctrl+クリックなど) には対応しない。
'
' COMPLETIONS
'     sort            選択範囲の行を並べ替え
'     sort r          降順で並べ替え
'     uniq            重複した行を取り除く
'     hsort           選択範囲の列を並べ替え
'     hsort r         降順で並べ替え
'     huniq           重複した列を取り除く
'===============================================================================

Public Sub sort(Optional ByVal opt As String): SortRange "sort", opt, False: End Sub
Public Sub hsort(Optional ByVal opt As String): SortRange "hsort", opt, True: End Sub

' horizontal=False なら行を (列 KEY で) 上から下へ、True なら列を (行 KEY で) 左から右へ並べ替える
Private Sub SortRange(ByVal cmd As String, ByVal opt As String, ByVal horizontal As Boolean)
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim rng As Range
    Dim keyIdx As Long
    Dim descending As Boolean
    Dim keyRange As Range
    Dim unitName As String

    If Not TryGetTarget(cmd, rng) Then Exit Sub
    If Not ParseSortOption(LCase$(Trim$(opt)), keyIdx, descending) Then
        Application.StatusBar = "[" & cmd & "] Invalid option: " & opt
        Exit Sub
    End If

    unitName = IIf(horizontal, "Row", "Column")
    If keyIdx > IIf(horizontal, rng.Rows.Count, rng.Columns.Count) Then
        Application.StatusBar = "[" & cmd & "] " & unitName & " " & keyIdx & " is outside the selection."
        Exit Sub
    End If

    If horizontal Then
        Set keyRange = rng.Rows(keyIdx)
    Else
        Set keyRange = rng.Columns(keyIdx)
    End If

    rng.Sort Key1:=keyRange, _
             Order1:=IIf(descending, xlDescending, xlAscending), _
             Header:=xlNo, _
             Orientation:=IIf(horizontal, xlLeftToRight, xlTopToBottom), _
             MatchCase:=False

    Application.StatusBar = "[" & cmd & "] " & _
        IIf(horizontal, rng.Columns.Count & " columns", rng.Rows.Count & " rows") & _
        " by " & LCase$(unitName) & " " & keyIdx & IIf(descending, " (desc)", " (asc)")
    Exit Sub

FINALLY:
    Application.StatusBar = "[" & cmd & "] E" & Err.Number & ": " & Err.Description
End Sub

Public Sub uniq()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim rng As Range
    Dim cols() As Variant
    Dim i As Long

    If Not TryGetTarget("uniq", rng) Then Exit Sub

    ReDim cols(0 To rng.Columns.Count - 1)
    For i = 0 To UBound(cols)
        cols(i) = i + 1
    Next i

    ' Columns には配列を括弧で囲んで渡す (そのままだと実行時エラーになる)
    rng.RemoveDuplicates Columns:=(cols), Header:=xlNo

    Application.StatusBar = "[uniq] Removed duplicate rows in " & rng.Address(False, False)
    Exit Sub

FINALLY:
    Application.StatusBar = "[uniq] E" & Err.Number & ": " & Err.Description
End Sub

Public Sub huniq()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim rng As Range
    Dim seen As Object
    Dim key As String
    Dim c As Long, kept As Long

    If Not TryGetTarget("huniq", rng) Then Exit Sub

    ' 初出の列だけを左から順に詰めていく。書き込み先 (kept) は常に読み出し元 (c) 以下なので、
    ' まだ読んでいない列を上書きすることはない。
    Set seen = CreateObject("Scripting.Dictionary")
    For c = 1 To rng.Columns.Count
        key = ColumnKey(rng.Columns(c))
        If Not seen.Exists(key) Then
            seen.Add key, True
            kept = kept + 1
            If kept < c Then rng.Columns(c).Copy Destination:=rng.Columns(kept)
        End If
    Next c

    If kept < rng.Columns.Count Then
        rng.Columns(kept + 1).Resize(, rng.Columns.Count - kept).ClearContents
    End If

    Application.StatusBar = "[huniq] Removed " & (rng.Columns.Count - kept) & " duplicate column(s)."
    Exit Sub

FINALLY:
    Application.StatusBar = "[huniq] E" & Err.Number & ": " & Err.Description
End Sub

' 列の全セルの値を連結した比較用キー (大文字・小文字は区別しない)
Private Function ColumnKey(ByVal col As Range) As String
    Dim v As Variant
    Dim r As Long
    Dim parts() As String

    v = col.Value2
    If Not IsArray(v) Then
        ColumnKey = CellKey(v)
        Exit Function
    End If

    ReDim parts(1 To UBound(v, 1))
    For r = 1 To UBound(v, 1)
        parts(r) = CellKey(v(r, 1))
    Next r
    ColumnKey = Join(parts, vbNullChar)
End Function

' 型と値から作るセル1つ分のキー (数値の 1 と文字列の "1" を区別する)
Private Function CellKey(ByVal v As Variant) As String
    CellKey = VarType(v) & ":" & LCase$(CStr(v))
End Function

' 選択範囲を対象範囲として返す。単一領域でなければ失敗。使用範囲の外側は除く。
Private Function TryGetTarget(ByVal cmd As String, ByRef outRng As Range) As Boolean
    Dim rng As Range

    If TypeName(Selection) <> "Range" Then Exit Function
    If Selection.Areas.Count > 1 Then
        Application.StatusBar = "[" & cmd & "] Multiple areas are not supported."
        Exit Function
    End If

    Set rng = Application.Intersect(Selection, Selection.Worksheet.UsedRange)
    If rng Is Nothing Then Exit Function

    Set outRng = rng
    TryGetTarget = True
End Function

' opt ("", "r", "2", "2r" など) を列番号と降順フラグに分解する
Private Function ParseSortOption(ByVal opt As String, ByRef outCol As Long, ByRef outDesc As Boolean) As Boolean
    Dim digits As String

    outDesc = False
    If Right$(opt, 1) = "r" Then
        outDesc = True
        opt = Left$(opt, Len(opt) - 1)
    End If

    digits = opt
    If Len(digits) = 0 Then
        outCol = 1
    ElseIf digits Like String$(Len(digits), "#") Then
        outCol = CLng(digits)
        If outCol < 1 Then Exit Function
    Else
        Exit Function
    End If

    ParseSortOption = True
End Function
