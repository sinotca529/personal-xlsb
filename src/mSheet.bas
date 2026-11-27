Attribute VB_Name = "mSheet"
Option Explicit
'===============================================================================
' NAME
'     ff - シート名のあいまい検索でシートへ移動する
'
' SYNOPSIS
'     ff TEXT
'
' DESCRIPTION
'     アクティブなブックの表示中のシート (グラフシートを含む) から、名前が TEXT に
'     あいまい一致するものを探し、いちばん良く一致するシートへ移動する。
'
'     ランチャーで "ff " に続けて入力すると、一致するシートが良い順に候補として出る。
'     Tab / 上下キーで選んで Enter で移動する。選ばずに Enter を押すと、候補の先頭
'     (いちばん良く一致するシート) へ移動する。
'
'     あいまい一致の規則:
'     - TEXT の文字が、その順でシート名に含まれていれば一致とする (間に別の文字が
'       あってよい)。例: "設書" は "設計書" に一致する。TEXT の空白は無視する。
'     - 大文字・小文字、全角・半角は区別しない。
'     - 並び順: 名前が TEXT と完全に一致するシートが最優先。ほかは点数の高い順で、
'       連続して一致する・名前の先頭や区切り (空白 _ - など) の直後で一致すると
'       点数が高い。点数が同じなら名前の短い順、それも同じならタブの並び順。
'
' NOTES
'     - 非表示のシートは対象外。
'     - 空白を含む名前を候補から選ぶと、ランチャーが "..." で囲んで入力する。
'
' COMPLETIONS
'     ff              シート名のあいまい検索でシートへ移動
'     ff @mSheet.SheetCandidates
'===============================================================================

' 区切りの直後の一致に与える加点の対象 (名前の単語の切れ目とみなす文字)
Private Const SEPARATORS As String = " _-.・()（）[]「」"

Public Sub ff(Optional ByVal query As String)
    On Error GoTo FINALLY

    If Len(query) = 0 Then
        Application.StatusBar = "[ff] Please specify part of a sheet name."
        Exit Sub
    End If
    If ActiveWorkbook Is Nothing Then
        Application.StatusBar = "[ff] No workbook is open."
        Exit Sub
    End If

    Dim ranked As Collection
    Set ranked = RankSheets(ActiveWorkbook, query)
    If ranked.Count = 0 Then
        Application.StatusBar = "[ff] No sheet matches: " & query
        Exit Sub
    End If

    ranked(1).Activate
    Application.StatusBar = "[ff] " & ranked(1).Name
    Exit Sub

FINALLY:
    Application.StatusBar = "[ff] E" & Err.Number & ": " & Err.Description
End Sub

' ランチャーの補完候補 (COMPLETIONS の @mSheet.SheetCandidates)。
' Array(シート名, 説明) の配列を良い順に返す。query が空なら表示中の全シートをタブ順に返す
Public Function SheetCandidates(ByVal query As String) As Variant
    Dim ranked As Collection
    Dim pairs() As Variant
    Dim i As Long

    SheetCandidates = Array()
    If ActiveWorkbook Is Nothing Then Exit Function
    Set ranked = RankSheets(ActiveWorkbook, query)
    If ranked.Count = 0 Then Exit Function

    ReDim pairs(0 To ranked.Count - 1)
    For i = 1 To ranked.Count
        pairs(i - 1) = Array(ranked(i).Name, "")
    Next i
    SheetCandidates = pairs
End Function

' wb の表示中のシートのうち query にあいまい一致するものを、良い順に並べる
Public Function RankSheets(ByVal wb As Workbook, ByVal query As String) As Collection
    Dim result As New Collection
    Dim scores As New Collection
    Dim sht As Object
    Dim score As Long
    Dim i As Long, pos As Long

    For Each sht In wb.Sheets
        If sht.Visible = xlSheetVisible Then
            score = FuzzyScore(sht.Name, query)
            If score >= 0 Then
                ' 安定な挿入ソート (点数の高い順、同点なら名前の短い順、それも同じならタブ順)
                pos = result.Count + 1
                For i = 1 To result.Count
                    If Better(score, Len(sht.Name), scores(i), Len(result(i).Name)) Then pos = i: Exit For
                Next i
                If pos > result.Count Then
                    result.Add sht
                    scores.Add score
                Else
                    result.Add sht, Before:=pos
                    scores.Add score, Before:=pos
                End If
            End If
        End If
    Next sht

    Set RankSheets = result
End Function

' text が query にあいまい一致すれば 0 以上の点数、一致しなければ -1。
' 完全一致 (大文字・小文字、全角・半角の違いを除く) は最高点
Public Function FuzzyScore(ByVal text As String, ByVal query As String) As Long
    Const EXACT As Long = 1000000
    Dim t As String, q As String
    Dim best As Long, score As Long
    Dim start As Long

    t = Normalize(text)
    q = Replace(Normalize(query), " ", "")
    If Len(q) = 0 Then FuzzyScore = 0: Exit Function
    If t = Normalize(query) Then FuzzyScore = EXACT: Exit Function

    ' q の先頭の文字が現れる位置ごとに、そこから貪欲に照合して最も良い点数を取る
    best = -1
    start = InStr(1, t, Left$(q, 1))
    Do While start > 0
        score = GreedyScore(t, q, start)
        If score > best Then best = score
        start = InStr(start + 1, t, Left$(q, 1))
    Loop
    FuzzyScore = best
End Function

' t の start から q の文字を順に探して点数を付ける。最後まで見つからなければ -1
Private Function GreedyScore(ByVal t As String, ByVal q As String, ByVal start As Long) As Long
    Dim i As Long, pos As Long, prev As Long
    Dim score As Long

    prev = 0
    pos = start
    For i = 1 To Len(q)
        If i > 1 Then pos = InStr(prev + 1, t, Mid$(q, i, 1))
        If pos = 0 Then GreedyScore = -1: Exit Function

        score = score + 1
        If pos = prev + 1 And i > 1 Then score = score + 5      ' 直前の文字に続いて一致
        If pos = 1 Then
            score = score + 8                                    ' 名前の先頭
        ElseIf InStr(SEPARATORS, Mid$(t, pos - 1, 1)) > 0 Then
            score = score + 4                                    ' 区切りの直後
        End If
        prev = pos
    Next i
    GreedyScore = score
End Function

' 比較用に、全角英数を半角に、大文字を小文字にそろえる
Private Function Normalize(ByVal s As String) As String
    Normalize = LCase$(StrConv(s, vbNarrow))
End Function

Private Function Better(ByVal scoreA As Long, ByVal lenA As Long, ByVal scoreB As Long, ByVal lenB As Long) As Boolean
    If scoreA <> scoreB Then
        Better = scoreA > scoreB
    Else
        Better = lenA < lenB
    End If
End Function
