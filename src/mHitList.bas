Attribute VB_Name = "mHitList"
Option Explicit
Option Private Module
' ヒット一覧 (セル・図形内テキストの位置のリスト) を n / N で順にたどる仕組み。
' find や lint が一覧を作り、Begin で移動モードを開始する。
' キー操作の受け口は mFind の fn / fp / fq (ランチャーからも呼べるコマンド)。

' ヒット1件 = Array(種別, 対象, 行, 列, 開始位置, 長さ, 説明)
'   種別 "cell": 対象は Range。開始位置・長さは未使用
'   種別 "shape": 対象は Shape。開始位置・長さはテキスト内のヒット位置 (その文字を選択する)
'   説明: ステータスバーに添える文字列 (例: "#REF!")。空でもよい
Public Const H_KIND As Long = 0
Public Const H_TARGET As Long = 1
Public Const H_ROW As Long = 2
Public Const H_COL As Long = 3
Public Const H_START As Long = 4
Public Const H_LEN As Long = 5
Public Const H_NOTE As Long = 6

Private gHits As Collection
Private gIndex As Long    ' 現在位置 (1 始まり)
Private gLabel As String  ' ステータスバーの見出し (find / lint)

'========================
' ヒットの作成・並べ替え
'========================
Public Function CellHit(ByVal cell As Range, Optional ByVal note As String) As Variant
    CellHit = Array("cell", cell, cell.Row, cell.Column, 0, 0, note)
End Function

Public Function ShapeHit(ByVal shp As Shape, ByVal topRow As Long, ByVal leftCol As Long, _
                         ByVal startPos As Long, ByVal length As Long, Optional ByVal note As String) As Variant
    ShapeHit = Array("shape", shp, topRow, leftCol, startPos, length, note)
End Function

' (行, 列, 開始位置) の昇順に並べた新しい Collection を返す (図形など少数向けの挿入ソート)
Public Function SortHits(ByVal hits As Collection) As Collection
    Dim result As New Collection
    Dim i As Long, j As Long
    Dim inserted As Boolean

    For i = 1 To hits.Count
        inserted = False
        For j = 1 To result.Count
            If ComesBefore(hits(i), result(j)) Then
                result.Add hits(i), Before:=j
                inserted = True
                Exit For
            End If
        Next j
        If Not inserted Then result.Add hits(i)
    Next i

    Set SortHits = result
End Function

' 位置順に並んだ cellHits と shapeHits を、位置順を保って result の末尾に追加する。
' 同じ位置ではセルを先にする
Public Sub MergeByPosition(ByVal cellHits As Collection, ByVal shapeHits As Collection, ByVal result As Collection)
    Dim i As Long, j As Long
    i = 1: j = 1
    Do While i <= cellHits.Count Or j <= shapeHits.Count
        If j > shapeHits.Count Then
            result.Add cellHits(i): i = i + 1
        ElseIf i > cellHits.Count Then
            result.Add shapeHits(j): j = j + 1
        ElseIf ComesBefore(shapeHits(j), cellHits(i)) Then
            result.Add shapeHits(j): j = j + 1
        Else
            result.Add cellHits(i): i = i + 1
        End If
    Loop
End Sub

Private Function ComesBefore(ByVal a As Variant, ByVal b As Variant) As Boolean
    If a(H_ROW) <> b(H_ROW) Then
        ComesBefore = a(H_ROW) < b(H_ROW)
    ElseIf a(H_COL) <> b(H_COL) Then
        ComesBefore = a(H_COL) < b(H_COL)
    Else
        ComesBefore = a(H_START) < b(H_START)
    End If
End Function

'========================
' 移動モード
'========================

' hits の1件目へ移動し、n / N / Esc を割り当てる。hits が空なら何もしない
Public Sub Begin(ByVal label As String, ByVal hits As Collection)
    Set gHits = hits
    gLabel = label
    gIndex = 1
    If gHits.Count = 0 Then EndSession: Exit Sub
    BindKeys
    ShowHit
End Sub

' 前回の一覧のまま移動モードに戻る。一覧がなければ False
Public Function ResumeSession() As Boolean
    If gHits Is Nothing Then Exit Function
    If gHits.Count = 0 Then Exit Function
    BindKeys
    ShowHit
    ResumeSession = True
End Function

Public Sub MoveBy(ByVal delta As Long)
    If gHits Is Nothing Then
        Application.StatusBar = "[hit] No results. Run find or lint first."
        Exit Sub
    End If
    If gHits.Count = 0 Then Exit Sub

    gIndex = ((gIndex - 1 + delta) Mod gHits.Count + gHits.Count) Mod gHits.Count + 1
    ShowHit
End Sub

Public Sub EndSession()
    Application.OnKey "n"
    Application.OnKey "+n"
    Application.OnKey "{ESC}"
End Sub

Private Sub BindKeys()
    Dim book As String
    book = "'" & ThisWorkbook.Name & "'!"
    Application.OnKey "n", book & "fn"
    Application.OnKey "+n", book & "fp"
    Application.OnKey "{ESC}", book & "fq"

    ' 日本語で入力した直後でも n / N が効くよう、IME をオフ (半角英数) にする
    mIme.ImeOffLater
End Sub

Private Sub ShowHit()
    Dim hit As Variant
    Dim ws As Worksheet
    Dim shp As Shape
    Dim cell As Range
    Dim where As String

    On Error GoTo STALE
    hit = gHits(gIndex)

    If hit(H_KIND) = "cell" Then
        Set cell = hit(H_TARGET)
        Set ws = cell.Worksheet
        ws.Parent.Activate
        ws.Activate
        Application.Goto cell, False
        where = ws.Name & "!" & cell.Address(False, False)
    Else
        Set shp = hit(H_TARGET)
        Set ws = ShapeSheet(shp)
        ws.Parent.Activate
        ws.Activate
        If Application.Intersect(ActiveWindow.VisibleRange, shp.TopLeftCell) Is Nothing Then
            Application.Goto shp.TopLeftCell, True
        End If
        On Error Resume Next
        shp.TextFrame2.TextRange.Characters(hit(H_START), hit(H_LEN)).Select
        If Err.Number <> 0 Then Err.Clear: shp.Select
        On Error GoTo STALE
        where = ws.Name & "!" & shp.Name
    End If

    If Len(hit(H_NOTE)) > 0 Then where = where & "  " & hit(H_NOTE)
    Application.StatusBar = "[" & gLabel & "] " & gIndex & "/" & gHits.Count & "  " & where & _
                            "   (n: 次 / N: 前 / Esc: 終了)"
    Exit Sub

STALE:
    Application.StatusBar = "[" & gLabel & "] The result is no longer valid. Run " & gLabel & " again. (E" & Err.Number & ")"
End Sub

' 図形 (グループ内の図形を含む) が載っているシート
Private Function ShapeSheet(ByVal shp As Shape) As Worksheet
    Dim p As Object
    Set p = shp.Parent
    Do While Not TypeOf p Is Worksheet
        Set p = p.Parent
    Loop
    Set ShapeSheet = p
End Function
