Attribute VB_Name = "mPolygon"
Option Explicit
'===============================================================================
' NAME
'     polygon, box - 選択範囲から閉じた図形を作成する
'
' SYNOPSIS
'     polygon
'     box
'
' DESCRIPTION
'     polygon は、選択範囲内の数値 (文字列としての数値でも可) が書き込まれたセルを
'     対象に、値の昇順で各セル左上の頂点を結んだ閉じた多角形 (フリーフォーム図形)
'     を作成する。
'     box は、選択範囲のセル境界に外接する矩形の図形を作成する。
'     いずれも線の太さは 0.5pt・線色は黒・塗りつぶしは白で作成される。
'
' NOTES
'     - polygon: 数値以外 (空欄・IsNumeric が False の文字列など) のセルは無視する。
'     - polygon: 同値のセルが複数ある場合は、シート走査順 (行→列) を維持する。
'     - polygon: 対象セルが3個未満の場合は何もしない。
'     - polygon, box とも、複数領域の選択 (Ctrl+クリックなど) には対応しない。
'===============================================================================

Private Type PolyNode
    x As Double
    y As Double
    v As Double
End Type

Public Sub polygon()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Areas.Count > 1 Then
        Application.StatusBar = "[polygon] 複数領域の選択には対応していません。"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Selection.Worksheet

    Dim nodes() As PolyNode
    Dim n As Long
    n = CollectNodes(Selection, nodes)
    If n < 3 Then
        Application.StatusBar = "[polygon] 数値セルが3個未満です。"
        Exit Sub
    End If

    SortNodesByValue nodes, n

    Dim shp As shape
    Set shp = BuildPolygonShape(ws, nodes, n)
    ApplyShapeStyle shp

    Application.StatusBar = "[polygon] " & n & " 頂点の多角形を作成しました。"
    Exit Sub

FINALLY:
    Application.StatusBar = "[polygon] E" & Err.Number & ": " & Err.Description
End Sub

Public Sub box()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Areas.Count > 1 Then
        Application.StatusBar = "[box] 複数領域の選択には対応していません。"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = Selection.Worksheet

    Dim shp As shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, Selection.Left, Selection.Top, Selection.Width, Selection.Height)
    ApplyShapeStyle shp

    Application.StatusBar = "[box] 矩形を作成しました。"
    Exit Sub

FINALLY:
    Application.StatusBar = "[box] E" & Err.Number & ": " & Err.Description
End Sub

' polygon, box に共通の図形スタイル (線0.5pt・黒線・白塗り) を適用する。
Private Sub ApplyShapeStyle(ByVal shp As shape)
    With shp
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .line.Visible = msoTrue
        .line.ForeColor.RGB = RGB(0, 0, 0)
        .line.Weight = 0.5
    End With
End Sub

' 選択範囲から数値セルを収集し、セル左上座標と値を nodes に格納する。
Private Function CollectNodes(ByVal rng As Range, ByRef nodes() As PolyNode) As Long
    Dim c As Range
    Dim cnt As Long
    ReDim nodes(1 To rng.Cells.Count)

    For Each c In rng.Cells
        If Not IsEmpty(c.Value) Then
            If IsNumeric(c.Value) Then
                cnt = cnt + 1
                nodes(cnt).x = c.Left
                nodes(cnt).y = c.Top
                nodes(cnt).v = CDbl(c.Value)
            End If
        End If
    Next c

    CollectNodes = cnt
End Function

' 値の昇順で安定ソート (挿入ソート)。同値は走査順を維持する。
Private Sub SortNodesByValue(ByRef nodes() As PolyNode, ByVal n As Long)
    Dim i As Long, j As Long
    Dim key As PolyNode

    For i = 2 To n
        key = nodes(i)
        j = i - 1
        Do While j >= 1
            If nodes(j).v > key.v Then
                nodes(j + 1) = nodes(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        nodes(j + 1) = key
    Next i
End Sub

' nodes の順に頂点を結んだフリーフォーム図形 (自動的に閉多角形になる) を作成する。
Private Function BuildPolygonShape(ByVal ws As Worksheet, ByRef nodes() As PolyNode, ByVal n As Long) As shape
    Dim fb As FreeformBuilder
    Dim i As Long

    Set fb = ws.Shapes.BuildFreeform(msoEditingCorner, nodes(1).x, nodes(1).y)
    For i = 2 To n
        fb.AddNodes msoSegmentLine, msoEditingCorner, nodes(i).x, nodes(i).y
    Next i
    ' ConvertToShape は始点に自動では戻らないため、最後に始点への辺を明示的に追加して閉じる。
    fb.AddNodes msoSegmentLine, msoEditingCorner, nodes(1).x, nodes(1).y

    Set BuildPolygonShape = fb.ConvertToShape
End Function
