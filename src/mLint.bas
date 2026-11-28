Attribute VB_Name = "mLint"
Option Explicit
'===============================================================================
' NAME
'     lint - 数式のエラーと外部ブックへの参照を調べ、該当箇所を順に移動する
'
' SYNOPSIS
'     lint
'
' DESCRIPTION
'     アクティブなブックを調べ、問題のあるセルの一覧を作って1件目へ移動する。
'     移動は find と共通で、n (次へ) / N (前へ) / Esc (終了)。ステータスバーに
'     セルの位置と問題の内容 (例: #REF!、外部参照 [Book2.xlsx]) を表示する。
'
'     調べる内容:
'         - エラー値のセル (#REF! #N/A #DIV/0! #NAME? #VALUE! など)
'         - 数式に #REF! を含むセル (IFERROR などで表示上はエラーにならないもの)
'         - 外部ブックを参照する数式 ([Book.xlsx]Sheet1!A1 など)
'         - 外部ブックを参照する入力規則 (リストの参照先など)
'         - ほかのファイルを指すハイパーリンク (セル・図形。Web とメールのリンクは除く)
'         - 外部ブックのマクロが登録された図形・ボタン ('Book2.xlsm'!Macro など)
'         - 外部ブックを参照するグラフの系列
'         - 外部ブックを参照するリンクされた画像、画像ファイルへのリンク
'         - 参照先に #REF! や外部ブックを含む名前定義
'         - Excel が把握している外部リンク (Excel の「リンクの編集」に出るもの)
'
'     セルと図形の問題は、位置 (図形は左上のセル) の順に並べて移動できるようにする。
'     非表示シートの問題、非表示の図形、グラフシート、名前定義、Excel が把握している
'     外部リンクは移動できないため、件数と一覧を最初にダイアログでまとめて表示する
'     (該当がなければ表示しない)。
'
' NOTES
'     - テーブルの構造化参照 (Table1[列]) は外部参照とみなさない。
'     - 入力規則は、同じ規則が続く範囲 (SpecialCells の1つの領域) ごとに1件だけ出す。
'       2000 セルを超える領域は先頭のセルだけを調べる。
'     - 条件付き書式の中身は調べない。
'
' COMPLETIONS
'     lint            数式のエラーと外部参照を調べる
'===============================================================================

Private Const MAX_LIST As Long = 15 ' ダイアログに並べる件数の上限 (項目ごと)

Public Sub lint()
    Dim booster As cBooster: Set booster = New cBooster
    On Error GoTo FINALLY

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim hits As New Collection
    Dim sheetHits As Collection
    Dim hit As Variant
    Dim hiddenCount As Long
    Dim hiddenSheets As New Collection
    Dim nameIssues As Collection
    Dim links As Collection
    Dim unreachable As New Collection ' 移動できない場所の問題 (非表示の図形・グラフシート)
    Dim extRe As Object
    Dim summary As String

    Set wb = ActiveWorkbook
    Set extRe = ExternalRefRegExp()

    For Each ws In wb.Worksheets
        Set sheetHits = ScanSheet(ws, extRe, unreachable)
        If ws.Visible = xlSheetVisible Then
            For Each hit In sheetHits
                hits.Add hit
            Next hit
        ElseIf sheetHits.Count > 0 Then
            hiddenCount = hiddenCount + sheetHits.Count
            hiddenSheets.Add ws.Name & " (" & sheetHits.Count & ")"
        End If
    Next ws

    AddChartSheetIssues wb, extRe, unreachable
    Set nameIssues = ScanNames(wb, extRe)
    Set links = ExternalLinks(wb)
    Set booster = Nothing ' 画面更新を戻してから表示・移動する

    summary = BuildSummary(hits.Count, hiddenCount, hiddenSheets, unreachable, nameIssues, links)
    If Len(summary) > 0 Then MsgBox summary, vbExclamation, "lint"

    If hits.Count = 0 Then
        mHitList.EndSession
        If Len(summary) = 0 Then
            Application.StatusBar = "[lint] No problems found."
        Else
            Application.StatusBar = "[lint] No problem cells in visible sheets."
        End If
        Exit Sub
    End If
    mHitList.Begin "lint", hits
    Exit Sub

FINALLY:
    Application.StatusBar = "[lint] E" & Err.Number & ": " & Err.Description
End Sub

'========================
' シート
'========================

' シートの問題 (セル・入力規則・ハイパーリンク・図形) を位置の順に集める。
' 移動できない問題 (非表示の図形) は unreachable に足す
Public Function ScanSheet(ByVal ws As Worksheet, ByVal extRe As Object, ByVal unreachable As Collection) As Collection
    Dim result As New Collection
    Dim others As New Collection
    Dim shp As Shape

    AddValidationHits ws, extRe, others
    AddHyperlinkHits ws, others
    For Each shp In ws.Shapes
        AddShapeHits ws, shp, extRe, others, unreachable
    Next shp

    mHitList.MergeByPosition ScanCells(ws, extRe), mHitList.SortHits(others), result
    Set ScanSheet = result
End Function

'========================
' セル
'========================

' シートの問題セルを行→列の順に集める。値と数式は配列で一括に読む
Private Function ScanCells(ByVal ws As Worksheet, ByVal extRe As Object) As Collection
    Dim result As New Collection
    Dim used As Range
    Dim vals As Variant, fmls As Variant
    Dim r As Long, c As Long
    Dim note As String
    Dim fml As String

    Set ScanCells = result
    Set used = ws.UsedRange
    vals = AsArray(used.Value2)
    fmls = AsArray(used.Formula)

    For r = 1 To UBound(vals, 1)
        For c = 1 To UBound(vals, 2)
            note = ""
            If IsError(vals(r, c)) Then note = ErrorText(vals(r, c))

            If VarType(fmls(r, c)) = vbString Then
                fml = fmls(r, c)
                If Left$(fml, 1) = "=" Then
                    If InStr(fml, "#REF!") > 0 And note <> "#REF!" Then note = JoinNote(note, "数式に #REF!")
                    If extRe.Test(fml) Then note = JoinNote(note, "外部参照 " & ExternalBook(extRe, fml))
                End If
            End If

            If Len(note) > 0 Then result.Add mHitList.CellHit(used.Cells(r, c), note)
        Next c
    Next r
End Function

'========================
' 入力規則・ハイパーリンク
'========================

' 外部ブックを参照する入力規則。SpecialCells の領域ごとに最初の1件だけ
Private Sub AddValidationHits(ByVal ws As Worksheet, ByVal extRe As Object, ByVal found As Collection)
    Const MAX_CELLS_PER_AREA As Double = 2000
    Dim validated As Range, area As Range, cell As Range

    On Error Resume Next ' 入力規則のセルがなければ SpecialCells は例外になる
    Set validated = ws.Cells.SpecialCells(xlCellTypeAllValidation)
    On Error GoTo 0
    If validated Is Nothing Then Exit Sub

    For Each area In validated.Areas
        ' 列全体などの広い領域は、セル数が Long に収まらないことがあるので行数 × 列数で比べる
        If CDbl(area.Rows.Count) * area.Columns.Count > MAX_CELLS_PER_AREA Then
            AddValidationHit area.Cells(1, 1), extRe, found
        Else
            For Each cell In area.Cells
                If AddValidationHit(cell, extRe, found) Then Exit For
            Next cell
        End If
    Next area
End Sub

' cell の入力規則が外部ブックを参照していれば found に足して True
Private Function AddValidationHit(ByVal cell As Range, ByVal extRe As Object, ByVal found As Collection) As Boolean
    Dim f As String
    On Error Resume Next ' 入力規則の種類によっては Formula1 / Formula2 が例外になる
    f = cell.Validation.Formula1
    f = f & " " & cell.Validation.Formula2
    On Error GoTo 0
    If Not extRe.Test(f) Then Exit Function
    found.Add mHitList.CellHit(cell, "入力規則 外部参照 " & ExternalBook(extRe, f))
    AddValidationHit = True
End Function

' ほかのファイルを指すハイパーリンク (セル・図形)
Private Sub AddHyperlinkHits(ByVal ws As Worksheet, ByVal found As Collection)
    Dim hl As Hyperlink
    Dim addr As String, note As String
    Dim shp As Shape

    For Each hl In ws.Hyperlinks
        addr = ""
        On Error Resume Next ' 壊れたハイパーリンクはプロパティの読み出しが例外になることがある
        addr = hl.Address
        On Error GoTo 0
        If IsFileLink(addr) Then
            note = "リンク先のファイル: " & addr
            If hl.Type = msoHyperlinkShape Then
                Set shp = hl.Shape
                found.Add ShapeHitAt(shp, note)
            Else
                found.Add mHitList.CellHit(hl.Range.Cells(1, 1), note)
            End If
        End If
    Next hl
End Sub

' ハイパーリンクのアドレスがファイルを指すか (空 = ブック内のリンク、Web とメールは除く)
Public Function IsFileLink(ByVal addr As String) As Boolean
    Dim a As String
    a = LCase$(Trim$(addr))
    If Len(a) = 0 Then Exit Function
    If a Like "http://*" Or a Like "https://*" Or a Like "mailto:*" Or a Like "ftp://*" Then Exit Function
    IsFileLink = True
End Function

'========================
' 図形・グラフ
'========================

' 外部ブックとつながる図形 (グループの中も調べる)。非表示の図形は unreachable に足す
Private Sub AddShapeHits(ByVal ws As Worksheet, ByVal shp As Shape, ByVal extRe As Object, _
                         ByVal found As Collection, ByVal unreachable As Collection)
    Dim child As Shape
    Dim note As String

    If IsGroup(shp) Then
        For Each child In shp.GroupItems
            AddShapeHits ws, child, extRe, found, unreachable
        Next child
        ' グループそのものにもマクロを登録できるので、続けて調べる
    End If

    note = AddNote(note, MacroNote(ws.Parent, shp))
    note = AddNote(note, ChartNote(shp, extRe))
    note = AddNote(note, PictureNote(shp, extRe))
    If Len(note) = 0 Then Exit Sub

    If Not shp.Visible Then
        unreachable.Add ws.Name & "!" & shp.Name & " (非表示の図形)  " & note
        Exit Sub
    End If
    found.Add ShapeHitAt(shp, note)
End Sub

Private Function IsGroup(ByVal shp As Shape) As Boolean
    On Error Resume Next
    IsGroup = (shp.Type = msoGroup)
End Function

' 図形そのものを選ぶ移動先 (並び順は左上のセル。取れなければ先頭に置く)
Private Function ShapeHitAt(ByVal shp As Shape, ByVal note As String) As Variant
    Dim topRow As Long, leftCol As Long
    On Error Resume Next
    topRow = shp.TopLeftCell.Row
    leftCol = shp.TopLeftCell.Column
    On Error GoTo 0
    ShapeHitAt = mHitList.ShapeHit(shp, topRow, leftCol, 0, 0, note)
End Function

' 図形に登録されたマクロがほかのブックのものなら、その説明
Private Function MacroNote(ByVal wb As Workbook, ByVal shp As Shape) As String
    Dim act As String, book As String
    Dim p As Long

    On Error Resume Next ' マクロを登録できない図形は OnAction が例外になる
    act = shp.OnAction
    On Error GoTo 0
    p = InStrRev(act, "!")
    If p = 0 Then Exit Function

    book = Replace(Left$(act, p - 1), "'", "")
    If StrComp(book, wb.Name, vbTextCompare) = 0 Or StrComp(book, wb.FullName, vbTextCompare) = 0 Then Exit Function
    MacroNote = "マクロの登録先 [" & book & "]"
End Function

' グラフの図形の系列が外部ブックを参照していれば、その説明
Private Function ChartNote(ByVal shp As Shape, ByVal extRe As Object) As String
    Dim isChart As Boolean
    On Error Resume Next
    isChart = shp.HasChart
    On Error GoTo 0
    If isChart Then ChartNote = SeriesNote(shp.Chart, extRe)
End Function

' グラフの系列の数式に外部参照があれば、その説明
' (On Error Resume Next の中で If の条件が例外になると If の中へ進んでしまうので、
'  条件は先に変数へ入れる。例外なら変数は初期値の False のまま)
Private Function SeriesNote(ByVal cht As Chart, ByVal extRe As Object) As String
    Dim i As Long, n As Long
    Dim f As String
    Dim matched As Boolean

    On Error Resume Next ' データのない系列などは Formula が例外になる
    n = cht.SeriesCollection.Count
    For i = 1 To n
        f = ""
        matched = False
        f = cht.SeriesCollection(i).Formula
        matched = extRe.Test(f)
        If matched Then
            SeriesNote = "グラフの系列 外部参照 " & ExternalBook(extRe, f)
            Exit Function
        End If
    Next i
End Function

' リンクされた画像 (セル範囲の画像・画像ファイル) がほかのブック・ファイルを指していれば、その説明
Private Function PictureNote(ByVal shp As Shape, ByVal extRe As Object) As String
    Dim f As String
    Dim matched As Boolean, linkedFile As Boolean

    On Error Resume Next ' 画像以外の図形はどちらのプロパティも例外になる (条件は先に変数へ。SeriesNote を参照)
    f = shp.DrawingObject.Formula
    matched = extRe.Test(f)
    If matched Then
        PictureNote = "リンクされた画像 外部参照 " & ExternalBook(extRe, f)
        Exit Function
    End If

    linkedFile = (shp.Type = msoLinkedPicture)
    If linkedFile Then
        f = ""
        f = shp.LinkFormat.SourceFullName
        PictureNote = "画像ファイルへのリンク: " & f
    End If
End Function

' グラフシートの系列の外部参照 (グラフシートは移動先にできないので unreachable に足す)
Private Sub AddChartSheetIssues(ByVal wb As Workbook, ByVal extRe As Object, ByVal unreachable As Collection)
    Dim cht As Chart
    Dim note As String
    For Each cht In wb.Charts
        note = SeriesNote(cht, extRe)
        If Len(note) > 0 Then unreachable.Add cht.Name & " (グラフシート)  " & note
    Next cht
End Sub

' 説明を ", " でつなぐ (空の説明は足さない)
Private Function AddNote(ByVal note As String, ByVal more As String) As String
    If Len(more) = 0 Then
        AddNote = note
    Else
        AddNote = JoinNote(note, more)
    End If
End Function

' 単一セルの Value2 / Formula はスカラーで返るので 1x1 の配列にそろえる
Private Function AsArray(ByVal v As Variant) As Variant
    If IsArray(v) Then AsArray = v: Exit Function
    Dim one(1 To 1, 1 To 1) As Variant
    one(1, 1) = v
    AsArray = one
End Function

Private Function JoinNote(ByVal note As String, ByVal more As String) As String
    JoinNote = IIf(Len(note) = 0, more, note & ", " & more)
End Function

' エラー値の表示文字列 (v はエラー値であること)
Private Function ErrorText(ByVal v As Variant) As String
    Select Case True
        Case v = CVErr(xlErrNull): ErrorText = "#NULL!"
        Case v = CVErr(xlErrDiv0): ErrorText = "#DIV/0!"
        Case v = CVErr(xlErrValue): ErrorText = "#VALUE!"
        Case v = CVErr(xlErrRef): ErrorText = "#REF!"
        Case v = CVErr(xlErrName): ErrorText = "#NAME?"
        Case v = CVErr(xlErrNum): ErrorText = "#NUM!"
        Case v = CVErr(xlErrNA): ErrorText = "#N/A"
        Case Else: ErrorText = "#ERROR"
    End Select
End Function

'========================
' 外部参照の判定
'========================

' 外部ブックへの参照に一致する正規表現。
'   [Book.xlsx]Sheet1!A1 / 'C:\dir\[Book.xlsx]Sheet 1'!A1 / [1]Sheet1!A1   … 角括弧のあとにシート名と !
'   Book.xlsx!Name / 'Book.xlsx'!Name                                    … ブック名.xl* のあとに !
' テーブルの構造化参照 (Table1[列]) は、角括弧のあとに演算子などを挟まずに ! が続くことはないので一致しない
Private Function ExternalRefRegExp() As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = "\[([^\[\]]+)\](?:[^'!\s()+\-*/,;=<>&^\[\]]+!|[^'\[\]]*'!)" & _
                 "|([^'!\s()+\-*/,;=<>&^\[\]]+\.xl(?:sx|sm|sb|am|tx|tm|s|a|t))'?!"
    re.IgnoreCase = True
    re.Global = False
    Set ExternalRefRegExp = re
End Function

' 数式中の最初の外部参照のブック名 (例: [Book2.xlsx])
Private Function ExternalBook(ByVal extRe As Object, ByVal fml As String) As String
    Dim m As Object
    Set m = extRe.Execute(fml)(0)
    If Not IsEmpty(m.SubMatches(0)) And Len(m.SubMatches(0)) > 0 Then
        ExternalBook = "[" & m.SubMatches(0) & "]"
    Else
        ExternalBook = "[" & m.SubMatches(1) & "]"
    End If
End Function

'========================
' 名前定義・外部リンク
'========================
Private Function ScanNames(ByVal wb As Workbook, ByVal extRe As Object) As Collection
    Dim result As New Collection
    Dim nm As Name
    Dim refText As String

    For Each nm In wb.Names
        ' Excel が内部で使う名前 (_xlfn. など) は対象外
        If Not (nm.Name Like "_xl*" Or nm.Name Like "*!_xl*") Then
            refText = ""
            On Error Resume Next
            refText = nm.RefersTo
            On Error GoTo 0
            If InStr(refText, "#REF!") > 0 Then
                result.Add nm.Name & "  #REF!"
            ElseIf extRe.Test(refText) Then
                result.Add nm.Name & "  外部参照 " & ExternalBook(extRe, refText)
            End If
        End If
    Next nm

    Set ScanNames = result
End Function

Private Function ExternalLinks(ByVal wb As Workbook) As Collection
    Dim result As New Collection
    Dim lk As Variant
    Dim i As Long

    lk = wb.LinkSources(xlExcelLinks)
    If IsArray(lk) Then
        For i = LBound(lk) To UBound(lk)
            result.Add lk(i)
        Next i
    End If
    Set ExternalLinks = result
End Function

'========================
' まとめ
'========================

' 移動できない問題 (非表示シート・非表示の図形やグラフシート・名前・外部リンク) のまとめ。
' 該当がなければ空文字列
Private Function BuildSummary(ByVal cellCount As Long, ByVal hiddenCount As Long, ByVal hiddenSheets As Collection, _
                              ByVal unreachable As Collection, ByVal nameIssues As Collection, _
                              ByVal links As Collection) As String
    Dim s As String

    If hiddenCount = 0 And unreachable.Count = 0 And nameIssues.Count = 0 And links.Count = 0 Then Exit Function

    s = "表示中のシートの問題 (セル・図形): " & cellCount & " 件"
    If cellCount > 0 Then s = s & " (OK のあと n / N で移動)"
    s = s & vbLf
    If hiddenCount > 0 Then
        s = s & vbLf & "非表示シートの問題: " & hiddenCount & " 件" & vbLf & ListText(hiddenSheets)
    End If
    If unreachable.Count > 0 Then
        s = s & vbLf & "移動できない場所の外部リンク: " & unreachable.Count & " 件" & vbLf & ListText(unreachable)
    End If
    If nameIssues.Count > 0 Then
        s = s & vbLf & "問題のある名前定義: " & nameIssues.Count & " 件" & vbLf & ListText(nameIssues)
    End If
    If links.Count > 0 Then
        s = s & vbLf & "外部リンク: " & links.Count & " 件" & vbLf & ListText(links)
    End If
    BuildSummary = s
End Function

Private Function ListText(ByVal items As Collection) As String
    Dim i As Long
    Dim s As String
    For i = 1 To items.Count
        If i > MAX_LIST Then
            s = s & "    ... ほか " & (items.Count - MAX_LIST) & " 件" & vbLf
            Exit For
        End If
        s = s & "    " & items(i) & vbLf
    Next i
    ListText = s
End Function
