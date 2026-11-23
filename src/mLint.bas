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
'         - 参照先に #REF! や外部ブックを含む名前定義
'         - Excel が把握している外部リンク (グラフ・入力規則などに残ったものを含む)
'
'     非表示シートの問題、名前定義、外部リンクは移動できないため、件数と一覧を
'     最初にダイアログでまとめて表示する (該当がなければ表示しない)。
'
' NOTES
'     - テーブルの構造化参照 (Table1[列]) は外部参照とみなさない。
'     - 図形・グラフ・条件付き書式・入力規則の中身は個別には調べない
'       (外部リンクとしては一覧に出る)。
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
    Dim extRe As Object
    Dim summary As String

    Set wb = ActiveWorkbook
    Set extRe = ExternalRefRegExp()

    For Each ws In wb.Worksheets
        Set sheetHits = ScanSheet(ws, extRe)
        If ws.Visible = xlSheetVisible Then
            For Each hit In sheetHits
                hits.Add hit
            Next hit
        ElseIf sheetHits.Count > 0 Then
            hiddenCount = hiddenCount + sheetHits.Count
            hiddenSheets.Add ws.Name & " (" & sheetHits.Count & ")"
        End If
    Next ws

    Set nameIssues = ScanNames(wb, extRe)
    Set links = ExternalLinks(wb)
    Set booster = Nothing ' 画面更新を戻してから表示・移動する

    summary = BuildSummary(hits.Count, hiddenCount, hiddenSheets, nameIssues, links)
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
' セル
'========================

' シートの問題セルを行→列の順に集める。値と数式は配列で一括に読む
Private Function ScanSheet(ByVal ws As Worksheet, ByVal extRe As Object) As Collection
    Dim result As New Collection
    Dim used As Range
    Dim vals As Variant, fmls As Variant
    Dim r As Long, c As Long
    Dim note As String
    Dim fml As String

    Set ScanSheet = result
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

' 移動できない問題 (非表示シート・名前・外部リンク) のまとめ。該当がなければ空文字列
Private Function BuildSummary(ByVal cellCount As Long, ByVal hiddenCount As Long, ByVal hiddenSheets As Collection, _
                              ByVal nameIssues As Collection, ByVal links As Collection) As String
    Dim s As String

    If hiddenCount = 0 And nameIssues.Count = 0 And links.Count = 0 Then Exit Function

    s = "表示中のシートの問題セル: " & cellCount & " 件"
    If cellCount > 0 Then s = s & " (OK のあと n / N で移動)"
    s = s & vbLf
    If hiddenCount > 0 Then
        s = s & vbLf & "非表示シートの問題セル: " & hiddenCount & " 件" & vbLf & ListText(hiddenSheets)
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
