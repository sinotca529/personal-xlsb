Attribute VB_Name = "Test_mHtml"
Option Explicit
'===============================================================================
' html の表組み立て (mHtml.TableHtml) のテスト。
'
' クリップボードへの書き込み (Win32 API) は Linux では動かないので、TableHtml が
' 返す HTML のソースを調べる。LibreOffice では次の値を読み戻せないため検証しない
' (docs/TESTING.md の「再現度の限界」)。
'   - 罫線の種類 (LineStyle は常に xlContinuous)
'   - 横位置 (HorizontalAlignment は Nothing)
'   - セル内の一部だけの書式 (書式が混在してもセルの Font.Bold などが Null にならない)
'===============================================================================

Public Sub Run()
    Dim ws As Object
    Dim h As String, td As String
    Dim nR As Long, nC As Long

    Set ws = ActiveWorkbook.Worksheets.Add

    ' ---- 基本: 行・列の数、エスケープ、改行 ----
    ws.Range("A1").Value = "a<b&c"
    ws.Range("B1").Value = 12
    ws.Range("A2").Value = "x" & vbLf & "y"
    ws.Range("B2").Value = "plain"
    h = mHtml.TableHtml(ws.Range("A1:B2"), nR, nC)
    TestLib.IsTrue "table で始まる", Left$(h, 6) = "<table"
    TestLib.Eq "行数を返す", nR, 2
    TestLib.Eq "列数を返す", nC, 2
    TestLib.Eq "<tr> は 2 つ", CountOf(h, "<tr"), 2
    TestLib.Eq "<td> は 4 つ", CountOf(h, "<td"), 4
    TestLib.Eq "<col> は 2 つ", CountOf(h, "<col "), 2
    TestLib.IsTrue "< と & をエスケープする", InStr(h, ">a&lt;b&amp;c</td>") > 0
    TestLib.IsTrue "セル内の改行を <br> にする", InStr(h, ">x<br>y</td>") > 0
    TestLib.IsTrue "標準の数値は右寄せ", InStr(TdOf(h, "12"), "text-align:right;") > 0
    TestLib.IsTrue "標準の文字列は text-align を付けない", InStr(TdOf(h, "plain"), "text-align") = 0
    TestLib.IsTrue "縦位置は既定で下", InStr(TdOf(h, "plain"), "vertical-align:bottom;") > 0
    TestLib.IsTrue "折り返さないセルは nowrap", InStr(TdOf(h, "plain"), "white-space:nowrap;") > 0

    ' ---- 塗り・文字の書式 ----
    ws.Range("D1").Value = "fill"
    ws.Range("D1").Interior.Color = RGB(255, 0, 0)
    ws.Range("E1").Value = "blue"
    ws.Range("E1").Font.Color = RGB(0, 0, 255)
    ws.Range("F1").Value = "bold"
    ws.Range("F1").Font.Bold = True
    ws.Range("G1").Value = "deco"
    ws.Range("G1").Font.Italic = True
    ws.Range("G1").Font.Underline = xlUnderlineStyleSingle
    ws.Range("G1").Font.Strikethrough = True
    ws.Range("H1").Value = "wrap"
    ws.Range("H1").WrapText = True
    ws.Range("H1").VerticalAlignment = xlTop
    h = mHtml.TableHtml(ws.Range("D1:H1"))
    TestLib.IsTrue "塗り色を background にする", InStr(TdOf(h, "fill"), "background:#FF0000;") > 0
    TestLib.IsTrue "塗りのないセルに background を付けない", InStr(TdOf(h, "blue"), "background") = 0
    TestLib.IsTrue "文字色を color にする", InStr(TdOf(h, "blue"), "color:#0000FF;") > 0
    TestLib.IsTrue "自動の文字色は出さない", InStr(TdOf(h, "fill"), "color:") = 0
    TestLib.IsTrue "太字", InStr(TdOf(h, "bold"), "font-weight:bold;") > 0
    TestLib.IsTrue "斜体", InStr(TdOf(h, "deco"), "font-style:italic;") > 0
    TestLib.IsTrue "下線と取り消し線", InStr(TdOf(h, "deco"), "text-decoration:underline line-through;") > 0
    TestLib.IsTrue "折り返すセルは nowrap にしない", InStr(TdOf(h, "wrap"), "nowrap") = 0
    TestLib.IsTrue "縦位置 上", InStr(TdOf(h, "wrap"), "vertical-align:top;") > 0

    ' ---- 罫線 (種類は読めないので太さと色) ----
    ws.Range("J1").Value = "border"
    With ws.Range("J1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .Color = RGB(0, 128, 0)
    End With
    h = mHtml.TableHtml(ws.Range("J1"))
    TestLib.IsTrue "太い罫線を 3px にする", InStr(TdOf(h, "border"), "border-bottom:3px solid #008000;") > 0

    ' ---- 結合セル ----
    ws.Range("A4").Value = "wide"
    ws.Range("A4:B4").Merge
    ws.Range("C4").Value = "tall"
    ws.Range("C4:C5").Merge
    ws.Range("A5").Value = "p"
    ws.Range("B5").Value = "q"
    h = mHtml.TableHtml(ws.Range("A4:C5"))
    TestLib.IsTrue "横の結合は colspan", InStr(TdOf(h, "wide"), " colspan=""2""") > 0
    TestLib.IsTrue "縦の結合は rowspan", InStr(TdOf(h, "tall"), " rowspan=""2""") > 0
    TestLib.Eq "結合で隠れたセルは出さない", CountOf(h, "<td"), 4

    ' 選択範囲からはみ出した結合セル: 範囲内だけを結合扱いにし、値は左上から取る
    ws.Range("E4").Value = "outside"
    ws.Range("E4:F4").Merge
    h = mHtml.TableHtml(ws.Range("F4:G4"), nR, nC)
    TestLib.IsTrue "はみ出した結合セルの値を出す", InStr(h, ">outside</td>") > 0
    TestLib.Eq "はみ出した結合セルに colspan を付けない", CountOf(h, "colspan"), 0
    TestLib.Eq "はみ出した結合セルの範囲の列数", nC, 2

    ' ---- 非表示の行・列 ----
    ws.Range("A8").Value = "r1"
    ws.Range("A9").Value = "hiddenrow"
    ws.Range("A10").Value = "r3"
    ws.Range("B8").Value = "hiddencol"
    ws.Rows(9).Hidden = True
    ws.Columns(2).Hidden = True
    h = mHtml.TableHtml(ws.Range("A8:C10"), nR, nC)
    TestLib.Eq "非表示の行を数えない", nR, 2
    TestLib.Eq "非表示の列を数えない", nC, 2
    TestLib.IsTrue "非表示の行の値を出さない", InStr(h, "hiddenrow") = 0
    TestLib.IsTrue "非表示の列の値を出さない", InStr(h, "hiddencol") = 0
    TestLib.Eq "すべて非表示なら空文字列", mHtml.TableHtml(ws.Range("A9:C9")), ""
    ws.Rows(9).Hidden = False
    ws.Columns(2).Hidden = False

    ' ---- 列幅・行の高さ ----
    ws.Range("A12").Value = "size"
    ws.Rows(12).RowHeight = 30
    h = mHtml.TableHtml(ws.Range("A12"))
    TestLib.IsTrue "行の高さを pt で書く", InStr(h, "<tr style=""height:30pt"">") > 0
    TestLib.IsTrue "列幅を pt で書く", InStr(h, "<col style=""width:" & Trim$(Str$(Round(ws.Columns(1).Width, 2))) & "pt"">") > 0
End Sub

Private Function CountOf(ByVal s As String, ByVal sub_ As String) As Long
    CountOf = (Len(s) - Len(Replace(s, sub_, ""))) \ Len(sub_)
End Function

' 中身が text の <td ...> の開きタグ (見つからなければ空文字列)
Private Function TdOf(ByVal h As String, ByVal text As String) As String
    Dim p As Long, q As Long
    p = InStr(h, ">" & text & "</td>")
    If p = 0 Then Exit Function
    q = InStrRev(h, "<td", p)
    If q = 0 Then Exit Function
    TdOf = Mid$(h, q, p - q + 1)
End Function
