Attribute VB_Name = "Test_mDitto"
Option Explicit
'===============================================================================
' ditto (すぐ上と同じ値のセルの文字色をグレーにする) のテスト。
'===============================================================================

Public Sub Run()
    Dim ws As Object
    Dim gray As Long, red As Long
    Dim grayed As Long, restored As Long

    Call mColorUtil.TryParseHexColor(mPalette.LookupHex("gray", "400"), gray)
    red = RGB(255, 0, 0)
    Set ws = ActiveWorkbook.Worksheets.Add

    ' A 列: 部署 / B 列: 課 (列ごとに独立して比べる)
    ws.Range("A1").Value = "営業"
    ws.Range("A2").Value = "営業"
    ws.Range("A3").Value = "営業"
    ws.Range("A4").Value = "開発"
    ws.Range("A5").Value = "開発"
    ws.Range("B1").Value = "1課"
    ws.Range("B2").Value = "1課"
    ws.Range("B3").Value = "2課"
    ws.Range("B4").Value = "2課"
    ws.Range("B5").Value = "2課"

    mDitto.ApplyDitto ws.Range("A1:B5"), grayed, restored
    TestLib.IsFalse "1行目はグレーにしない", IsGray(ws.Range("A1"), gray)
    TestLib.IsTrue "上と同じならグレー", IsGray(ws.Range("A2"), gray)
    TestLib.IsTrue "2つ上から続いてもグレー", IsGray(ws.Range("A3"), gray)
    TestLib.IsFalse "上と違えばグレーにしない", IsGray(ws.Range("A4"), gray)
    TestLib.IsTrue "列ごとに独立して比べる (左の列が変わってもグレー)", IsGray(ws.Range("B4"), gray)
    TestLib.Eq "グレーにした数", grayed, 6

    ' ---- 表示どおりの文字列で比べる ----
    ws.Range("D1").Value = 1.5
    ws.Range("D1").NumberFormat = "0"
    ws.Range("D2").Value = 2.4
    ws.Range("D2").NumberFormat = "0"
    mDitto.ApplyDitto ws.Range("D1:D2")
    TestLib.IsTrue "表示が同じならグレー (値は違っても)", IsGray(ws.Range("D2"), gray)

    ' ---- 空のセル ----
    ws.Range("E1").Value = "x"
    ws.Range("E3").Value = "x"
    mDitto.ApplyDitto ws.Range("E1:E3")
    TestLib.IsFalse "空のセルはグレーにしない", IsGray(ws.Range("E2"), gray)
    TestLib.IsFalse "空のセルのすぐ下は比べない", IsGray(ws.Range("E3"), gray)

    ' ---- 非表示の行は飛ばす ----
    ws.Range("F1").Value = "y"
    ws.Range("F2").Value = "z"
    ws.Range("F3").Value = "y"
    ws.Rows(2).Hidden = True
    mDitto.ApplyDitto ws.Range("F1:F3")
    TestLib.IsTrue "非表示の行を飛ばして、上の表示されている行と比べる", IsGray(ws.Range("F3"), gray)
    ws.Rows(2).Hidden = False

    ' ---- 結合セル ----
    ws.Range("G1").Value = "m"
    ws.Range("G1:H1").Merge
    ws.Range("G2").Value = "m"
    mDitto.ApplyDitto ws.Range("G1:G2")
    TestLib.IsFalse "結合セルのすぐ下は比べない", IsGray(ws.Range("G2"), gray)

    ' 結合セルをはさんで上下が同じ値でも、結合セルの下は比べない
    ws.Range("K1").Value = "m"
    ws.Range("K2").Value = "m"
    ws.Range("K2:L2").Merge
    ws.Range("K3").Value = "m"
    mDitto.ApplyDitto ws.Range("K1:K3")
    TestLib.IsFalse "結合セルをはさんだ下は、その上と同じ値でも比べない", IsGray(ws.Range("K3"), gray)

    ' 前の列の最後の値と、次の列の先頭が同じでも、次の列の1行目は比べない
    ws.Range("N1").Value = "p"
    ws.Range("N2").Value = "q"
    ws.Range("O1").Value = "q"
    mDitto.ApplyDitto ws.Range("N1:O2")
    TestLib.IsFalse "列が変わったら1行目は比べない (前の列の最後と同じ値でも)", IsGray(ws.Range("O1"), gray)

    ' ---- 再実行: 上と違う値になったセルを自動に戻す。ほかの色には触らない ----
    ws.Range("A2").Value = "総務"
    ws.Range("B1").Font.Color = red
    mDitto.ApplyDitto ws.Range("A1:B5"), grayed, restored
    TestLib.IsFalse "再実行: 上と違う値になったらグレーを外す", IsGray(ws.Range("A2"), gray)
    ' (「自動」に戻ったかは確かめない。LibreOffice は ColorIndex = xlColorIndexAutomatic を黒にする)
    TestLib.IsFalse "再実行: 次のセルも上と違うのでグレーを外す", IsGray(ws.Range("A3"), gray)
    TestLib.Eq "再実行: 自動に戻した数", restored, 2
    TestLib.Eq "再実行: ほかの色を付けたセルには触らない", CLng(ws.Range("B1").Font.Color), red

    ' ---- コマンド全体 ----
    ws.Range("J1").Value = "k"
    ws.Range("J2").Value = "k"
    ws.Range("J1:J2").Select
    mDitto.ditto
    TestLib.IsTrue "ditto: 選択範囲に色を付ける", IsGray(ws.Range("J2"), gray)
End Sub

Private Function IsGray(ByVal cell As Object, ByVal gray As Long) As Boolean
    IsGray = (CLng(cell.Font.Color) = gray)
End Function
