Attribute VB_Name = "Test_mMarkdown"
Option Explicit
'===============================================================================
' md のテスト。表の組み立て (mMarkdown.MarkdownTable) と、コマンド全体 (mMarkdown.md)。
'
' クリップボードは偽物 (tests/fakes/mClipboard.bas) に差し替わっていて、書き込もうとした
' 内容を mClipboard.LastText で読める。
' 列幅の揃え (全角を2文字ぶんと数える StrConv / LenB) は LibreOffice では Excel と
' 結果が違うため、空白と - の数は無視して比べる (Squash)。
'===============================================================================

Public Sub Run()
    Dim ws As Object
    Dim got As String

    Set ws = ActiveWorkbook.Worksheets.Add

    ws.Range("A1").Value = "名前"
    ws.Range("B1").Value = "数"
    ws.Range("A2").Value = "a"
    ws.Range("B2").Value = 1
    ws.Range("A3").Value = "bb"
    ws.Range("B3").Value = 22
    got = mMarkdown.MarkdownTable(ws.Range("A1:B3"))
    TestLib.Eq "見出し・区切り・データ行。数値だけの列は右寄せ", Squash(got), _
               "|名前|数|" & vbCrLf & "|-|-:|" & vbCrLf & "|a|1|" & vbCrLf & "|bb|22|" & vbCrLf
    TestLib.IsTrue "データ行の数値を右に寄せる", InStr(got, " 1 |") > 0 And InStr(got, "| 1") = 0

    ws.Range("D1").Value = "h"
    ws.Range("D2").Value = "a|b" & vbLf & "c"
    TestLib.Eq "| をエスケープし、改行を <br> にする", Squash(mMarkdown.MarkdownTable(ws.Range("D1:D2"))), _
               "|h|" & vbCrLf & "|-|" & vbCrLf & "|a\|b<br>c|" & vbCrLf

    ws.Range("F1").Value = 5
    TestLib.Eq "1 行だけの表は右寄せにしない", Squash(mMarkdown.MarkdownTable(ws.Range("F1"))), _
               "|5|" & vbCrLf & "|-|" & vbCrLf

    ws.Range("H1").Value = "n"
    ws.Range("H2").Value = 1
    ws.Range("H3").Value = "x"
    TestLib.Eq "文字列が混じる列は右寄せにしない", Squash(mMarkdown.MarkdownTable(ws.Range("H1:H3"))), _
               "|n|" & vbCrLf & "|-|" & vbCrLf & "|1|" & vbCrLf & "|x|" & vbCrLf

    ' ---- コマンド全体: 選択範囲をクリップボードへ ----
    mClipboard.Reset
    ws.Range("A1:B3").Select
    mMarkdown.md
    TestLib.Eq "md: 選択範囲の表をクリップボードに書き込む", mClipboard.LastText(), _
               mMarkdown.MarkdownTable(ws.Range("A1:B3"))
    TestLib.Eq "md: 書き込みは 1 回", mClipboard.CallCount(), 1

    ' 列全体を選んでも使用範囲 (UsedRange) と重なる部分だけにする
    mClipboard.Reset
    ws.Columns(1).Select
    mMarkdown.md
    TestLib.Eq "md: 列全体の選択は使用範囲に絞る", Squash(mClipboard.LastText()), _
               "|名前|" & vbCrLf & "|-|" & vbCrLf & "|a|" & vbCrLf & "|bb|" & vbCrLf

    ' 使用範囲の外だけを選んだら書き込まない
    mClipboard.Reset
    ws.Range("Z100").Select
    mMarkdown.md
    TestLib.Eq "md: 空の選択範囲では書き込まない", mClipboard.CallCount(), 0
End Sub

' 空白を消し、続く - を 1 つにまとめる
Private Function Squash(ByVal s As String) As String
    s = Replace(s, " ", "")
    Do While InStr(s, "--") > 0
        s = Replace(s, "--", "-")
    Loop
    Squash = s
End Function
