Attribute VB_Name = "mClipboard"
Option Explicit
Option Private Module
'===============================================================================
' テスト用の偽物の mClipboard (本物は src/mClipboard.bas)。
'
' 本物は Win32 API でクリップボードに書き込むため Linux では動かない。テストでは
' 本物の代わりにこれを読み込み、書き込もうとした内容を覚えておく。テストからは
' Reset で消し、LastText / LastHtml / CallCount で調べる。
' 公開する関数の名前は本物と揃える (tests/run.py が照合する)。
'===============================================================================

Private mLastText As String
Private mLastHtml As String
Private mCalls As Long

' ---- 本物と同じ公開関数 ----

Public Function SetClipboardText(ByVal text As String) As Boolean
    mLastText = text
    mLastHtml = ""
    mCalls = mCalls + 1
    SetClipboardText = True
End Function

Public Function SetClipboardHtml(ByVal fragment As String, ByVal plainText As String) As Boolean
    mLastText = plainText
    mLastHtml = fragment
    mCalls = mCalls + 1
    SetClipboardHtml = True
End Function

' ---- テスト用 ----

Public Sub Reset()
    mLastText = ""
    mLastHtml = ""
    mCalls = 0
End Sub

' 最後に書き込んだテキスト形式の内容
Public Function LastText() As String
    LastText = mLastText
End Function

' 最後に書き込んだ HTML 形式の内容 (SetClipboardText なら空文字列)
Public Function LastHtml() As String
    LastHtml = mLastHtml
End Function

Public Function CallCount() As Long
    CallCount = mCalls
End Function
