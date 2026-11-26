Attribute VB_Name = "mLauncherFormUtil"
Option Explicit
Option Private Module
'===============================================================================
' テスト用の偽物の mLauncherFormUtil (本物は src/mLauncherFormUtil.bas)。
'
' 本物は Win32 API (user32) でフォームのウィンドウを透過させるため Linux では
' 動かない。テストでは何もしない。公開する手続きの名前は本物と揃える
' (tests/run.py が照合する)。
'===============================================================================

Public Sub MakeFormTransparent(frm As Object)
End Sub
