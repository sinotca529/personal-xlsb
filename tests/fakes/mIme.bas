Attribute VB_Name = "mIme"
Option Explicit
Option Private Module
'===============================================================================
' テスト用の偽物の mIme (本物は src/mIme.bas)。
'
' 本物は Win32 API (imm32) で IME を切るため Linux では動かない。テストでは
' 何もしない。公開する手続きの名前は本物と揃える (tests/run.py が照合する)。
'===============================================================================

Public Sub ImeOffLater()
End Sub

Public Sub ImeOff()
End Sub
