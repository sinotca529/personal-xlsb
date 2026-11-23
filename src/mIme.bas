Attribute VB_Name = "mIme"
Option Explicit
Option Private Module
' IME (日本語入力) の状態を操作する。

Private Declare PtrSafe Function GetFocus Lib "user32" () As LongPtr
Private Declare PtrSafe Function ImmGetContext Lib "imm32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ImmSetOpenStatus Lib "imm32" (ByVal hIMC As LongPtr, ByVal fOpen As Long) As Long
Private Declare PtrSafe Function ImmReleaseContext Lib "imm32" (ByVal hWnd As LongPtr, ByVal hIMC As LongPtr) As Long

' 現在のマクロ (とランチャーのフォーム) が終わった直後に IME をオフにするよう予約する。
' マクロ実行中はフォームなどにフォーカスがあるため、その場でオフにしても Excel 本体には効かない。
Public Sub ImeOffLater()
    Application.OnTime Now, "'" & ThisWorkbook.Name & "'!mIme.ImeOff"
End Sub

' フォーカスのあるウィンドウの IME をオフにする。Microsoft IME では入力モード「半角英数」に当たる
Public Sub ImeOff()
    Dim hWnd As LongPtr
    Dim hIMC As LongPtr

    hWnd = GetFocus()
    If hWnd = 0 Then Exit Sub
    hIMC = ImmGetContext(hWnd)
    If hIMC = 0 Then Exit Sub
    ImmSetOpenStatus hIMC, 0
    ImmReleaseContext hWnd, hIMC
End Sub
