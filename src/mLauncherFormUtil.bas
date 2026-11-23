Attribute VB_Name = "mLauncherFormUtil"
Option Explicit
Option Private Module

Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String _
) As Long

Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIndex As Long _
) As LongPtr

Private Declare PtrSafe Function SetWindowLongPtrA Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As LongPtr _
) As LongPtr

Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hWnd As LongPtr, _
    ByVal crey As Long, _
    ByVal bAlpha As Byte, _
    ByVal dwFlags As Long _
) As Long

Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_STYLE As Long = -16
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_CAPTION As Long = &HC00000
Private Const LWA_COLORKEY As Long = &H1

Public Sub MakeFormTransparent(frm As Object)
    Dim hWnd As LongPtr
    Dim style As LongPtr

    hWnd = FindWindowA(vbNullString, frm.Caption)
    If hWnd = 0 Then Exit Sub

    ' タイトルバーを消す
    style = GetWindowLongPtrA(hWnd, GWL_STYLE)
    style = style And Not WS_CAPTION
    SetWindowLongPtrA hWnd, GWL_STYLE, style

    ' 枠線がなく、透過できるウィンドウにする
    style = GetWindowLongPtrA(hWnd, GWL_EXSTYLE)
    style = (style Or WS_EX_LAYERED) And Not WS_EX_DLGMODALFRAME
    SetWindowLongPtrA hWnd, GWL_EXSTYLE, style

    frm.BackColor = &HFF00FF
    SetLayeredWindowAttributes hWnd, frm.BackColor, 0, LWA_COLORKEY
End Sub
