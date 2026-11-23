Attribute VB_Name = "mClipboard"
Option Explicit
Option Private Module
' クリップボードへ Unicode 文字列を書き込む。
' MSForms.DataObject はエクスプローラーのウィンドウが開いていると文字化け (??) する
' 既知の不具合があるため、Win32 API で CF_UNICODETEXT を直接書き込む。

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWndNewOwner As LongPtr) As Long
Private Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal uFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal dest As LongPtr, ByVal src As LongPtr, ByVal length As LongPtr)

Private Const CF_UNICODETEXT As Long = 13
Private Const GHND As Long = &H42 ' GMEM_MOVEABLE Or GMEM_ZEROINIT

' text をクリップボードに書き込む。成功すれば True
Public Function SetClipboardText(ByVal text As String) As Boolean
    Dim hMem As LongPtr
    Dim ptr As LongPtr
    Dim bytes As LongPtr

    bytes = (Len(text) + 1) * 2 ' 終端の NUL を含む UTF-16 (ゼロ初期化で NUL になる)
    hMem = GlobalAlloc(GHND, bytes)
    If hMem = 0 Then Exit Function

    ptr = GlobalLock(hMem)
    If ptr = 0 Then GlobalFree hMem: Exit Function
    If Len(text) > 0 Then CopyMemory ptr, StrPtr(text), bytes - 2
    GlobalUnlock hMem

    If OpenClipboard(0) = 0 Then GlobalFree hMem: Exit Function
    EmptyClipboard
    If SetClipboardData(CF_UNICODETEXT, hMem) = 0 Then
        GlobalFree hMem ' 所有権が移らなかったので自分で解放する
    Else
        SetClipboardText = True
    End If
    CloseClipboard
End Function
