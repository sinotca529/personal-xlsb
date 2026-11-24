Attribute VB_Name = "mClipboard"
Option Explicit
Option Private Module
' クリップボードへ Unicode 文字列 (と HTML) を書き込む。
' MSForms.DataObject はエクスプローラーのウィンドウが開いていると文字化け (??) する
' 既知の不具合があるため、Win32 API で CF_UNICODETEXT / CF_HTML を直接書き込む。

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
Private Declare PtrSafe Function RegisterClipboardFormatW Lib "user32" (ByVal lpszFormat As LongPtr) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, ByVal dwFlags As Long, _
    ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As LongPtr, ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long

Private Const CF_UNICODETEXT As Long = 13
Private Const GHND As Long = &H42 ' GMEM_MOVEABLE Or GMEM_ZEROINIT
Private Const CP_UTF8 As Long = 65001

' text をクリップボードに書き込む。成功すれば True
Public Function SetClipboardText(ByVal text As String) As Boolean
    Dim hText As LongPtr
    hText = AllocUnicode(text)
    If hText = 0 Then Exit Function

    If OpenClipboard(0) = 0 Then GlobalFree hText: Exit Function
    EmptyClipboard
    SetClipboardText = PutData(CF_UNICODETEXT, hText)
    CloseClipboard
End Function

' fragment (HTML の断片) を HTML 形式 (CF_HTML) で、plainText をテキスト形式で
' 同時に書き込む。貼り付け先のアプリが受け取れる方を選ぶ。両方成功すれば True
Public Function SetClipboardHtml(ByVal fragment As String, ByVal plainText As String) As Boolean
    Dim fmt As Long
    Dim hHtml As LongPtr, hText As LongPtr
    Dim okHtml As Boolean, okText As Boolean
    Dim data() As Byte

    fmt = RegisterClipboardFormatW(StrPtr("HTML Format"))
    If fmt = 0 Then Exit Function

    data = CfHtml(fragment)
    hHtml = AllocBytes(data)
    If hHtml = 0 Then Exit Function
    hText = AllocUnicode(plainText)
    If hText = 0 Then GlobalFree hHtml: Exit Function

    If OpenClipboard(0) = 0 Then GlobalFree hHtml: GlobalFree hText: Exit Function
    EmptyClipboard
    okHtml = PutData(fmt, hHtml)
    okText = PutData(CF_UNICODETEXT, hText)
    CloseClipboard
    SetClipboardHtml = okHtml And okText
End Function

' =====================

' SetClipboardData で渡す。所有権が移らなかったときは自分で解放する
Private Function PutData(ByVal fmt As Long, ByVal hMem As LongPtr) As Boolean
    If SetClipboardData(fmt, hMem) = 0 Then
        GlobalFree hMem
    Else
        PutData = True
    End If
End Function

' NUL 終端の UTF-16 文字列を入れたグローバルメモリ。失敗すれば 0
Private Function AllocUnicode(ByVal text As String) As LongPtr
    Dim hMem As LongPtr
    Dim ptr As LongPtr
    Dim bytes As LongPtr

    bytes = (Len(text) + 1) * 2 ' 終端の NUL を含む (ゼロ初期化で NUL になる)
    hMem = GlobalAlloc(GHND, bytes)
    If hMem = 0 Then Exit Function

    ptr = GlobalLock(hMem)
    If ptr = 0 Then GlobalFree hMem: Exit Function
    If Len(text) > 0 Then CopyMemory ptr, StrPtr(text), bytes - 2
    GlobalUnlock hMem
    AllocUnicode = hMem
End Function

' バイト列をそのまま入れたグローバルメモリ。失敗すれば 0
Private Function AllocBytes(ByRef data() As Byte) As LongPtr
    Dim hMem As LongPtr
    Dim ptr As LongPtr
    Dim bytes As LongPtr

    bytes = UBound(data) - LBound(data) + 1
    hMem = GlobalAlloc(GHND, bytes)
    If hMem = 0 Then Exit Function

    ptr = GlobalLock(hMem)
    If ptr = 0 Then GlobalFree hMem: Exit Function
    CopyMemory ptr, VarPtr(data(LBound(data))), bytes
    GlobalUnlock hMem
    AllocBytes = hMem
End Function

' CF_HTML の書式 (ヘッダーに UTF-8 のバイト位置を書いた HTML) にした、NUL 終端の UTF-8
Private Function CfHtml(ByVal fragment As String) As Byte()
    Dim pre As String, post As String, header As String
    Dim startHtml As Long, startFrag As Long, endFrag As Long, endHtml As Long

    pre = "<html><head><meta charset=""utf-8""></head><body>" & vbCrLf & "<!--StartFragment-->"
    post = "<!--EndFragment-->" & vbCrLf & "</body></html>"

    ' 位置は10桁固定なので、ヘッダーの長さは数値によらない (ASCII なので文字数 = バイト数)
    startHtml = Len(CfHtmlHeader(0, 0, 0, 0))
    startFrag = startHtml + Utf8Length(pre)
    endFrag = startFrag + Utf8Length(fragment)
    endHtml = endFrag + Utf8Length(post)
    header = CfHtmlHeader(startHtml, endHtml, startFrag, endFrag)

    CfHtml = ToUtf8(header & pre & fragment & post)
End Function

Private Function CfHtmlHeader(ByVal startHtml As Long, ByVal endHtml As Long, _
                              ByVal startFrag As Long, ByVal endFrag As Long) As String
    CfHtmlHeader = "Version:0.9" & vbCrLf & _
                   "StartHTML:" & Format$(startHtml, "0000000000") & vbCrLf & _
                   "EndHTML:" & Format$(endHtml, "0000000000") & vbCrLf & _
                   "StartFragment:" & Format$(startFrag, "0000000000") & vbCrLf & _
                   "EndFragment:" & Format$(endFrag, "0000000000") & vbCrLf
End Function

Private Function Utf8Length(ByVal s As String) As Long
    If Len(s) = 0 Then Exit Function
    Utf8Length = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
End Function

' UTF-8 に変換し、末尾に NUL を1バイト付ける
Private Function ToUtf8(ByVal s As String) As Byte()
    Dim n As Long
    Dim data() As Byte

    n = Utf8Length(s)
    ReDim data(0 To n) ' data(n) は NUL (Byte 配列は 0 で初期化される)
    If n > 0 Then WideCharToMultiByte CP_UTF8, 0, StrPtr(s), Len(s), VarPtr(data(0)), n, 0, 0
    ToUtf8 = data
End Function
