Attribute VB_Name = "mLauncher"
Option Explicit
Option Private Module
'===============================================================================
' NAME
'     launcher - 「マクロ名 引数...」を入力してマクロを実行するランチャー
'
' DESCRIPTION
'     Ctrl+L でランチャー (Launcher フォーム) を開く。入力欄に
'     「マクロ名 引数...」を入力して Enter で実行、Esc で閉じる。
'     入力中は補完候補を表示し、Tab / Shift+Tab / 上下キーで選択する。
'
'     補完候補は build.py が各モジュールのヘッダーの COMPLETIONS 節から
'     生成し、mLauncherCompData モジュールとして埋め込む。
'
'     マクロは、個人用マクロブック (このブック) → アクティブなブックの順に探す。
'     引数は空白区切り。"..." で囲むと空白を含められ、\n \t \" \\ が使える。
'
' NOTES
'     - PERSONAL.XLSB を開いたとき (Workbook_Open) に Ctrl+L を割り当てる。
'     - Option Private Module のため、このモジュールのマクロは
'       マクロ一覧 (Alt+F8) には表示されない。
'===============================================================================

'========================
' ショートカット
'========================
Public Sub ActivateLauncher()
    Application.OnKey "^l", "'" & ThisWorkbook.Name & "'!ShowLauncher"
End Sub

Public Sub DeactivateLauncher()
    Application.OnKey "^l"
End Sub

Public Sub ShowLauncher()
    Load Launcher
    Launcher.Show
End Sub

'========================
' InputBox 版ランチャー
'========================
Public Sub PromptAndRun()
    Dim cmd As String
    cmd = Application.InputBox("", "実行するコマンドを入力")
    TryRun cmd
End Sub

'========================
' 実行
'========================
Public Sub TryRun(cmd As String)
    Dim src As Variant
    Dim a As Variant
    Dim m As String
    
    ' 1. 入力をスペースで分割して、マクロ名/引数のリストを得る
    a = SplitToken(cmd)
    If UBound(a) < 0 Then Exit Sub
    
    ' 2. 個人用マクロブック (このブック) にそのマクロがあれば、実行して終了
    ' 3. 現在アクティブなブックにそのマクロがあれば、実行して終了
    On Error Resume Next
    For Each src In Array(ThisWorkbook.Name, ActiveWorkbook.Name)
        m = "'" & src & "'!" & a(0)
    
        Select Case UBound(a)
            Case 0: Run m
            Case 1: Run m, a(1)
            Case 2: Run m, a(1), a(2)
            Case 3: Run m, a(1), a(2), a(3)
            Case 4: Run m, a(1), a(2), a(3), a(4)
            Case 5: Run m, a(1), a(2), a(3), a(4), a(5)
            Case 6: Run m, a(1), a(2), a(3), a(4), a(5), a(6)
            Case 7: Run m, a(1), a(2), a(3), a(4), a(5), a(6), a(7)
            Case 8: Run m, a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8)
            Case 9: Run m, a(1), a(2), a(3), a(4), a(5), a(6), a(7), a(8), a(9)
            Case Else: MsgBox "Too many arguments (max 9).": Exit Sub
        End Select
        
        ' 正常終了
        If Err.Number = 0 Then Exit Sub
        
        ' 実行時エラー
        ' マクロが見つからないときのエラー (1004) を無視する
        If Err.Number <> 1004 Then
            MsgBox Err.Description
            Exit Sub
        Else
            Debug.Print Err.Description
            Err.Clear
        End If
    Next src
    
    
    ' 4. マクロの実行に失敗したら、エラーダイアログを出す
    MsgBox "マクロの実行に失敗しました"
End Sub

Private Function SplitToken(text As String) As Variant
    Dim tokens() As String
    Dim token As String
    Dim cursor As Long
    Dim ch As String
    Dim nextCh As String
    Dim inQuote As Boolean
    Dim n As Long

    text = Trim$(text)
    If Len(text) = 0 Then
        SplitToken = Array()
        Exit Function
    End If

    ReDim tokens(0 To Len(text))

    For cursor = 1 To Len(text)
        ch = Mid$(text, cursor, 1)

        If ch = "\" Then
            If cursor < Len(text) Then
                nextCh = Mid$(text, cursor + 1, 1)
                Select Case nextCh
                    Case "n"
                        token = token & vbCrLf
                    Case "t"
                        token = token & vbTab
                    Case """"
                        token = token & """"
                    Case "\"
                        token = token & "\"
                    Case Else
                        token = token & "\" & nextCh
                End Select
                cursor = cursor + 1
            Else
                token = token & "\"
            End If

        ElseIf ch = """" Then
            inQuote = Not inQuote

        ElseIf ch = " " Then
            If inQuote Then
                token = token & ch
            ElseIf Len(token) > 0 Then
                tokens(n) = token
                n = n + 1
                token = ""
            End If

        Else
            token = token & ch
        End If
    Next cursor

    If Len(token) > 0 Then
        tokens(n) = token
        n = n + 1
    End If

    ReDim Preserve tokens(0 To n - 1)
    SplitToken = tokens
End Function
