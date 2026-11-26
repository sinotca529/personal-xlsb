Attribute VB_Name = "TestLib"
Option Explicit
'===============================================================================
' TestLib - VBA テスト用のごく小さなアサーションライブラリ
'
'   テストモジュールは Public Sub Run() を 1 つ公開し、その中から
'   TestLib.Eq / TestLib.IsTrue / TestLib.IsFalse を呼ぶ。
'   結果は 1 行 1 件のタブ区切りテキストとして蓄積し、Python 側が解釈する。
'===============================================================================

Private mLines As String
Private mModule As String
Private mPass As Long
Private mFail As Long

Public Sub TBegin()
    mLines = ""
    mModule = ""
    mPass = 0
    mFail = 0
End Sub

Public Sub TModule(ByVal moduleName As String)
    mModule = moduleName
End Sub

Private Sub Emit(ByVal kind As String, ByVal desc As String, ByVal got As String, ByVal want As String)
    mLines = mLines & kind & vbTab & mModule & vbTab & OneLine(desc) & vbTab & OneLine(got) & vbTab & OneLine(want) & vbLf
End Sub

' 1 件 1 行のタブ区切りを崩さないよう、改行とタブを \r \n \t と書く
Private Function OneLine(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    OneLine = Replace(s, vbTab, "\t")
End Function

' 期待値と実測値を文字列化して比較する (VBA の型強制の差を避けるため)。
Public Sub Eq(ByVal desc As String, ByVal actual As Variant, ByVal expected As Variant)
    Dim a As String, e As String
    a = Stringify(actual)
    e = Stringify(expected)
    If a = e Then
        mPass = mPass + 1
        Emit "PASS", desc, a, e
    Else
        mFail = mFail + 1
        Emit "FAIL", desc, a, e
    End If
End Sub

Public Sub IsTrue(ByVal desc As String, ByVal actual As Variant)
    Eq desc, CBool(actual), True
End Sub

Public Sub IsFalse(ByVal desc As String, ByVal actual As Variant)
    Eq desc, CBool(actual), False
End Sub

' 想定外の実行時エラーをテスト失敗として記録する。
Public Sub Failed(ByVal desc As String, ByVal detail As String)
    mFail = mFail + 1
    Emit "FAIL", desc, detail, "(no error)"
End Sub

Private Function Stringify(ByVal v As Variant) As String
    If IsObject(v) Then
        If v Is Nothing Then
            Stringify = "Nothing"
        Else
            Stringify = TypeName(v)
        End If
    ElseIf IsEmpty(v) Then
        Stringify = "Empty"
    ElseIf IsNull(v) Then
        Stringify = "Null"
    ElseIf VarType(v) = vbBoolean Then
        Stringify = IIf(v, "True", "False")
    Else
        Stringify = CStr(v)
    End If
End Function

Public Function TReport() As String
    TReport = mLines & "SUMMARY" & vbTab & mPass & vbTab & mFail
End Function
