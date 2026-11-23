Attribute VB_Name = "mLlp"
Option Explicit
'===============================================================================
' NAME
'     llp - Lightweight Linked Picture (軽量リンク画像) を管理する
'
' SYNOPSIS
'     llp { insert | i | update | u | migrate | m | clean | c | goto | g | localize | l }
'
' DESCRIPTION
'     リンク画像 (リンクされた図) の参照範囲を「名前定義 + IF(LLP_LINK, 参照式, "")」
'     で管理し、参照範囲の追跡・再リンクを容易にする。名前 LLP_LINK は普段 FALSE で、
'     描画を更新するときだけ一時的に TRUE にする。
'
'     解説: https://zenn.dev/sinotca/articles/18638886a004d4
'
' COMMANDS
'     insert, i      リンク画像を挿入する (事前に参照範囲をコピーしておく)。
'     update, u      すべての LLP 画像の描画を更新する。
'     migrate, m     既存のリンク画像をすべて LLP 管理下に置き換える。
'     clean, c       使われなくなった LLP 関連の名前定義を削除する。
'     goto, g        選択中の LLP 画像から参照範囲へジャンプする。
'     localize, l    LLP 参照式から外部ブック名の修飾を削除する。
'
' COMPLETIONS
'     llp             軽量リンク画像の管理
'     llp insert      リンク画像を挿入
'     llp update      すべての描画を更新
'     llp migrate     既存のリンク画像を LLP 化
'     llp clean       不要な名前定義を削除
'     llp goto        参照範囲へジャンプ
'     llp localize    外部ブック名の修飾を削除
'===============================================================================

Private Const MSO_PICTURE As Long = 13

Public Sub llp(Optional ByVal cmd As String = "")
    Dim errNum As Long, errDesc As String

    On Error GoTo FINALLY
    Randomize
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    InitLlpUpdateFlag
    Select Case LCase$(cmd)
        Case "insert", "i": InsertLlp
        Case "update", "u": UpdateAll
        Case "migrate", "m": MigrateAll
        Case "clean", "c": CleanOrphanNames
        Case "goto", "g": GotoSelected
        Case "localize", "l": LocalizeBookRefs
    End Select

FINALLY:
    ' Err はここで確定させる。後片付け自体が失敗しても
    ' Application の状態復元が止まらないよう On Error Resume Next で継続する。
    errNum = Err.Number: errDesc = Err.Description
    On Error Resume Next
    SetLlpUpdateFlag False
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error GoTo 0
    If errNum <> 0 Then
        Application.StatusBar = "E" & errNum & ": " & errDesc
    End If
End Sub

Private Sub InitLlpUpdateFlag()
    On Error Resume Next
    ActiveWorkbook.Names.add name:="LLP_LINK", RefersTo:="=FALSE"
End Sub

Private Sub SetLlpUpdateFlag(ByVal state As Boolean)
    ActiveWorkbook.Names("LLP_LINK").RefersTo = "=" & state
End Sub

Function GenId(Optional ByVal length As Long = 21) As String
    Const chars As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Dim i As Long

    GenId = ""
    For i = 1 To length
        GenId = GenId & Mid$(chars, Int(Rnd() * Len(chars)) + 1, 1)
    Next i
End Function

' LLP_LINK フラグが TRUE のときに f と評価される名前を作り、返す。
Private Function NewLlpName(ByVal f As String) As String
    NewLlpName = "LLP_" & GenId()
    ActiveWorkbook.Names.add name:=NewLlpName, RefersTo:="=IF(LLP_LINK," & f & ","""")"
End Function

Private Sub InsertLlp()
    Dim s As shape
    Dim f As String

    If Application.CutCopyMode <> xlCopy Then
        Err.Raise 601, "llp insert", "参照範囲をコピーした状態で実行してください。"
    End If

    ' リンクされた図を貼り付ける
    Application.CommandBars.ExecuteMso "PastePictureLink"
    Set s = ActiveSheet.Shapes(ActiveSheet.Shapes.Count)

    ' 貼り付けた図から、参照先を表す式を取得する
    f = GetFormulaOrEmpty(s)
    If Len(f) = 0 Then Err.Raise 602, "llp insert", "参照範囲を取得できませんでした。"

    ' 範囲として評価できない式はセットできないので、一度フラグを TRUE にする
    SetLlpUpdateFlag True
    s.DrawingObject.formula = NewLlpName(f)
    SetLlpUpdateFlag False
End Sub

Private Sub UpdateAll()
    SetLlpUpdateFlag True
    Application.Wait Now + TimeSerial(0, 0, 1)
    SetLlpUpdateFlag False
End Sub

Private Sub MigrateAll()
    Dim ws As Worksheet
    Dim s As shape
    Dim f As String

    SetLlpUpdateFlag True
    For Each ws In ActiveWorkbook.Worksheets
        For Each s In ws.Shapes
            f = GetFormulaOrEmpty(s)
            If Len(f) <> 0 And Not (f Like "LLP_*") Then
                s.DrawingObject.formula = "=" & NewLlpName(f)
            End If
        Next s
    Next ws
    SetLlpUpdateFlag False
    Application.StatusBar = "移行が完了しました。"
End Sub

Private Sub CleanOrphanNames()
    Dim used As Object
    Dim ws As Worksheet
    Dim s As shape
    Dim nm As name
    Dim llpName As String
    Set used = CreateObject("Scripting.Dictionary")

    ' 使用中の LLP_* を収集
    used.add "LLP_LINK", True
    For Each ws In ActiveWorkbook.Worksheets
        For Each s In ws.Shapes
            llpName = GetLlpNameOrEmpty(s)
            If Not used.Exists(llpName) Then used.add llpName, True
        Next s
    Next ws

    ' 未使用の LLP_* を削除 (LLP_LINK は除外)
    For Each nm In ActiveWorkbook.Names
        If (nm.name Like "LLP_*") And (Not used.Exists(nm.name)) Then nm.Delete
    Next nm

    Application.StatusBar = "孤児リンクを削除しました。"
End Sub

Private Sub GotoSelected()
    Dim src As String

    ' 選択中の画像の参照範囲を取得
    On Error GoTo HANDLE_ERROR
    If (Selection.ShapeRange.Count <> 1) Then GoTo HANDLE_ERROR
    src = GetLlpSrcRangeOrEmptyS(Selection.ShapeRange(1))
    If Len(src) = 0 Then GoTo HANDLE_ERROR

    ' 選択範囲に移動
    Application.Goto Application.Range(src), True
    Exit Sub

HANDLE_ERROR:
    On Error GoTo 0
    Err.Raise 603, "llp goto ", "LLP 画像を1つだけ選択してください。"
End Sub

Private Sub LocalizeBookRefs()
    Dim nm As name
    Dim srcRange As String
    Dim srcRangeWoBk As String

    For Each nm In ActiveWorkbook.Names
        If nm.name = "LLP_LINK" Then GoTo NEXT_NAME
        If Not (nm.name Like "LLP_*") Then GoTo NEXT_NAME

        srcRange = GetLlpSrcRangeOrEmptyF(nm.RefersTo)
        If Len(srcRange) = 0 Then GoTo NEXT_NAME

        srcRangeWoBk = RemoveBookQualifier(srcRange)
        If srcRangeWoBk = srcRange Then GoTo NEXT_NAME

        nm.RefersTo = "=IF(LLP_LINK," & srcRangeWoBk & ","""")"
NEXT_NAME:
    Next nm

    Application.StatusBar = "LLP 画像から外部ブック名を削除しました。"
End Sub

' 画像から参照先の式を取得
Private Function GetFormulaOrEmpty(ByRef s As shape) As String
    On Error Resume Next
    ' リンク付き画像は msoPicuture なので、 msoLinkedPicture は見なくて良い。
    If s.Type <> MSO_PICTURE Then Exit Function
    GetFormulaOrEmpty = Trim$(s.DrawingObject.formula)
    On Error GoTo 0
End Function

' LLP 管理下の Shape について、参照している名前を返す
Private Function GetLlpNameOrEmpty(ByRef s As shape) As String
    Dim f As String
    f = GetFormulaOrEmpty(s)
    If f Like "LLP_*" Then GetLlpNameOrEmpty = f
    On Error GoTo 0
End Function


' LLP 管理下の Shape について、参照先の範囲を返す
Private Function GetLlpSrcRangeOrEmptyS(ByRef s As shape) As String
    Dim nm As String
    Dim f As String
    nm = GetLlpNameOrEmpty(s)
    If Len(nm) = 0 Then Exit Function

    f = ActiveWorkbook.Names(Trim$(nm)).RefersTo
    GetLlpSrcRangeOrEmptyS = GetLlpSrcRangeOrEmptyF(f)
End Function

' LLP 式から参照先の範囲を取り出す
Private Function GetLlpSrcRangeOrEmptyF(ByRef f As String) As String
    Const prefix As String = "=IF(LLP_LINK,"
    Const SUFFIX As String = ","""")"
    If Not f Like (prefix & "*") Then Exit Function
    If Not f Like ("*" & SUFFIX) Then Exit Function
    GetLlpSrcRangeOrEmptyF = Mid$(f, Len(prefix) + 1, Len(f) - Len(prefix) - Len(SUFFIX))
End Function

' 範囲から 外部ブック修飾を除去
'  例:  "'[C:\path\[Book.xlsx]Sheet A'!$A$1" → "'Sheet A'!$A$1"
'       "[Book.xlsx]Sheet1!$A$1:$B$2"        → "Sheet1!$A$1:$B$2"
Function RemoveBookQualifier(ByRef RefText As String) As String
    Dim exclPos As Long
    Dim bracketPos As Long

    RemoveBookQualifier = RefText

    ' 末尾から ! を探す。
    exclPos = InStrRev(RefText, "!")
    If exclPos = 0 Then Exit Function

    ' ] を探す
    bracketPos = InStrRev(Left$(RefText, exclPos - 1), "]")
    If bracketPos = 0 Then Exit Function


    RemoveBookQualifier = Mid$(RefText, bracketPos + 1)

    ' ! の直前がシングルクォートなら、 ' 頭に ' を付与
    If Mid$(RefText, exclPos - 1, 1) = "'" Then
        RemoveBookQualifier = "'" & RemoveBookQualifier
    End If
End Function

