Attribute VB_Name = "Launcher"
Private Const MAX_ROWS As Long = 5    ' 補完の最大表示行数
Private Const ROW_HEIGHT As Long = 18 ' 補完候補の行高

Private Sub UserForm_Initialize()
    ' フォーム背景の透過
    mLauncherFormUtil.MakeFormTransparent Me

    ' 部品のプロパティを調整
    With Me.TextBox1
        .Font.Name = "MS Gothic"
        .Font.Size = 18
        .Top = 0
        .Left = 0
        .Width = WorksheetFunction.Min(360, ActiveWindow.Width * 0.7)
        .Height = 24
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = &HFFFFFF
    End With

    With Me.ListBox1
        .IntegralHeight = False
        .Font.Name = Me.TextBox1.Font.Name
        .Font.Size = Me.TextBox1.Font.Size
        .Left = Me.TextBox1.Left
        .Top = Me.TextBox1.Top + Me.TextBox1.Height
        .Width = Me.TextBox1.Width
        .Height = MAX_ROWS * ROW_HEIGHT + 1
        .BorderStyle = fmBorderStyleSingle
        .SpecialEffect = fmSpecialEffectFlat
        .BackColor = &HF0F0F0
        .Visible = False
        .Clear
    End With

    Me.Width = Me.TextBox1.Width
    Me.Height = Me.TextBox1.Height + Me.ListBox1.Height

    ' マクロ一覧の取得
     CP_LoadCompletions
     UpdateCandidates
End Sub

Private Sub TextBox1_KeyDown( _
    ByVal KeyCode As MSForms.ReturnInteger, _
    ByVal Shift As Integer _
)
    ' 候補選択
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
        SelectCandidate IIf(KeyCode = vbKeyDown, 1, -1)
        KeyCode = 0
        Exit Sub
    End If

    If KeyCode = vbKeyTab Then
        SelectCandidate IIf((Shift And 1) <> 0, -1, 1)
        KeyCode = 0
        Exit Sub
    End If

    ' 選択解除
    Me.ListBox1.ListIndex = -1

    ' 実行/キャンセル
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyEscape Then
        If KeyCode = vbKeyReturn Then mLauncher.TryRun Me.TextBox1.text
        Unload Me
        KeyCode = 0
        Exit Sub
    End If
End Sub

Private Sub TextBox1_Change()
    If Me.ListBox1.ListIndex = -1 Then UpdateCandidates
End Sub

Private Sub SelectCandidate(idxDelta As Integer)
    If Me.ListBox1.ListCount = 0 Then Exit Sub

    ' 選択対象の更新
    With Me.ListBox1
        .ListIndex = (.ListCount + .ListIndex + idxDelta) Mod .ListCount
    End With

    ' 入力欄に選択対象を反映
    With Me.TextBox1
        .text = Me.ListBox1.List(Me.ListBox1.ListIndex)
        .SelStart = Len(Me.TextBox1.text)
        .SelLength = 0
    End With
End Sub

Private Sub UpdateCandidates()
    Dim ctx As String, pre As String
    Dim cands As Collection, pair As Variant
    Dim i As Long, numShow As Long

    Me.ListBox1.Clear

    Set cands = CP_ListCandidates(Me.TextBox1.text)

    For i = 1 To cands.Count
        pair = cands(i)
        Me.ListBox1.AddItem CP_ReplaceLastToken(Me.TextBox1.text, CStr(pair(0)))
        ' 説明を2列目に出す場合:
        ' Me.ListBox1.List(Me.ListBox1.ListCount - 1, 1) = CStr(pair(1))
    Next i

    numShow = WorksheetFunction.Min(Me.ListBox1.ListCount, MAX_ROWS)
    With Me.ListBox1
        .Height = IIf(numShow > 0, numShow * ROW_HEIGHT + 1, 0)
        .ListIndex = -1
        .Visible = (numShow > 0)
    End With
End Sub
