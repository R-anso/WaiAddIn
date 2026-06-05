Option Explicit

' ================================================================
' 使用说明：
' 1. 在 VBA 编辑器中：菜单「插入」→「用户窗体」
' 2. 在属性窗口把窗体名称改为 frmSidebar，Caption 改为 "WAI"
' 3. 双击窗体打开代码窗口，把本文件全部内容粘贴进去
' 4. 保存。窗体控件会在运行时自动创建。
' ================================================================

' ===== 控件：运行时动态创建 =====
Private WithEvents txtSelection As MSForms.TextBox
Private WithEvents cboInstruction As MSForms.ComboBox
Private WithEvents btnSend As MSForms.CommandButton
Private WithEvents txtResponse As MSForms.TextBox
Private WithEvents btnAccept As MSForms.CommandButton
Private WithEvents btnReject As MSForms.CommandButton
Private WithEvents btnInsert As MSForms.CommandButton
Private WithEvents btnNewChat As MSForms.CommandButton
Private WithEvents btnSettings As MSForms.CommandButton
Private WithEvents cboModel As MSForms.ComboBox
Private WithEvents optEditMode As MSForms.OptionButton
Private WithEvents optAnswerMode As MSForms.OptionButton
Private WithEvents lblSelection As MSForms.Label
Private WithEvents lblInstruction As MSForms.Label
Private WithEvents lblResponse As MSForms.Label
Private WithEvents lblModel As MSForms.Label
Private WithEvents btnRefresh As MSForms.CommandButton

Private pIsEditMode As Boolean
Private pLastUserMessage As String
Private pLastReply As String
Private pIsUpdating As Boolean

Private Const FORM_PADDING As Long = 12
Private Const BUTTON_HEIGHT As Long = 24
Private Const LABEL_HEIGHT As Long = 16

' 预设指令
Private Function PresetInstructions() As String()
    PresetInstructions = Split( _
        "润色这段文字|改写这段文字|纠错这段文字|翻译成英文|总结要点|解释这段内容|扩写为200字|压缩到100字以内", "|")
End Function

Private Sub UserForm_Initialize()
    pIsEditMode = GetAISettingAsBoolean("EditModeDefault", True)
    pIsUpdating = True
    BuildUI
    LoadSettings
    pIsUpdating = False
End Sub

Private Sub BuildUI()
    Dim yPos As Long
    Dim ctrlWidth As Long
    
    ctrlWidth = Me.Width - FORM_PADDING * 2
    
    ' 选区行：标签 + 刷新按钮
    yPos = FORM_PADDING
    Set lblSelection = AddLabel("lblSelection", FORM_PADDING, yPos, ctrlWidth - 80, LABEL_HEIGHT, "选区：未选中", True)
    Set btnRefresh = AddButton("btnRefresh", FORM_PADDING + ctrlWidth - 76, yPos - 2, 76, 20, "抓取选区")
    yPos = yPos + LABEL_HEIGHT + 4
    
    ' 选区文本框
    Set txtSelection = AddTextBox("txtSelection", FORM_PADDING, yPos, ctrlWidth, 80, True, True)
    yPos = yPos + 80 + 8
    
    ' 编辑/问答 模式
    Set optEditMode = AddOptionButton("optEditMode", FORM_PADDING, yPos, 80, LABEL_HEIGHT, "编辑模式", True, True)
    Set optAnswerMode = AddOptionButton("optAnswerMode", FORM_PADDING + 100, yPos, 80, LABEL_HEIGHT, "问答模式", False, True)
    yPos = yPos + LABEL_HEIGHT + 8
    
    ' 指令行
    Set lblInstruction = AddLabel("lblInstruction", FORM_PADDING, yPos, ctrlWidth, LABEL_HEIGHT, "指令：", False)
    yPos = yPos + LABEL_HEIGHT + 2
    
    Set cboInstruction = AddComboBox("cboInstruction", FORM_PADDING, yPos, ctrlWidth - 80, True)
    Dim preset
    For Each preset In PresetInstructions()
        cboInstruction.AddItem preset
    Next
    cboInstruction.Text = "润色这段文字"
    
    Set btnSend = AddButton("btnSend", FORM_PADDING + ctrlWidth - 76, yPos, 76, BUTTON_HEIGHT, "发送")
    yPos = yPos + BUTTON_HEIGHT + 8
    
    ' 回复区
    Set lblResponse = AddLabel("lblResponse", FORM_PADDING, yPos, ctrlWidth, LABEL_HEIGHT, "AI 回复：", False)
    yPos = yPos + LABEL_HEIGHT + 2
    
    Set txtResponse = AddTextBox("txtResponse", FORM_PADDING, yPos, ctrlWidth, 260, True, True)
    yPos = yPos + 260 + 8
    
    ' 操作按钮行
    Set btnAccept = AddButton("btnAccept", FORM_PADDING, yPos, 90, BUTTON_HEIGHT, "接受修改")
    Set btnReject = AddButton("btnReject", FORM_PADDING + 96, yPos, 90, BUTTON_HEIGHT, "放弃")
    Set btnInsert = AddButton("btnInsert", FORM_PADDING + 192, yPos, 90, BUTTON_HEIGHT, "插入光标")
    yPos = yPos + BUTTON_HEIGHT + 8
    
    ' 底部工具栏
    Set btnNewChat = AddButton("btnNewChat", FORM_PADDING, yPos, 88, BUTTON_HEIGHT, "新对话")
    Set btnSettings = AddButton("btnSettings", FORM_PADDING + 94, yPos, 88, BUTTON_HEIGHT, "设置")
    
    Set lblModel = AddLabel("lblModel", FORM_PADDING + 200, yPos + 4, 36, LABEL_HEIGHT, "模型：", False)
    Set cboModel = AddComboBox("cboModel", FORM_PADDING + 240, yPos, ctrlWidth - 240, False)
    cboModel.AddItem "deepseek-v4-pro"
    cboModel.AddItem "deepseek-v4-flash"
    cboModel.Text = GetAISetting("Model", "deepseek-v4-flash")
    
    UpdateModeUI
End Sub

Private Sub LoadSettings()
    optEditMode.Value = GetAISettingAsBoolean("EditModeDefault", True)
    optAnswerMode.Value = Not optEditMode.Value
    pIsEditMode = optEditMode.Value
End Sub

Private Sub UpdateModeUI()
    If pIsEditMode Then
        btnAccept.Enabled = True
        btnReject.Enabled = True
        Me.Caption = "WAI — 编辑模式"
    Else
        btnAccept.Enabled = False
        btnReject.Enabled = False
        Me.Caption = "WAI — 问答模式"
    End If
End Sub

' ===== 按钮事件 =====
Private Sub btnRefresh_Click()
    RefreshSelection
End Sub

Private Sub btnSend_Click()
    Dim sourceText As String
    Dim instruction As String
    Dim prompt As String
    Dim reply As String
    
    RefreshSelection
    sourceText = Trim$(txtSelection.Text)
    
    If Len(sourceText) = 0 Then
        txtResponse.Text = "请先在 Word 中选中文字，再点击「抓取选区」。"
        Exit Sub
    End If
    
    instruction = Trim$(cboInstruction.Text)
    If Len(instruction) = 0 Then instruction = "请处理这段文字"
    
    On Error Resume Next
    Dim i As Long, found As Boolean
    For i = 0 To cboInstruction.ListCount - 1
        If cboInstruction.List(i) = instruction Then found = True: Exit For
    Next
    If Not found Then cboInstruction.AddItem instruction, 0
    On Error GoTo 0
    
    gLastSelection = sourceText
    gLastInstruction = instruction
    
    prompt = "请根据以下原文执行用户要求。" & vbCrLf & _
             "要求：" & instruction & vbCrLf & _
             "原文：" & vbCrLf & sourceText
    
    pLastUserMessage = prompt
    
    Me.Caption = "WAI — 思考中..."
    txtResponse.Text = "思考中..."
    btnSend.Enabled = False
    DoEvents
    
    reply = CallDeepSeek(AI_SYSTEM_PROMPT, prompt, True)
    
    pLastReply = reply
    txtResponse.Text = reply
    btnSend.Enabled = True
    UpdateModeUI
    Me.Caption = IIf(pIsEditMode, "WAI — 编辑模式", "WAI — 问答模式")
End Sub

Private Sub btnAccept_Click()
    If Len(Trim$(pLastReply)) = 0 Then Exit Sub
    Selection.TypeText pLastReply
    txtResponse.Text = "已应用修改。可以继续选中其他内容操作。"
End Sub

Private Sub btnReject_Click()
    pLastReply = ""
    txtResponse.Text = ""
End Sub

Private Sub btnInsert_Click()
    If Len(Trim$(txtResponse.Text)) = 0 Then
        MsgBox "没有可插入的内容", vbInformation
        Exit Sub
    End If
    Selection.TypeText txtResponse.Text
End Sub

Private Sub btnNewChat_Click()
    ClearChatHistory
    gLastReply = ""
    gLastSelection = ""
    gLastInstruction = ""
    pLastReply = ""
    txtResponse.Text = ""
End Sub

Private Sub btnSettings_Click()
    ConfigureAISettings
    cboModel.Text = GetAISetting("Model", "deepseek-v4-flash")
End Sub

Private Sub optEditMode_Click()
    If pIsUpdating Then Exit Sub
    pIsEditMode = True
    SetAISetting "EditModeDefault", "1"
    UpdateModeUI
End Sub

Private Sub optAnswerMode_Click()
    If pIsUpdating Then Exit Sub
    pIsEditMode = False
    SetAISetting "EditModeDefault", "0"
    UpdateModeUI
End Sub

Private Sub cboModel_Change()
    If pIsUpdating Then Exit Sub
    Dim newModel As String
    newModel = Trim$(cboModel.Text)
    If Len(newModel) > 0 Then SetAISetting "Model", newModel
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        SetAISetting "PanelVisible", "0"
    End If
End Sub

' ===== 公共方法（供 UI.bas 调用） =====
Public Sub RefreshSelection()
    Dim selText As String
    selText = Selection.Text
    
    If Len(Trim$(selText)) > 0 Then
        lblSelection.Caption = "选区：已选中 " & Len(selText) & " 字"
        txtSelection.Text = selText
    Else
        lblSelection.Caption = "选区：未选中"
    End If
End Sub

Public Sub ShowResponse(ByVal responseText As String, ByVal isEdit As Boolean)
    pLastReply = responseText
    pIsEditMode = isEdit
    txtResponse.Text = responseText
    UpdateModeUI
End Sub

Public Sub SetBusy(ByVal isBusy As Boolean)
    btnSend.Enabled = Not isBusy
    If isBusy Then
        Me.Caption = "WAI — 思考中..."
    Else
        Me.Caption = IIf(pIsEditMode, "WAI — 编辑模式", "WAI — 问答模式")
    End If
End Sub

Public Property Get IsEditMode() As Boolean
    IsEditMode = pIsEditMode
End Property

Public Property Get SelectedInstruction() As String
    SelectedInstruction = Trim$(cboInstruction.Text)
End Property

' ===== 动态创建控件的辅助函数 =====
Private Function AddLabel(ByVal ctrlName As String, ByVal cLeft As Long, ByVal cTop As Long, ByVal cWidth As Long, ByVal cHeight As Long, ByVal cCaption As String, ByVal cBold As Boolean) As MSForms.Label
    Dim ctrl As MSForms.Label
    Set ctrl = Me.Controls.Add("Forms.Label.1", ctrlName, True)
    ctrl.Left = cLeft
    ctrl.Top = cTop
    ctrl.Width = cWidth
    ctrl.Height = cHeight
    ctrl.Caption = cCaption
    ctrl.Font.Bold = cBold
    Set AddLabel = ctrl
End Function

Private Function AddTextBox(ByVal ctrlName As String, ByVal cLeft As Long, ByVal cTop As Long, ByVal cWidth As Long, ByVal cHeight As Long, ByVal cMultiLine As Boolean, ByVal cScrollBars As Boolean) As MSForms.TextBox
    Dim ctrl As MSForms.TextBox
    Set ctrl = Me.Controls.Add("Forms.TextBox.1", ctrlName, True)
    ctrl.Left = cLeft
    ctrl.Top = cTop
    ctrl.Width = cWidth
    ctrl.Height = cHeight
    ctrl.MultiLine = cMultiLine
    If cScrollBars Then
        ctrl.ScrollBars = fmScrollBarsVertical
    End If
    ctrl.Locked = True
    ctrl.BackColor = &HF5F5F5
    Set AddTextBox = ctrl
End Function

Private Function AddComboBox(ByVal ctrlName As String, ByVal cLeft As Long, ByVal cTop As Long, ByVal cWidth As Long, ByVal cEditable As Boolean) As MSForms.ComboBox
    Dim ctrl As MSForms.ComboBox
    Set ctrl = Me.Controls.Add("Forms.ComboBox.1", ctrlName, True)
    ctrl.Left = cLeft
    ctrl.Top = cTop
    ctrl.Width = cWidth
    ctrl.Height = BUTTON_HEIGHT
    If Not cEditable Then
        ctrl.Style = fmStyleDropDownList
    Else
        ctrl.Style = fmStyleDropDownCombo
    End If
    Set AddComboBox = ctrl
End Function

Private Function AddButton(ByVal ctrlName As String, ByVal cLeft As Long, ByVal cTop As Long, ByVal cWidth As Long, ByVal cHeight As Long, ByVal cCaption As String) As MSForms.CommandButton
    Dim ctrl As MSForms.CommandButton
    Set ctrl = Me.Controls.Add("Forms.CommandButton.1", ctrlName, True)
    ctrl.Left = cLeft
    ctrl.Top = cTop
    ctrl.Width = cWidth
    ctrl.Height = cHeight
    ctrl.Caption = cCaption
    Set AddButton = ctrl
End Function

Private Function AddOptionButton(ByVal ctrlName As String, ByVal cLeft As Long, ByVal cTop As Long, ByVal cWidth As Long, ByVal cHeight As Long, ByVal cCaption As String, ByVal cValue As Boolean, ByVal cGroup As Boolean) As MSForms.OptionButton
    Dim ctrl As MSForms.OptionButton
    Set ctrl = Me.Controls.Add("Forms.OptionButton.1", ctrlName, True)
    ctrl.Left = cLeft
    ctrl.Top = cTop
    ctrl.Width = cWidth
    ctrl.Height = cHeight
    ctrl.Caption = cCaption
    ctrl.Value = cValue
    If cGroup Then ctrl.GroupName = "ModeGroup"
    Set AddOptionButton = ctrl
End Function
