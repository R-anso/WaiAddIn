Option Explicit

' ===== 侧边栏管理 =====
Private gSidebar As frmSidebar
Private gEventHandler As AppEventHandler

' 公开给 AppEventHandler 使用
Public gSidebarForm As Object

Public Sub ShowSidebar()
    On Error GoTo showFailed
    
    If gSidebar Is Nothing Then
        Set gSidebar = New frmSidebar
        Set gSidebarForm = gSidebar
        
        ' 启动窗口事件监听，实现侧边栏跟随
        If gEventHandler Is Nothing Then
            Set gEventHandler = New AppEventHandler
            Set gEventHandler.App = Application
        End If
    End If
    
    If Not gSidebar.Visible Then
        gSidebar.RefreshSelection
        gSidebar.Show vbModeless
        PositionSidebar
    End If
    
    SetAISetting "PanelVisible", "1"
    Exit Sub
    
showFailed:
    MsgBox "侧边栏启动失败：" & Err.Description, vbExclamation, "WAI"
End Sub

Private Sub PositionSidebar()
    On Error Resume Next
    ' 贴到文档窗口右侧，像真正的侧边栏
    gSidebar.Left = ActiveWindow.Left + ActiveWindow.Width - gSidebar.Width - 20
    gSidebar.Top = ActiveWindow.Top + 40
    ' 让侧边栏高度跟随文档窗口
    gSidebar.Height = ActiveWindow.Height - 80
    On Error GoTo 0
End Sub

Public Sub HideSidebar()
    On Error Resume Next
    If Not gSidebar Is Nothing Then
        gSidebar.Hide
    End If
    On Error GoTo 0
    SetAISetting "PanelVisible", "0"
End Sub

Public Function IsSidebarVisible() As Boolean
    On Error Resume Next
    If gSidebar Is Nothing Then
        IsSidebarVisible = False
    Else
        IsSidebarVisible = gSidebar.Visible
    End If
    On Error GoTo 0
End Function

Public Sub RefreshSidebarSelection()
    On Error Resume Next
    If Not gSidebar Is Nothing Then
        If gSidebar.Visible Then
            gSidebar.RefreshSelection
        End If
    End If
    On Error GoTo 0
End Sub

Public Sub ShowSidebarResponse(ByVal responseText As String, ByVal isEdit As Boolean)
    On Error Resume Next
    If gSidebar Is Nothing Then
        ShowSidebar
    End If
    
    If Not gSidebar.Visible Then
        gSidebar.Show vbModeless
    End If
    
    gSidebar.RefreshSelection
    gSidebar.ShowResponse responseText, isEdit
    On Error GoTo 0
End Sub

Public Sub SetSidebarBusy(ByVal isBusy As Boolean)
    On Error Resume Next
    If Not gSidebar Is Nothing Then
        gSidebar.SetBusy isBusy
    End If
    On Error GoTo 0
End Sub

Public Function GetSidebarEditMode() As Boolean
    On Error Resume Next
    If gSidebar Is Nothing Then
        GetSidebarEditMode = GetAISettingAsBoolean("EditModeDefault", True)
    Else
        GetSidebarEditMode = gSidebar.IsEditMode
    End If
    On Error GoTo 0
End Function

Public Function GetSidebarInstruction() As String
    On Error Resume Next
    If gSidebar Is Nothing Then
        GetSidebarInstruction = "润色这段文字"
    Else
        GetSidebarInstruction = gSidebar.SelectedInstruction
    End If
    On Error GoTo 0
End Function

' ===== 显示对话结果（兼容旧版，同时更新侧边栏） =====
Public Sub ShowChatResult(context As String, question As String, reply As String)
    gLastSelection = context
    gLastInstruction = question
    gLastReply = reply

    If Len(Trim$(reply)) = 0 Then
        MsgBox "没有收到有效回复。", vbExclamation, "AI"
        Exit Sub
    End If
    
    ShowSidebarResponse reply, GetSidebarEditMode()
End Sub

' ===== 清除对话历史 =====
Public Sub ClearHistory()
    ClearChatHistory
    gLastReply = ""
    gLastSelection = ""
    gLastInstruction = ""
    
    On Error Resume Next
    If Not gSidebar Is Nothing Then
        gSidebar.ShowResponse "", GetSidebarEditMode()
    End If
    On Error GoTo 0
    
    MsgBox "对话历史已清除", vbInformation, "AI"
End Sub

Public Sub InsertLastReplyToSelection()
    If Len(Trim$(gLastReply)) = 0 Then
        MsgBox "当前没有可插入的 AI 结果。", vbInformation, "AI"
        Exit Sub
    End If

    Selection.TypeText gLastReply
End Sub

Public Sub ReplaceSelectionWithLastReply()
    If Len(Trim$(gLastReply)) = 0 Then
        MsgBox "当前没有可替换的 AI 结果。", vbInformation, "AI"
        Exit Sub
    End If

    Selection.TypeText gLastReply
End Sub

Private Function PreviewText(valueText As String, Optional maxLength As Long = 300) As String
    Dim cleaned As String

    cleaned = Replace(valueText, vbCrLf, vbCrLf)
    cleaned = Replace(cleaned, vbCr, vbCrLf)
    cleaned = Replace(cleaned, vbLf, vbCrLf)

    If Len(cleaned) > maxLength Then
        PreviewText = Left$(cleaned, maxLength) & vbCrLf & "..."
    Else
        PreviewText = cleaned
    End If
End Function
