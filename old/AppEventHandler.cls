Option Explicit

' ================================================================
' 应用事件监听：Word 窗口移动/缩放时自动重新定位侧边栏
' 使用方法：VBA 编辑器 → 插入 → 类模块 → 命名为 AppEventHandler
' ================================================================

Public WithEvents App As Word.Application

Private Sub App_WindowSize(ByVal Doc As Document, ByVal Wn As Window)
    RepositionSidebar
End Sub

Private Sub App_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
    RepositionSidebar
End Sub

Public Sub RepositionSidebar()
    On Error Resume Next
    If Not gSidebarForm Is Nothing Then
        If gSidebarForm.Visible Then
            gSidebarForm.Left = ActiveWindow.Left + ActiveWindow.Width - gSidebarForm.Width - 20
            gSidebarForm.Top = ActiveWindow.Top + 40
        End If
    End If
    On Error GoTo 0
End Sub
