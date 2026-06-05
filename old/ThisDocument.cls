Option Explicit

Public Sub AutoExec()
    InitializeAISettings
    InstallAIKeyBindings
End Sub

Private Sub Document_Open()
    AutoExec
End Sub

Private Sub Document_Close()
    On Error Resume Next
    HideSidebar
    On Error GoTo 0
End Sub
