Option Explicit

Public Const AI_SYSTEM_PROMPT As String = "你是一个中文文档助手，擅长润色、改写、纠错、总结和续写。回答要直接、准确、尽量保留原文结构。"

' ===== 快捷键入口 =====
Public Sub ChatWithSelection()
    If Not IsSidebarVisible() Then ShowSidebar
    RefreshSidebarSelection
End Sub

Public Sub PolishSelection()
    If Not IsSidebarVisible() Then ShowSidebar
    ProcessSelectionTask "润色选区", "请在不改变原意的前提下润色这段文字，使表达更自然、更专业。只输出润色后的结果。", True, False
End Sub

Public Sub RewriteSelection()
    If Not IsSidebarVisible() Then ShowSidebar
    ProcessSelectionTask "改写选区", "请改写这段文字，保持原意不变，但让表达更清晰、更简洁。只输出改写结果。", True, False
End Sub

Public Sub CorrectSelection()
    If Not IsSidebarVisible() Then ShowSidebar
    ProcessSelectionTask "纠错选区", "请检查并修正这段文字中的错别字、语法、标点和明显逻辑问题。只输出修正后的文本。", True, False
End Sub

Public Sub TranslateSelection()
    If Not IsSidebarVisible() Then ShowSidebar
    ProcessSelectionTask "翻译选区", "请将这段文字翻译成英文，保留原有格式和术语。只输出译文。", True, False
End Sub

Public Sub SummarizeSelection()
    If Not IsSidebarVisible() Then ShowSidebar
    ProcessSelectionTask "总结选区", "请将这段内容总结为 3 到 5 条要点，保留核心信息。", False, False
End Sub

Public Sub SummarizeDocument()
    Dim fullText As String
    Dim reply As String
    Dim prompt As String

    fullText = ActiveDocument.Content.Text
    If Len(Trim$(fullText)) < 50 Then
        MsgBox "文档内容太少", vbInformation
        Exit Sub
    End If

    If Not IsSidebarVisible() Then ShowSidebar

    prompt = "请对以下文档进行结构化总结，包含主要议题、关键论点和重要结论。用中文回答，直接输出结果。" & vbCrLf & vbCrLf & fullText

    Application.StatusBar = "AI 正在总结全文..."
    SetSidebarBusy True
    reply = CallDeepSeek(AI_SYSTEM_PROMPT, prompt, False)
    SetSidebarBusy False
    Application.StatusBar = False

    ShowChatResult "（全文）", "总结文档", reply
End Sub

Public Sub AutocompleteFromContext()
    Dim context As String
    Dim prompt As String
    Dim reply As String
    Dim confirm As VbMsgBoxResult

    context = GetContextBeforeSelection(2000)
    If Len(Trim$(context)) < 20 Then
        MsgBox "上文太少，请先写一些内容", vbInformation
        Exit Sub
    End If

    If Not IsSidebarVisible() Then ShowSidebar
    RefreshSidebarSelection

    prompt = "请根据以下上下文续写下一段，保持文风一致，长度大约 100 到 150 字，只输出续写内容，不要解释。" & vbCrLf & vbCrLf & context

    Application.StatusBar = "AI 正在补全内容..."
    SetSidebarBusy True
    reply = CallDeepSeek(AI_SYSTEM_PROMPT, prompt, False)
    SetSidebarBusy False
    Application.StatusBar = False

    ShowChatResult "（上文续写）", "根据上下文补全", reply
End Sub

' ===== 面板和设置 =====
Public Sub OpenAISettings()
    ConfigureAISettings
End Sub

Public Sub ToggleAIChatPanel()
    If IsSidebarVisible() Then
        HideSidebar
        MsgBox "侧边栏已关闭", vbInformation, "WAI"
    Else
        ShowSidebar
    End If
End Sub

' ===== 内部处理 =====
Private Sub ProcessSelectionTask(actionTitle As String, instruction As String, Optional replaceSelection As Boolean = False, Optional keepHistory As Boolean = False)
    Dim sourceText As String
    Dim prompt As String
    Dim reply As String

    sourceText = GetSelectedText()
    If Len(Trim$(sourceText)) = 0 Then
        MsgBox "请先选中文字", vbInformation, actionTitle
        Exit Sub
    End If

    gLastSelection = sourceText
    gLastInstruction = instruction
    RefreshSidebarSelection

    prompt = "请根据以下原文执行用户要求。" & vbCrLf & _
             "要求：" & instruction & vbCrLf & _
             "原文：" & vbCrLf & sourceText

    Application.StatusBar = "AI 正在处理：" & actionTitle & "..."
    SetSidebarBusy True
    reply = CallDeepSeek(AI_SYSTEM_PROMPT, prompt, keepHistory)
    SetSidebarBusy False
    Application.StatusBar = False

    ShowChatResult sourceText, instruction, reply
End Sub

Private Function GetSelectedText() As String
    GetSelectedText = Selection.Text
End Function

Private Function GetContextBeforeSelection(Optional maxChars As Long = 2000) As String
    Dim rng As Range
    Dim startPos As Long

    Set rng = Selection.Range.Duplicate
    startPos = rng.Start - maxChars
    If startPos < 0 Then startPos = 0
    rng.Start = startPos
    GetContextBeforeSelection = rng.Text
End Function

Public Sub InsertAIText(textToInsert As String, replaceSelection As Boolean)
    If replaceSelection Then
        Selection.TypeText textToInsert
    Else
        Selection.TypeText textToInsert
    End If
End Sub
