' ===== 配置区，只需改这里 =====
Option Explicit

' ===== 默认配置 =====
Public Const DEEPSEEK_DEFAULT_BASE_URL As String = "https://api.deepseek.com/chat/completions"
Public Const DEEPSEEK_DEFAULT_MODEL As String = "deepseek-v4-flash"
Private Const SETTING_PREFIX As String = "WAI_"
Private Const SETTING_API_KEY As String = "ApiKey"
Private Const SETTING_BASE_URL As String = "BaseUrl"
Private Const SETTING_MODEL As String = "Model"
Private Const SETTING_MAX_TOKENS As String = "MaxTokens"
Private Const SETTING_TEMPERATURE As String = "Temperature"
Private Const SETTING_PANEL_VISIBLE As String = "PanelVisible"
Private Const SETTING_BINDINGS_INSTALLED As String = "BindingsInstalled"
Private Const MAX_HISTORY_CHARS As Long = 12000

Public gChatHistory As String
Public gLastReply As String
Public gLastSelection As String
Public gLastInstruction As String

Public Sub InitializeAISettings()
    EnsureSetting SETTING_API_KEY, ""
    EnsureSetting SETTING_BASE_URL, DEEPSEEK_DEFAULT_BASE_URL
    EnsureSetting SETTING_MODEL, DEEPSEEK_DEFAULT_MODEL
    EnsureSetting SETTING_MAX_TOKENS, "2048"
    EnsureSetting SETTING_TEMPERATURE, "0.2"
    EnsureSetting SETTING_PANEL_VISIBLE, "1"
    EnsureSetting SETTING_BINDINGS_INSTALLED, "0"
End Sub

Public Sub ConfigureAISettings()
    Dim currentValue As String

    currentValue = InputBox("请输入 DeepSeek API Key（留空表示不修改）", "AI 设置", GetAISetting(SETTING_API_KEY, ""))
    If Len(Trim$(currentValue)) > 0 Then SetAISetting SETTING_API_KEY, Trim$(currentValue)

    currentValue = InputBox("请输入 API 基础地址", "AI 设置", GetAISetting(SETTING_BASE_URL, DEEPSEEK_DEFAULT_BASE_URL))
    If Len(Trim$(currentValue)) > 0 Then SetAISetting SETTING_BASE_URL, Trim$(currentValue)

    currentValue = InputBox("请输入模型名", "AI 设置", GetAISetting(SETTING_MODEL, DEEPSEEK_DEFAULT_MODEL))
    If Len(Trim$(currentValue)) > 0 Then SetAISetting SETTING_MODEL, Trim$(currentValue)

    currentValue = InputBox("请输入最大输出 token 数", "AI 设置", GetAISetting(SETTING_MAX_TOKENS, "2048"))
    If Len(Trim$(currentValue)) > 0 Then SetAISetting SETTING_MAX_TOKENS, Trim$(currentValue)

    currentValue = InputBox("请输入 temperature（例如 0.2）", "AI 设置", GetAISetting(SETTING_TEMPERATURE, "0.2"))
    If Len(Trim$(currentValue)) > 0 Then SetAISetting SETTING_TEMPERATURE, Trim$(currentValue)

    MsgBox "设置已保存。", vbInformation, "AI 设置"
End Sub

Public Function GetAISetting(settingName As String, Optional defaultValue As String = "") As String
    Dim variableName As String

    variableName = SETTING_PREFIX & settingName
    On Error Resume Next
    GetAISetting = ThisDocument.Variables(variableName).Value
    If Err.Number <> 0 Then
        Err.Clear
        GetAISetting = defaultValue
    End If
    On Error GoTo 0
End Function

Public Function GetAISettingAsLong(settingName As String, Optional defaultValue As Long = 0) As Long
    Dim valueText As String

    valueText = Trim$(GetAISetting(settingName, CStr(defaultValue)))
    If Len(valueText) = 0 Then
        GetAISettingAsLong = defaultValue
    ElseIf IsNumeric(valueText) Then
        GetAISettingAsLong = CLng(valueText)
    Else
        GetAISettingAsLong = defaultValue
    End If
End Function

Public Function GetAISettingAsDouble(settingName As String, Optional defaultValue As Double = 0#) As Double
    Dim valueText As String

    valueText = Trim$(GetAISetting(settingName, Replace$(Str$(defaultValue), " ", "")))
    If Len(valueText) = 0 Then
        GetAISettingAsDouble = defaultValue
    ElseIf IsNumeric(Replace$(valueText, ",", ".")) Then
        GetAISettingAsDouble = CDbl(Replace$(valueText, ",", "."))
    Else
        GetAISettingAsDouble = defaultValue
    End If
End Function

Public Function GetAISettingAsBoolean(settingName As String, Optional defaultValue As Boolean = False) As Boolean
    Dim valueText As String

    valueText = LCase$(Trim$(GetAISetting(settingName, IIf(defaultValue, "1", "0"))))
    GetAISettingAsBoolean = (valueText = "1" Or valueText = "true" Or valueText = "yes" Or valueText = "on")
End Function

Public Sub SetAISetting(settingName As String, settingValue As String)
    Dim variableName As String

    variableName = SETTING_PREFIX & settingName
    On Error Resume Next
    ThisDocument.Variables(variableName).Value = settingValue
    If Err.Number <> 0 Then
        Err.Clear
        ThisDocument.Variables.Add Name:=variableName, Value:=settingValue
    End If
    On Error GoTo 0
End Sub

Public Sub InstallAIKeyBindings(Optional force As Boolean = False)
    If Not force Then
        If GetAISetting(SETTING_BINDINGS_INSTALLED, "0") = "1" Then Exit Sub
    End If

    ' 快捷键选择原则：不与 Word 内置快捷键冲突
    ' 全部使用 Ctrl+Shift+Alt 四键组合，确保不会覆盖 Word 原生功能
    CustomizationContext = ThisDocument
    
    On Error Resume Next
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyW, wdKeyShift, wdKeyControl, wdKeyAlt), _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:="ToggleAIChatPanel"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyP, wdKeyShift, wdKeyControl, wdKeyAlt), _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:="PolishSelection"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyR, wdKeyShift, wdKeyControl, wdKeyAlt), _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:="RewriteSelection"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, wdKeyShift, wdKeyControl, wdKeyAlt), _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:="CorrectSelection"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyS, wdKeyShift, wdKeyControl, wdKeyAlt), _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:="SummarizeDocument"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeySpacebar, wdKeyShift, wdKeyControl, wdKeyAlt), _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:="AutocompleteFromContext"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyO, wdKeyShift, wdKeyControl, wdKeyAlt), _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:="OpenAISettings"
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyC, wdKeyShift, wdKeyControl, wdKeyAlt), _
                    KeyCategory:=wdKeyCategoryMacro, _
                    Command:="ChatWithSelection"
    On Error GoTo 0

    SetAISetting SETTING_BINDINGS_INSTALLED, "1"
End Sub

Public Sub RebindAIKeyBindings()
    Dim i As Long
    Dim removed As Long
    
    CustomizationContext = ThisDocument
    On Error Resume Next
    For i = KeyBindings.Count To 1 Step -1
        Select Case KeyBindings(i).Command
            Case "ChatWithSelection", "PolishSelection", "RewriteSelection", _
                 "CorrectSelection", "SummarizeDocument", "AutocompleteFromContext", _
                 "OpenAISettings", "ToggleAIChatPanel"
                KeyBindings(i).Clear
                removed = removed + 1
        End Select
    Next
    On Error GoTo 0
    
    SetAISetting SETTING_BINDINGS_INSTALLED, "0"
    InstallAIKeyBindings True
    
    MsgBox "快捷键已绑定（清除 " & removed & " 个旧绑定）。" & vbCrLf & vbCrLf & _
           "Ctrl+Shift+Alt+W  打开/关闭侧边栏" & vbCrLf & _
           "Ctrl+Shift+Alt+P  润色选区" & vbCrLf & _
           "Ctrl+Shift+Alt+R  改写选区" & vbCrLf & _
           "Ctrl+Shift+Alt+E  纠错选区" & vbCrLf & _
           "Ctrl+Shift+Alt+S  总结全文" & vbCrLf & _
           "Ctrl+Shift+Alt+O  打开设置" & vbCrLf & vbCrLf & _
           "所有快捷键均为 Ctrl+Shift+Alt 组合，不会覆盖 Word 原生功能。", vbInformation, "WAI"
End Sub

Public Function CallDeepSeek(systemPrompt As String, userMessage As String, _
                      Optional keepHistory As Boolean = False, _
                      Optional temperature As Variant) As String
    Dim http As Object
    Dim baseUrl As String
    Dim apiKey As String
    Dim modelName As String
    Dim requestBody As String
    Dim reply As String

    apiKey = Trim$(GetAISetting(SETTING_API_KEY, ""))
    If Len(apiKey) = 0 Then
        CallDeepSeek = "[错误] 尚未配置 API Key，请运行 AI 设置。"
        Exit Function
    End If

    baseUrl = Trim$(GetAISetting(SETTING_BASE_URL, DEEPSEEK_DEFAULT_BASE_URL))
    modelName = Trim$(GetAISetting(SETTING_MODEL, DEEPSEEK_DEFAULT_MODEL))

    Set http = CreateObject("MSXML2.XMLHTTP.6.0")

    requestBody = "{""model"":" & JsonQuote(modelName) & "," & _
                  Chr$(34) & "messages" & Chr$(34) & ":[{""role"":""system"",""content"":" & JsonQuote(CleanAiText(systemPrompt)) & "}," & _
                  BuildMessagesJson(userMessage, keepHistory) & "]," & _
                  Chr$(34) & "temperature" & Chr$(34) & ":" & JsonNumber(GetTemperatureValue(temperature)) & "," & _
                  Chr$(34) & "max_tokens" & Chr$(34) & ":" & CStr(GetAISettingAsLong(SETTING_MAX_TOKENS, 2048)) & "}"

    On Error GoTo RequestFailed
    http.Open "POST", baseUrl, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send requestBody

    If http.Status >= 200 And http.Status < 300 Then
        reply = ParseContent(http.responseText)
        gLastReply = reply
        If keepHistory Then UpdateChatHistory userMessage, reply
        CallDeepSeek = reply
    Else
        CallDeepSeek = "[错误 " & http.Status & "] " & http.responseText & vbCrLf & vbCrLf & "--- 请求体（前500字符）---" & vbCrLf & Left$(requestBody, 500)
    End If
    Exit Function

RequestFailed:
    CallDeepSeek = "[错误] 调用接口失败：" & Err.Description
End Function

Public Function BuildMessagesJson(userMessage As String, Optional keepHistory As Boolean = False) As String
    If keepHistory And Len(gChatHistory) > 0 Then
        BuildMessagesJson = gChatHistory & "," & "{""role"":""user"",""content"":" & JsonQuote(CleanAiText(userMessage)) & "}"
    Else
        BuildMessagesJson = "{""role"":""user"",""content"":" & JsonQuote(CleanAiText(userMessage)) & "}"
    End If
End Function

Public Sub UpdateChatHistory(userMessage As String, assistantMessage As String)
    Dim historyEntry As String

    historyEntry = "{""role"":""user"",""content"":" & JsonQuote(CleanAiText(userMessage)) & "}," & _
                   "{""role"":""assistant"",""content"":" & JsonQuote(CleanAiText(assistantMessage)) & "}"

    If Len(gChatHistory) = 0 Then
        gChatHistory = historyEntry
    Else
        gChatHistory = gChatHistory & "," & historyEntry
    End If

    If Len(gChatHistory) > MAX_HISTORY_CHARS Then
        gChatHistory = ""
    End If
End Sub

Public Sub ClearChatHistory()
    gChatHistory = ""
End Sub

Public Function ParseContent(json As String) As String
    Dim re As Object
    Dim matches As Object

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = False
    re.MultiLine = True
    re.Pattern = """content""\s*:\s*""((?:[^""\\]|\\.)*)"""

    If re.Test(json) Then
        Set matches = re.Execute(json)
        ParseContent = JsonUnescape(matches(0).SubMatches(0))
    Else
        ParseContent = json
    End If
End Function

Public Function JsonQuote(s As String) As String
    s = Replace(s, Chr$(92), Chr$(92) & Chr$(92))
    s = Replace(s, Chr$(34), Chr$(92) & Chr$(34))
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonQuote = Chr$(34) & s & Chr$(34)
End Function

Public Function JsonNumber(value As Double) As String
    Dim temp As String
    temp = Replace$(Trim$(Str$(value)), " ", "")
    If Left$(temp, 1) = "." Then temp = "0" & temp
    If Left$(temp, 1) = "-" And Mid$(temp, 2, 1) = "." Then temp = "-0" & Mid$(temp, 2)
    JsonNumber = temp
End Function

Private Function GetTemperatureValue(Optional temperature As Variant) As Double
    If IsMissing(temperature) Then
        GetTemperatureValue = GetAISettingAsDouble(SETTING_TEMPERATURE, 0.2)
    ElseIf IsNumeric(temperature) Then
        GetTemperatureValue = CDbl(temperature)
    Else
        GetTemperatureValue = GetAISettingAsDouble(SETTING_TEMPERATURE, 0.2)
    End If
End Function

Private Function JsonUnescape(s As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String
    Dim nextChar As String
    Dim hexCode As String

    i = 1
    Do While i <= Len(s)
        c = Mid$(s, i, 1)
        If c = Chr$(92) And i < Len(s) Then
            nextChar = Mid$(s, i + 1, 1)
            Select Case nextChar
                Case Chr$(34)
                    result = result & Chr$(34)
                    i = i + 2
                Case Chr$(92)
                    result = result & Chr$(92)
                    i = i + 2
                Case "/"
                    result = result & "/"
                    i = i + 2
                Case "b"
                    result = result & Chr$(8)
                    i = i + 2
                Case "f"
                    result = result & Chr$(12)
                    i = i + 2
                Case "n"
                    result = result & vbLf
                    i = i + 2
                Case "r"
                    result = result & vbCr
                    i = i + 2
                Case "t"
                    result = result & vbTab
                    i = i + 2
                Case "u"
                    If i + 5 <= Len(s) Then
                        hexCode = Mid$(s, i + 2, 4)
                        If IsHex4(hexCode) Then
                            result = result & ChrW$(CLng("&H" & hexCode))
                            i = i + 6
                        Else
                            result = result & nextChar
                            i = i + 2
                        End If
                    Else
                        result = result & nextChar
                        i = i + 2
                    End If
                Case Else
                    result = result & nextChar
                    i = i + 2
            End Select
        Else
            result = result & c
            i = i + 1
        End If
    Loop

    JsonUnescape = result
End Function

Private Function IsHex4(codeText As String) As Boolean
    Dim i As Long
    Dim c As String

    If Len(codeText) <> 4 Then Exit Function
    For i = 1 To 4
        c = Mid$(codeText, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", c, vbBinaryCompare) = 0 Then Exit Function
    Next i

    IsHex4 = True
End Function

Private Sub EnsureSetting(settingName As String, defaultValue As String)
    Dim variableName As String
    Dim currentValue As String

    variableName = SETTING_PREFIX & settingName
    On Error Resume Next
    currentValue = ThisDocument.Variables(variableName).Value
    If Err.Number <> 0 Then
        Err.Clear
        ThisDocument.Variables.Add Name:=variableName, Value:=defaultValue
    ElseIf Len(currentValue) = 0 Then
        ThisDocument.Variables(variableName).Value = defaultValue
    End If
    On Error GoTo 0
End Sub

Private Function CleanAiText(valueText As String) As String
    Dim i As Long
    Dim code As Long
    Dim c As String
    Dim cleaned As String

    For i = 1 To Len(valueText)
        c = Mid$(valueText, i, 1)
        code = AscW(c)
        Select Case code
            Case 0
                cleaned = cleaned
            Case 9
                cleaned = cleaned & vbTab
            Case 10, 13
                cleaned = cleaned & vbLf
            Case 7, 8, 11, 12, 14 To 31
                cleaned = cleaned & " "
            Case 127 To 159
                cleaned = cleaned & " "
            Case 8232, 8233
                cleaned = cleaned & vbLf
            Case Else
                cleaned = cleaned & c
        End Select
    Next i

    CleanAiText = cleaned
End Function



