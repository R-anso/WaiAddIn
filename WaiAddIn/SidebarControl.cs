using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WaiAddIn
{
    public partial class SidebarControl : UserControl
    {
        // ---- 控件 ----
        private Label lblHistory;
        private ComboBox cboHistory;
        private Button btnDeleteSession;
        private Label lblSelection;
        private Button btnRefresh;
        private TextBox txtSelection;
        private RadioButton optEditMode;
        private RadioButton optAnswerMode;
        private Label lblInstruction;
        private TextBox txtPrompt;
        private Button btnSend;
        private Button btnAutocomplete;
        private Label lblResponse;
        private TextBox txtResponse;
        private Button btnAccept;
        private Button btnReject;
        private Button btnInsert;
        private Button btnNewChat;
        private Button btnSettings;
        private Label lblModel;
        private ComboBox cboModel;
        private Panel pnlTopDivider;
        private Panel pnlBottomDivider;

        // ---- 状态 ----
        private bool _isEditMode = true;
        private string _lastReply = "";
        private bool _isUpdating = false;
        private bool _uiBuilt = false;
        private bool _isLoadingHistory = false;
        private SettingsData _pendingSettings;
        private int _selectionStart = -1;
        private int _selectionEnd = -1;
        private List<ChatSessionData> _sessions = new List<ChatSessionData>();
        private ChatSessionData _currentSession;

        // ---- 服务 ----
        private AIService _aiService;

        private const int PAD = 10;
        private const int BTN_H = 30;
        private const int LBL_H = 28;
        private const int HISTORY_ROW_H = 30;
        private const int SELECTION_BOX_H = 96;
        private const int DEFAULT_PROMPT_BOX_H = 140;
        private const int DEFAULT_RESPONSE_BOX_H = 210;
        private const int ACTION_ROW_H = 30;
        private const int MODEL_ROW_H = 30;

        private static readonly string[] Presets = {
            "润色这段文字", "改写这段文字", "纠错这段文字",
            "翻译成英文", "总结要点", "解释这段内容",
            "扩写为200字", "压缩到100字以内"
        };

        private static readonly string SystemPrompt =
            "你是一个中文文档助手，擅长润色、改写、纠错、总结和续写。回答要直接、准确、尽量保留原文结构。";

        public SidebarControl()
        {
            InitializeComponent();
            this.Load += SidebarControl_Load;
            this.Resize += SidebarControl_Resize;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            this.Dock = DockStyle.Fill;
            this.AutoScroll = true;
            this.AutoScrollMinSize = new Size(0, 680);
            this.BackColor = Color.FromArgb(250, 250, 250);
            this.ForeColor = Color.FromArgb(35, 35, 35);
            this.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
            this.Padding = new Padding(PAD);
            this.MinimumSize = new Size(320, 500);
            this.DoubleBuffered = true;
            this.ResumeLayout(false);
        }

        private void SidebarControl_Load(object sender, EventArgs e)
        {
            if (!_uiBuilt)
            {
                BuildUI();
                _uiBuilt = true;
            }

            LoadHistory();
            ApplyPendingSettings();
            LayoutUI();
        }

        private void SidebarControl_Resize(object sender, EventArgs e)
        {
            if (_uiBuilt)
            {
                LayoutUI();
            }
        }

        // ======================== 设置 AI 服务 ========================
        public void SetAIService(AIService service)
        {
            _aiService = service;
        }

        public void LoadSettings(SettingsData settings)
        {
            _pendingSettings = settings;

            if (!_uiBuilt)
            {
                return;
            }

            ApplyPendingSettings();
        }

        private void ApplyPendingSettings()
        {
            if (_pendingSettings == null || !_uiBuilt)
            {
                return;
            }

            _isUpdating = true;
            _isEditMode = _pendingSettings.EditModeDefault;
            optEditMode.Checked = _isEditMode;
            optAnswerMode.Checked = !_isEditMode;
            cboModel.Text = _pendingSettings.Model;
            btnAutocomplete.Visible = _pendingSettings.EnableAutocomplete;
            _isUpdating = false;
            UpdateModeUI();
            ApplyConfiguredHeights();

            if (!string.IsNullOrWhiteSpace(_pendingSettings.LastSessionId))
            {
                SelectSession(_pendingSettings.LastSessionId);
            }
            else if (_sessions.Count > 0)
            {
                SelectSession(_sessions[0].Id);
            }
        }

        private void ApplyConfiguredHeights()
        {
            if (_pendingSettings == null)
            {
                return;
            }

            if (txtSelection != null)
                txtSelection.Height = Clamp(_pendingSettings.SelectionBoxHeight, 48, 260);
            if (txtPrompt != null)
                txtPrompt.Height = Clamp(_pendingSettings.PromptBoxHeight, 60, 360);
            if (txtResponse != null)
                txtResponse.Height = Clamp(_pendingSettings.ResponseBoxHeight, 80, 420);
        }

        // ======================== 公共方法 ========================
        public void RefreshSelection(string selectedText)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(RefreshSelection), selectedText);
                return;
            }

            if (!string.IsNullOrWhiteSpace(selectedText))
            {
                lblSelection.Text = $"选区：已选中 {selectedText.Length} 字";
                txtSelection.Text = selectedText;
            }
            else
            {
                lblSelection.Text = "选区：未选中";
            }
        }

        public void SetBusy(bool busy)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<bool>(SetBusy), busy);
                return;
            }
            btnSend.Enabled = !busy;
        }

        private void RunOnUiThread(Action action)
        {
            if (this.IsDisposed)
                return;

            if (this.InvokeRequired)
            {
                this.BeginInvoke(action);
                return;
            }

            action();
        }

        // ======================== UI 构建 ========================
        private void BuildUI()
        {
            int y = PAD;
            int w = Math.Max(280, this.ClientSize.Width - PAD * 2);

            // 会话栏
            lblHistory = AddLabel("会话：", PAD, y, 46, LBL_H, false);
            cboHistory = new ComboBox
            {
                Left = PAD + 48,
                Top = y + 1,
                Width = Math.Max(120, w - 180),
                Height = HISTORY_ROW_H,
                FlatStyle = FlatStyle.System,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboHistory.SelectedIndexChanged += (s, e) => {
                if (!_isLoadingHistory)
                {
                    SelectCurrentHistoryFromUi();
                }
            };
            this.Controls.Add(cboHistory);

            btnNewChat = AddButton("新会话", PAD + w - 124, y, 58, HISTORY_ROW_H);
            btnNewChat.Click += (s, e) => StartNewSession();
            btnDeleteSession = AddButton("删除", PAD + w - 60, y, 60, HISTORY_ROW_H);
            btnDeleteSession.Click += (s, e) => DeleteCurrentSession();
            y += HISTORY_ROW_H + 8;

            // 选区标签 + 刷新
            lblSelection = AddLabel("选区：未选中", PAD, y, w - 86, LBL_H, true);
            btnRefresh = AddButton("抓取选区", PAD + w - 80, y - 1, 80, BTN_H);
            btnRefresh.Click += (s, e) => RefreshFromWord();
            y += LBL_H + 6;

            // 选区文本框
            txtSelection = AddTextBox(PAD, y, w, SELECTION_BOX_H, true, true, true);
            y += SELECTION_BOX_H + 8;

            // 模式
            optEditMode = AddRadioButton("编辑模式", PAD, y, 92, LBL_H, true);
            optAnswerMode = AddRadioButton("问答模式", PAD + 104, y, 92, LBL_H, false);
            optEditMode.CheckedChanged += (s, e) => { if (!_isUpdating) { _isEditMode = true; SaveMode(); UpdateModeUI(); } };
            optAnswerMode.CheckedChanged += (s, e) => { if (!_isUpdating) { _isEditMode = false; SaveMode(); UpdateModeUI(); } };
            y += LBL_H + 8;

            pnlTopDivider = AddDivider(PAD, y, w);
            y += 10;

            // Prompt
            lblInstruction = AddLabel("Prompt：", PAD, y, w, LBL_H, false);
            y += LBL_H + 1;
            txtPrompt = AddTextBox(PAD, y, w, DEFAULT_PROMPT_BOX_H, true, true, false);
            txtPrompt.Text = "请根据下面的原文和当前会话历史继续处理内容。";
            txtPrompt.WordWrap = true;
            y += DEFAULT_PROMPT_BOX_H + 8;

            btnSend = AddButton("发送", PAD, y, 80, BTN_H);
            btnSend.Click += BtnSend_Click;
            btnAutocomplete = AddButton("自动补全", PAD + 86, y, 90, BTN_H);
            btnAutocomplete.Click += BtnAutocomplete_Click;
            y += BTN_H + 10;

            pnlBottomDivider = AddDivider(PAD, y, w);
            y += 10;

            // 回复
            lblResponse = AddLabel("AI 回复：", PAD, y, w, LBL_H, bold: false);
            y += LBL_H + 1;
            txtResponse = AddTextBox(PAD, y, w, DEFAULT_RESPONSE_BOX_H, true, true, true);
            txtResponse.Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point);
            txtResponse.WordWrap = true;
            y += DEFAULT_RESPONSE_BOX_H + 8;

            // 操作按钮行
            btnAccept = AddButton("接受修改", PAD, y, 92, ACTION_ROW_H);
            btnAccept.Click += (s, e) => AcceptEdit();
            btnReject = AddButton("放弃", PAD + 98, y, 92, ACTION_ROW_H);
            btnReject.Click += (s, e) => { _lastReply = ""; txtResponse.Text = ""; };
            btnInsert = AddButton("插入光标", PAD + 196, y, 92, ACTION_ROW_H);
            btnInsert.Click += (s, e) => InsertAtCursor();
            y += ACTION_ROW_H + 8;

            // 底部
            btnSettings = AddButton("设置", PAD, y, 92, MODEL_ROW_H);
            btnSettings.Click += (s, e) => OpenSettings();
            lblModel = AddLabel("模型：", PAD + 102, y, 48, LBL_H, false);
            cboModel = new ComboBox { Left = PAD + 152, Top = y, Width = Math.Max(120, w - 152), Height = MODEL_ROW_H, FlatStyle = FlatStyle.System, DropDownStyle = ComboBoxStyle.DropDownList };
            cboModel.Items.AddRange(new[] { "deepseek-v4-pro", "deepseek-v4-flash" });
            cboModel.Text = "deepseek-v4-flash";
            cboModel.SelectedIndexChanged += (s, e) => { if (!_isUpdating && _aiService != null) SaveModel(); };
            this.Controls.Add(cboModel);

            UpdateModeUI();
        }

        private void LayoutUI()
        {
            int w = Math.Max(280, this.ClientSize.Width - PAD * 2);
            int h = Math.Max(560, this.ClientSize.Height - PAD * 2);
            int currentY = PAD;
            int bottomGap = 8;

            if (cboHistory != null)
                cboHistory.Width = Math.Max(120, w - 180);
            if (btnNewChat != null)
                btnNewChat.Left = PAD + w - 124;
            if (btnDeleteSession != null)
                btnDeleteSession.Left = PAD + w - 60;

            if (lblSelection != null)
                lblSelection.Width = Math.Max(0, w - 90);
            if (btnRefresh != null)
                btnRefresh.Left = PAD + w - 80;
            if (lblSelection != null)
                lblSelection.Top = currentY;
            if (btnRefresh != null)
                btnRefresh.Top = currentY - 1;

            currentY += LBL_H + 8;

            if (txtSelection != null)
            {
                txtSelection.Top = currentY;
                txtSelection.Width = w;
                txtSelection.Height = _pendingSettings != null ? Clamp(_pendingSettings.SelectionBoxHeight, 48, 260) : SELECTION_BOX_H;
            }
            currentY += (txtSelection != null ? txtSelection.Height : SELECTION_BOX_H) + 8;

            if (optAnswerMode != null)
            {
                optEditMode.Top = currentY;
                optAnswerMode.Top = currentY;
                optAnswerMode.Left = PAD + 104;
            }
            currentY += LBL_H + 8;

            if (pnlTopDivider != null)
            {
                pnlTopDivider.Left = PAD;
                pnlTopDivider.Top = currentY;
                pnlTopDivider.Width = w;
            }
            currentY += 10;

            if (lblInstruction != null)
                lblInstruction.Top = currentY;
            if (txtPrompt != null)
            {
                txtPrompt.Top = currentY + LBL_H + 2;
                txtPrompt.Width = w;
                txtPrompt.Height = _pendingSettings != null ? Clamp(_pendingSettings.PromptBoxHeight, 60, 360) : DEFAULT_PROMPT_BOX_H;
            }
            if (btnSend != null)
            {
                btnSend.Left = PAD;
                btnSend.Top = (txtPrompt != null ? txtPrompt.Bottom : currentY + DEFAULT_PROMPT_BOX_H) + 6;
            }
            if (btnAutocomplete != null)
            {
                btnAutocomplete.Left = PAD + 86;
                btnAutocomplete.Top = btnSend != null ? btnSend.Top : currentY + 6;
                btnAutocomplete.Visible = _pendingSettings == null || _pendingSettings.EnableAutocomplete;
            }

            currentY = (btnSend != null ? btnSend.Bottom : currentY + BTN_H) + 10;

            if (pnlBottomDivider != null)
            {
                pnlBottomDivider.Left = PAD;
                pnlBottomDivider.Top = currentY;
                pnlBottomDivider.Width = w;
            }
            currentY += 10;

            if (txtResponse != null)
            {
                lblResponse.Top = currentY;
                txtResponse.Top = currentY + LBL_H + 2;
                txtResponse.Width = w;
                txtResponse.Height = _pendingSettings != null ? Clamp(_pendingSettings.ResponseBoxHeight, 80, 420) : DEFAULT_RESPONSE_BOX_H;
            }
            currentY = (txtResponse != null ? txtResponse.Bottom : currentY + DEFAULT_RESPONSE_BOX_H) + 8;

            if (btnReject != null)
            {
                btnAccept.Top = currentY;
                btnReject.Top = currentY;
                btnInsert.Top = currentY;
                btnReject.Left = PAD + 96;
            }
            if (btnInsert != null)
                btnInsert.Left = PAD + 192;
            currentY += ACTION_ROW_H + 8;

            if (btnSettings != null)
            {
                btnSettings.Top = currentY;
            }
            if (lblModel != null)
            {
                lblModel.Top = currentY + 4;
                lblModel.Left = PAD + 102;
            }
            if (cboModel != null)
            {
                cboModel.Left = PAD + 152;
                cboModel.Top = currentY;
                cboModel.Width = Math.Max(120, w - 152);
            }

            if (this.AutoScroll)
            {
                this.AutoScrollMinSize = new Size(0, currentY + MODEL_ROW_H + bottomGap);
            }
        }

        private void UpdateModeUI()
        {
            if (btnSend != null)
                btnSend.Text = _isEditMode ? "直接修改" : "发送";

            if (btnAutocomplete != null)
            {
                var autocompleteEnabled = _pendingSettings == null || _pendingSettings.EnableAutocomplete;
                btnAutocomplete.Enabled = autocompleteEnabled;
                btnAutocomplete.Visible = autocompleteEnabled;
            }

            if (btnAccept != null)
                btnAccept.Visible = !_isEditMode;
            if (btnReject != null)
                btnReject.Visible = !_isEditMode;
        }

        // ======================== 事件处理 ========================
        private async void BtnSend_Click(object sender, EventArgs e)
        {
            RefreshFromWord();

            var sourceText = txtSelection.Text.Trim();
            if (string.IsNullOrEmpty(sourceText))
            {
                txtResponse.Text = "请先在 Word 中选中文字，再点击「抓取选区」。";
                return;
            }

            var promptText = string.IsNullOrWhiteSpace(txtPrompt?.Text) ? "请根据以下原文执行用户要求。" : txtPrompt.Text.Trim();
            var prompt = $"{promptText}\n\n原文：\n{sourceText}";

            CaptureSelectionRange();
            EnsureCurrentSession();
            var historyMessages = GetSessionHistoryMessages().ToList();

            SetBusy(true);
            txtResponse.Text = _isEditMode ? "思考中，完成后将直接替换选中文本..." : "思考中...";

            var reply = await _aiService.CallAsync(SystemPrompt, historyMessages, prompt);
            RunOnUiThread(() =>
            {
                _lastReply = reply;
                txtResponse.Text = reply;
                if (_isEditMode)
                {
                    ApplyReplyToSelection(reply);
                    txtResponse.Text = reply + Environment.NewLine + Environment.NewLine + "✓ 已直接替换原选区。";
                }
                AppendSessionMessage("assistant", reply);
                SetBusy(false);
                PersistHistoryState();
            });
        }

        private async void BtnAutocomplete_Click(object sender, EventArgs e)
        {
            if (_pendingSettings != null && !_pendingSettings.EnableAutocomplete)
            {
                txtResponse.Text = "自动补全已关闭，请先在设置里启用。";
                return;
            }

            RefreshFromWord();
            var sourceText = txtSelection.Text.Trim();
            if (string.IsNullOrEmpty(sourceText))
            {
                txtResponse.Text = "请先在 Word 中选中要补全的文本。";
                return;
            }

            var promptText = string.IsNullOrWhiteSpace(txtPrompt?.Text) ? "请补全以下文本，只输出补全后的完整文本。" : txtPrompt.Text.Trim();
            var prompt = $"{promptText}\n\n待补全文本：\n{sourceText}";

            CaptureSelectionRange();
            EnsureCurrentSession();
            var historyMessages = GetSessionHistoryMessages().ToList();

            SetBusy(true);
            txtResponse.Text = "自动补全中...";

            var reply = await _aiService.CallAsync(SystemPrompt, historyMessages, prompt, 0.2, _pendingSettings?.MaxTokens ?? 2048);
            RunOnUiThread(() =>
            {
                _lastReply = reply;
                txtResponse.Text = reply;
                ApplyReplyToSelection(reply);
                AppendSessionMessage("assistant", reply);
                SetBusy(false);
                PersistHistoryState();
            });
        }

        private void AcceptEdit()
        {
            if (string.IsNullOrEmpty(_lastReply)) return;
            ApplyReplyToSelection(_lastReply);
            txtResponse.Text = "✓ 已应用修改。";
        }

        private void InsertAtCursor()
        {
            if (string.IsNullOrEmpty(txtResponse.Text)) return;
            var word = Globals.ThisAddIn.Application;
            word.Selection.TypeText(txtResponse.Text);
        }

        private void RefreshFromWord()
        {
            var word = Globals.ThisAddIn.Application;
            var selText = word.Selection?.Text ?? "";
            if (word.Selection != null)
            {
                _selectionStart = word.Selection.Start;
                _selectionEnd = word.Selection.End;
            }
            RefreshSelection(selText);
        }

        private void CaptureSelectionRange()
        {
            var word = Globals.ThisAddIn.Application;
            if (word.Selection == null)
            {
                _selectionStart = -1;
                _selectionEnd = -1;
                return;
            }

            _selectionStart = word.Selection.Start;
            _selectionEnd = word.Selection.End;
        }

        private void ApplyReplyToSelection(string reply)
        {
            if (string.IsNullOrWhiteSpace(reply))
                return;

            var word = Globals.ThisAddIn.Application;
            if (_selectionStart >= 0 && _selectionEnd >= _selectionStart)
            {
                try
                {
                    var range = word.ActiveDocument.Range(_selectionStart, _selectionEnd);
                    range.Text = reply;
                    return;
                }
                catch
                {
                    // 回退到当前选区
                }
            }

            var selection = word.Selection;
            if (selection != null)
            {
                selection.TypeText(reply);
            }
        }

        private void SaveMode() { /* settings saving done via SettingsManager */ }
        private void SaveModel()
        {
            var settings = SettingsManager.Load();
            settings.Model = cboModel.Text;
            settings.Save();
            _aiService?.UpdateConfig(settings.ApiKey, settings.BaseUrl, settings.Model);
        }

        private void OpenSettings()
        {
            var settings = SettingsManager.Load();
            using (var dlg = new SettingsDialog(settings))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    dlg.Settings.Save();
                    _aiService?.UpdateConfig(dlg.Settings.ApiKey, dlg.Settings.BaseUrl, dlg.Settings.Model);
                    cboModel.Text = dlg.Settings.Model;
                    _pendingSettings = dlg.Settings;
                    ApplyConfiguredHeights();
                    if (btnAutocomplete != null)
                        btnAutocomplete.Visible = dlg.Settings.EnableAutocomplete;
                    LayoutUI();
                }
            }
        }

        private void LoadHistory()
        {
            _sessions = ConversationStore.Load();
            if (_sessions.Count == 0)
            {
                _sessions.Add(ConversationStore.CreateSession());
            }

            _sessions = _sessions.OrderByDescending(s => s.UpdatedAt).ToList();

            if (cboHistory == null)
                return;

            _isLoadingHistory = true;
            cboHistory.Items.Clear();
            foreach (var session in _sessions)
            {
                cboHistory.Items.Add(session);
            }
            _isLoadingHistory = false;
        }

        private void SelectSession(string sessionId)
        {
            if (string.IsNullOrWhiteSpace(sessionId))
                return;

            var session = _sessions.FirstOrDefault(s => s.Id == sessionId);
            if (session == null)
                return;

            _currentSession = session;
            if (cboHistory != null)
            {
                _isLoadingHistory = true;
                cboHistory.SelectedItem = session;
                _isLoadingHistory = false;
            }

            LoadSessionIntoUi(session);
            if (_pendingSettings != null)
            {
                _pendingSettings.LastSessionId = session.Id;
                SettingsManager.Save(_pendingSettings);
            }
        }

        private void SelectCurrentHistoryFromUi()
        {
            if (cboHistory?.SelectedItem is ChatSessionData session)
            {
                _currentSession = session;
                LoadSessionIntoUi(session);
                PersistHistoryState();
            }
        }

        private void LoadSessionIntoUi(ChatSessionData session)
        {
            if (session == null)
                return;

            var lastUser = session.Messages.LastOrDefault(m => string.Equals(m.Role, "user", StringComparison.OrdinalIgnoreCase));
            var lastAssistant = session.Messages.LastOrDefault(m => string.Equals(m.Role, "assistant", StringComparison.OrdinalIgnoreCase));

            if (!string.IsNullOrWhiteSpace(lastUser?.Content) && txtPrompt != null)
            {
                txtPrompt.Text = lastUser.Content;
            }

            if (!string.IsNullOrWhiteSpace(lastAssistant?.Content) && txtResponse != null)
            {
                txtResponse.Text = lastAssistant.Content;
                _lastReply = lastAssistant.Content;
            }
        }

        private void StartNewSession()
        {
            var session = ConversationStore.CreateSession();
            _sessions.Insert(0, session);
            if (cboHistory != null)
            {
                _isLoadingHistory = true;
                cboHistory.Items.Insert(0, session);
                cboHistory.SelectedItem = session;
                _isLoadingHistory = false;
            }
            _currentSession = session;
            txtPrompt.Text = "请根据下面的原文和当前会话历史继续处理内容。";
            txtResponse.Text = "";
            _lastReply = "";
            PersistHistoryState();
        }

        private void DeleteCurrentSession()
        {
            if (_currentSession == null)
                return;

            var currentId = _currentSession.Id;
            _sessions.RemoveAll(s => s.Id == currentId);
            ConversationStore.Save(_sessions);
            LoadHistory();
            if (_sessions.Count == 0)
            {
                StartNewSession();
                return;
            }

            SelectSession(_sessions[0].Id);
        }

        private void EnsureCurrentSession()
        {
            if (_currentSession != null)
                return;

            if (_sessions.Count == 0)
            {
                _sessions.Add(ConversationStore.CreateSession());
            }

            _currentSession = _sessions[0];
            if (cboHistory != null)
            {
                _isLoadingHistory = true;
                cboHistory.SelectedItem = _currentSession;
                _isLoadingHistory = false;
            }
        }

        private void AppendSessionMessage(string role, string content)
        {
            if (_currentSession == null || string.IsNullOrWhiteSpace(content))
                return;

            _currentSession.Messages.Add(new ChatMessageData { Role = role, Content = content, Timestamp = DateTime.Now });
            _currentSession.UpdatedAt = DateTime.Now;

            if (_currentSession.Messages.Count == 1)
            {
                _currentSession.Title = MakeSessionTitle(content);
            }

            RefreshHistoryUi();
        }

        private IEnumerable<ChatMessageData> GetSessionHistoryMessages()
        {
            if (_currentSession == null)
                yield break;

            foreach (var message in _currentSession.Messages)
            {
                yield return message;
            }
        }

        private void PersistHistoryState()
        {
            ConversationStore.Save(_sessions);
            if (_pendingSettings != null && _currentSession != null)
            {
                _pendingSettings.LastSessionId = _currentSession.Id;
                SettingsManager.Save(_pendingSettings);
            }
            RefreshHistoryUi();
        }

        private void RefreshHistoryUi()
        {
            if (cboHistory == null)
                return;

            var selectedId = _currentSession?.Id;
            _isLoadingHistory = true;
            cboHistory.Items.Clear();
            foreach (var session in _sessions.OrderByDescending(s => s.UpdatedAt))
            {
                cboHistory.Items.Add(session);
            }
            if (!string.IsNullOrWhiteSpace(selectedId))
            {
                var selected = _sessions.FirstOrDefault(s => s.Id == selectedId);
                if (selected != null)
                {
                    cboHistory.SelectedItem = selected;
                }
            }
            _isLoadingHistory = false;
        }

        private string MakeSessionTitle(string content)
        {
            var text = (content ?? string.Empty).Trim();
            if (text.Length <= 18)
                return text;

            return text.Substring(0, 18) + "...";
        }

        private int Clamp(int value, int min, int max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        // ======================== 控件工厂 ========================
        private Label AddLabel(string text, int left, int top, int width, int height, bool bold)
        {
            var lbl = new Label
            {
                Text = text,
                Left = left,
                Top = top,
                Width = width,
                Height = height,
                ForeColor = Color.FromArgb(35, 35, 35),
                AutoEllipsis = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            lbl.Font = bold ? new Font(this.Font, FontStyle.Bold) : this.Font;
            this.Controls.Add(lbl);
            return lbl;
        }

        private TextBox AddTextBox(int left, int top, int width, int height, bool multiLine, bool scrollBars, bool readOnly)
        {
            var tb = new TextBox
            {
                Left = left, Top = top, Width = width, Height = height,
                Multiline = multiLine, ReadOnly = readOnly,
                ScrollBars = scrollBars ? ScrollBars.Vertical : ScrollBars.None,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                ForeColor = Color.FromArgb(25, 25, 25),
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                WordWrap = false
            };
            this.Controls.Add(tb);
            return tb;
        }

        private Button AddButton(string text, int left, int top, int width, int height)
        {
            var btn = new Button
            {
                Text = text, Left = left, Top = top,
                Width = width, Height = height,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(248, 248, 248),
                ForeColor = Color.FromArgb(35, 35, 35),
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point)
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(210, 210, 210);
            this.Controls.Add(btn);
            return btn;
        }

        private Panel AddDivider(int left, int top, int width)
        {
            var panel = new Panel
            {
                Left = left,
                Top = top,
                Width = width,
                Height = 1,
                BackColor = Color.FromArgb(220, 220, 220),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(panel);
            return panel;
        }

        private RadioButton AddRadioButton(string text, int left, int top, int width, int height, bool @checked)
        {
            var rb = new RadioButton
            {
                Text = text,
                Left = left,
                Top = top,
                Width = width,
                Height = height,
                Checked = @checked,
                BackColor = Color.Transparent,
                ForeColor = Color.FromArgb(35, 35, 35),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei UI", 9.5F, FontStyle.Regular, GraphicsUnit.Point)
            };
            this.Controls.Add(rb);
            return rb;
        }
    }
}
