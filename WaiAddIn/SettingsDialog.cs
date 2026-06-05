using System;
using System.Drawing;
using System.Windows.Forms;

namespace WaiAddIn
{
    public class SettingsDialog : Form
    {
        public SettingsData Settings { get; private set; }

        private TextBox txtApiKey;
        private TextBox txtBaseUrl;
        private TextBox txtModel;
        private NumericUpDown numMaxTokens;
        private NumericUpDown numTemperature;
        private CheckBox chkEditMode;
        private NumericUpDown numSelectionHeight;
        private NumericUpDown numPromptHeight;
        private NumericUpDown numResponseHeight;
        private CheckBox chkAutocomplete;

        public SettingsDialog(SettingsData settings)
        {
            Settings = settings;
            Text = "WAI 设置";
            Width = 420;
            Height = 520;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            BuildUI();
            LoadSettings();
        }

        private void BuildUI()
        {
            int y = 12, pad = 12, lblW = 128, ctrlW = 230;

            AddLabel("API Key：", pad, y, lblW); y += 22;
            txtApiKey = AddTextBox(pad, y, ctrlW); y += 28;
            AddLabel("API 地址：", pad, y, lblW); y += 22;
            txtBaseUrl = AddTextBox(pad, y, ctrlW); y += 28;
            AddLabel("模型名：", pad, y, lblW); y += 22;
            txtModel = AddTextBox(pad, y, ctrlW); y += 30;

            AddLabel("最大 Token：", pad, y, lblW);
            numMaxTokens = new NumericUpDown { Left = pad + lblW + 10, Top = y, Width = 80, Minimum = 256, Maximum = 8192, Value = 2048 };
            Controls.Add(numMaxTokens);
            y += 28;

            AddLabel("Temperature：", pad, y, lblW);
            numTemperature = new NumericUpDown { Left = pad + lblW + 10, Top = y, Width = 80, Minimum = 0, Maximum = 2, DecimalPlaces = 2, Increment = 0.1m, Value = 0.2m };
            Controls.Add(numTemperature);
            y += 30;

            chkEditMode = new CheckBox { Text = "默认使用编辑模式（可直接修改原文）", Left = pad, Top = y, Width = 360, Checked = true };
            Controls.Add(chkEditMode);
            y += 28;

            AddLabel("选区框高度：", pad, y, lblW);
            numSelectionHeight = new NumericUpDown { Left = pad + lblW + 10, Top = y, Width = 80, Minimum = 48, Maximum = 260, Increment = 4, Value = 96 };
            Controls.Add(numSelectionHeight);
            y += 28;

            AddLabel("Prompt框高度：", pad, y, lblW);
            numPromptHeight = new NumericUpDown { Left = pad + lblW + 10, Top = y, Width = 80, Minimum = 60, Maximum = 360, Increment = 4, Value = 140 };
            Controls.Add(numPromptHeight);
            y += 28;

            AddLabel("输出框高度：", pad, y, lblW);
            numResponseHeight = new NumericUpDown { Left = pad + lblW + 10, Top = y, Width = 80, Minimum = 80, Maximum = 420, Increment = 4, Value = 210 };
            Controls.Add(numResponseHeight);
            y += 30;

            chkAutocomplete = new CheckBox { Text = "启用自动补全", Left = pad, Top = y, Width = 180, Checked = false };
            Controls.Add(chkAutocomplete);
            y += 32;

            var btnOK = new Button { Text = "保存", Left = 200, Top = y, Width = 80, Height = 28, DialogResult = DialogResult.OK };
            btnOK.Click += (s, e) => SaveSettings();
            Controls.Add(btnOK);

            var btnCancel = new Button { Text = "取消", Left = 290, Top = y, Width = 80, Height = 28, DialogResult = DialogResult.Cancel };
            Controls.Add(btnCancel);
        }

        private void LoadSettings()
        {
            txtApiKey.Text = Settings.ApiKey;
            txtBaseUrl.Text = Settings.BaseUrl;
            txtModel.Text = Settings.Model;
            numMaxTokens.Value = Settings.MaxTokens;
            numTemperature.Value = (decimal)Settings.Temperature;
            chkEditMode.Checked = Settings.EditModeDefault;
            numSelectionHeight.Value = Math.Max(numSelectionHeight.Minimum, Math.Min(numSelectionHeight.Maximum, Settings.SelectionBoxHeight));
            numPromptHeight.Value = Math.Max(numPromptHeight.Minimum, Math.Min(numPromptHeight.Maximum, Settings.PromptBoxHeight));
            numResponseHeight.Value = Math.Max(numResponseHeight.Minimum, Math.Min(numResponseHeight.Maximum, Settings.ResponseBoxHeight));
            chkAutocomplete.Checked = Settings.EnableAutocomplete;
        }

        private void SaveSettings()
        {
            Settings.ApiKey = txtApiKey.Text.Trim();
            Settings.BaseUrl = txtBaseUrl.Text.Trim();
            Settings.Model = txtModel.Text.Trim();
            Settings.MaxTokens = (int)numMaxTokens.Value;
            Settings.Temperature = (double)numTemperature.Value;
            Settings.EditModeDefault = chkEditMode.Checked;
            Settings.SelectionBoxHeight = (int)numSelectionHeight.Value;
            Settings.PromptBoxHeight = (int)numPromptHeight.Value;
            Settings.ResponseBoxHeight = (int)numResponseHeight.Value;
            Settings.EnableAutocomplete = chkAutocomplete.Checked;
        }

        private Label AddLabel(string text, int left, int top, int width)
        {
            var lbl = new Label { Text = text, Left = left, Top = top, Width = width, Height = 20 };
            Controls.Add(lbl);
            return lbl;
        }

        private TextBox AddTextBox(int left, int top, int width)
        {
            var tb = new TextBox { Left = left, Top = top, Width = width, BorderStyle = BorderStyle.FixedSingle };
            Controls.Add(tb);
            return tb;
        }
    }
}
