using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;

namespace WaiAddIn
{
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane _taskPane;
        private SidebarControl _sidebarControl;
        private AIService _aiService;

        private void InternalStartup()
        {
            this.Startup += ThisAddIn_Startup;
            this.Shutdown += ThisAddIn_Shutdown;
        }

        // ======================== 初始化 ========================
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                var settings = SettingsManager.Load();
                _aiService = new AIService(settings.ApiKey, settings.BaseUrl, settings.Model);

                // 创建侧边栏控件
                _sidebarControl = new SidebarControl();
                _sidebarControl.SetAIService(_aiService);
                _sidebarControl.LoadSettings(settings);

                // 创建真正的停靠任务窗格
                _taskPane = this.CustomTaskPanes.Add(_sidebarControl, "WAI");
                _taskPane.Width = 380;
                _taskPane.Visible = true;

                // 监听选区变化
                this.Application.WindowSelectionChange += Application_WindowSelectionChange;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"WAI 启动失败：{ex.Message}", "WAI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ======================== 选区变化 ========================
        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            try
            {
                if (_sidebarControl == null) return;

                var selText = Sel?.Text ?? "";
                if (!string.IsNullOrWhiteSpace(selText))
                {
                    _sidebarControl.RefreshSelection(selText);
                }
            }
            catch { }
        }

        // ======================== 清理 ========================
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (_taskPane != null)
                {
                    _taskPane.Visible = false;
                    _taskPane = null;
                }
                this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
            }
            catch { }
        }

        // ======================== 注意 ========================
        // InternalStartup() 供 VSTO 生成代码调用，用于绑定 Startup / Shutdown。
    }
}
