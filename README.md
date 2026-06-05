# WAI — Word AI 侧边栏插件

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![.NET Framework](https://img.shields.io/badge/.NET-4.8-blue.svg)](https://dotnet.microsoft.com/download/dotnet-framework/net48)
[![Platform](https://img.shields.io/badge/Platform-Windows%20Word-green.svg)]()

> 一个开源的 Word 右侧停靠 AI 侧边栏，支持润色、改写、纠错、翻译、总结、续写、自动补全、多轮对话。

---

## 📸 截图

（TODO：添加截图）

---

## ✨ 功能

| 功能 | 说明 |
|------|------|
| ✅ 真正停靠侧栏 | 使用 VSTO CustomTaskPane，贴合 Word 右侧 |
| ✅ 选区自动跟随 | 选中文字自动显示在侧栏中 |
| ✅ 编辑模式 | AI 返回后直接替换选中的原文 |
| ✅ 问答模式 | 手动确认后插入（可接受/放弃/插入光标） |
| ✅ 多轮对话历史 | 会话持久化存储，可切换历史会话 |
| ✅ 自定义指令 | 自由输入 Prompt，支持润色/翻译/总结等预设 |
| ✅ 自动补全 | 可选开关，选中不完整文本后自动补全 |
| ✅ 可调节界面 | 选区/Prompt/输出三个区域高度可在设置中调节 |
| ✅ 模型切换 | 支持 DeepSeek V4 Pro / Flash 切换 |
| ✅ 设置持久化 | API Key、模型、温度等配置保存到本地文件 |

---

## 📦 下载与安装

### 方式一：一键安装（推荐）

1. 前往 [Releases 页面](https://github.com/YOUR_USERNAME/WaiAddIn/releases) 下载最新版本的 `WAI-Installer.zip`
2. 解压到任意文件夹
3. **右键** `Install.bat` → **以管理员身份运行**
4. 等待安装完成
5. 重启 Word，右侧即可看到 **WAI** 面板

### 方式二：从源码编译

#### 环境要求

- **Windows 10 / 11**
- **Visual Studio 2022+**（Community 免费版即可）
  - 安装时勾选 **「Office/SharePoint 开发」** 工作负载
- **Microsoft Word 2019 / 2021 / 365**（桌面版）
- **.NET Framework 4.8**

#### 编译步骤

```bash
# 1. 克隆仓库
git clone https://github.com/YOUR_USERNAME/WaiAddIn.git
cd WaiAddIn

# 2. 用 Visual Studio 打开 WaiAddIn.sln（或直接打开 WaiAddIn/WaiAddIn.csproj）

# 3. 还原 NuGet 包
#    在 VS 中：工具 → NuGet 包管理器 → 还原

# 4. 按 F5 编译运行，Word 会自动打开
```

#### 发布安装包

在 Visual Studio 中：
1. 右键项目 → **发布**
2. 选择发布目标为 **文件夹**
3. 将发布出的文件（.vsto + .dll + setup.exe）复制到 `publish/` 目录
4. 和 `Install.bat`、`Uninstall.bat` 一同打包分发

---

## 🔧 初次使用配置

1. 打开 Word，右侧可见 **WAI** 面板
2. 点击面板底部的 **「设置」** 按钮
3. 填入你的 API Key（目前支持 DeepSeek API 兼容接口）
4. 可选：修改 API 地址、模型名、温度等参数
5. 点击 **保存**
6. 选中 Word 中的文字，点击 **发送** 或 **自动补全** 即可

> 💡 **获取 API Key**：前往 [DeepSeek 官网](https://platform.deepseek.com/) 注册并创建 API Key。

---

## 🗑️ 卸载

### 使用卸载脚本

1. 找到 `Uninstall.bat`
2. **右键** → **以管理员身份运行**
3. 重启 Word

### 手动卸载

1. 打开 Word → **文件** → **选项** → **加载项**
2. 在「管理」下拉框选择 **COM 加载项** → 点击 **转到**
3. 找到 **WAI** → 取消勾选或移除
4. 删除本地设置文件夹：`%LOCALAPPDATA%\WAI`
5. 删除 VSTO 缓存：`%LOCALAPPDATA%\Microsoft\VSTO\Cache`

---

## 📁 项目结构

```
WaiAddIn/
├── ThisAddIn.cs          # VSTO 入口，创建停靠面板、监听选区
├── SidebarControl.cs     # 侧边栏 UI 控件（WinForms）
├── AIService.cs          # DeepSeek API 调用封装
├── SettingsManager.cs    # 设置存储（JSON 文件）
├── SettingsDialog.cs     # 设置配置窗口
├── ConversationStore.cs  # 会话历史持久化存储
├── Properties/
│   └── AssemblyInfo.cs   # 程序集信息
└── WaiAddIn.csproj       # 项目文件 (.NET Framework 4.8)

publish/
├── Install.bat            # 一键安装脚本
└── Uninstall.bat          # 一键卸载脚本
```

---

## ⚙️ 配置文件位置

| 内容 | 路径 |
|------|------|
| 设置文件 | `%LOCALAPPDATA%\WAI\settings.json` |
| 会话历史 | `%LOCALAPPDATA%\WAI\sessions.json` |

删除这两个文件即可恢复出厂设置。

---

## 🚨 常见问题

### Q：安装时提示"无法验证发布者"

这是 ClickOnce 自签名证书的默认行为，不影响使用。可以：
- 点击 **仍然安装**
- 或者先运行 `Install.bat`（会自动处理）

### Q：Word 右侧没有显示 WAI 面板

1. 确认 Word 是桌面版，不是网页版
2. 检查 Word → 选项 → 加载项 → COM 加载项中 WAI 是否已启用
3. 重启 Word
4. 如果仍不显示，重新运行安装程序

### Q：提示"未配置 API Key"

打开 WAI 面板底部的 **设置** 按钮，填入有效的 API Key。

### Q：插件不工作了/报错了

1. 打开设置，确认 API Key 和 API 地址正确
2. 删除 `%LOCALAPPDATA%\WAI\` 下的设置文件重置配置
3. 重新安装插件

---

## 🛠️ 技术栈

- **语言**: C# (.NET Framework 4.8)
- **框架**: VSTO (Visual Studio Tools for Office)
- **UI**: Windows Forms (UserControl)
- **API**: DeepSeek Chat Completions API
- **存储**: Newtonsoft.Json (JSON 文件)

---

## 📄 许可证

本项目采用 **MIT 许可证** — 详见 [LICENSE](LICENSE) 文件。

---

## 🙏 贡献

欢迎提交 Issue 和 Pull Request！

1. Fork 本仓库
2. 创建功能分支 (`git checkout -b feature/amazing-feature`)
3. 提交修改 (`git commit -m 'Add amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 创建 Pull Request

---

## ⚠️ 重要提醒

- 本项目**不内置任何 API Key**，请自行申请
- API Key 请妥善保管，不要提交到代码仓库
- 本插件仅支持 **Windows 桌面版 Word**，不支持 Mac / 网页版
