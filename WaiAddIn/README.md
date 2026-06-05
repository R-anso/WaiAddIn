# WAI — Word AI 侧边栏（C# VSTO 版本）

真正的 Word 停靠侧边栏，使用 C# VSTO (Visual Studio Tools for Office) 开发。

## 环境要求

- **Visual Studio 2022**（Community 版免费）
  - 安装时勾选「Office/SharePoint 开发」工作负载
- **Microsoft Word 2016/2019/2021/365**（桌面版，不支持网页版）
- **.NET Framework 4.8**

## 快速开始

### 1. 在 Visual Studio 中创建项目

1. 打开 Visual Studio
2. 新建项目 → 搜索「**Word VSTO 外接程序**」
3. 项目名称：`WaiAddIn`
4. 位置：本项目目录下的 `WaiAddIn` 文件夹
5. 框架：**.NET Framework 4.8**
6. 点击「创建」

### 2. 安装 NuGet 包

在「工具 → NuGet 包管理器 → 管理解决方案的 NuGet 包」中安装：

```
Newtonsoft.Json   (版本 >= 13.0.3)
```

### 3. 替换代码文件

将本项目中的以下文件**覆盖**到 VS 项目中对应位置：

| 文件 | 说明 |
|------|------|
| `ThisAddIn.cs` | 插件入口，创建停靠面板、监听选区 |
| `SidebarControl.cs` | 侧边栏 UI 控件 |
| `AIService.cs` | DeepSeek API 调用 |
| `SettingsManager.cs` | 设置存储 |
| `SettingsDialog.cs` | 设置窗口 |

### 4. 编译和运行

- 按 `F5` 启动调试（会自动打开 Word）
- 首次运行会在 Word 右侧看到 **WAI** 停靠面板
- 选中文字后，面板自动显示选中内容

### 5. 发布安装

- 生成 → 发布（选择 ClickOnce 或 Windows Installer）
- 或直接在 `bin/Debug` 中找到 `.vsto` 文件双击安装

## 与 VBA 版的区别

| 功能 | VBA 版 | C# VSTO 版 |
|------|--------|-----------|
| 侧边栏 | 浮动 UserForm（假侧边栏） | **真正停靠面板**（Custom Task Pane） |
| 选区跟随 | 手动点击刷新 | **自动实时跟随** |
| 快捷键 | 需手动绑定，易冲突 | 不需要，面板一直可见 |
| 分发 | 分享 .docm 即可 | 需要安装 .vsto |

## 项目结构

```
WaiAddIn/
├── ThisAddIn.cs          # VSTO 入口，创建停靠面板
├── SidebarControl.cs     # 侧边栏 WinForms 控件
├── AIService.cs          # DeepSeek API 封装
├── SettingsManager.cs    # 设置读/写（JSON 文件）
├── SettingsDialog.cs     # 设置窗口
├── Properties/
│   └── AssemblyInfo.cs
└── WaiAddIn.csproj
```
