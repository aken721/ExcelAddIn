# HtmlEditorControl 控件使用指南

## 已完成的集成工作

本项目已经将 `HtmlEditorControl` 控件及其依赖文件集成到项目中：

### 已复制的文件
- ✅ `HtmlEditorControl.dll` - 主控件库
- ✅ `HtmlEditorControl.xml` - XML 文档文件
- ✅ `Microsoft.Web.WebView2.Core.dll` - WebView2 核心库
- ✅ `Microsoft.Web.WebView2.WinForms.dll` - WebView2 WinForms 支持库
- ✅ `runtimes/` 文件夹 - 包含不同平台的本地依赖（x86、x64、ARM64）

### 已添加的项目引用
已在 `ExcelAddIn.csproj` 文件中添加了以下引用：
- HtmlEditorControl
- Microsoft.Web.WebView2.Core
- Microsoft.Web.WebView2.WinForms

## 在 Visual Studio 工具箱中添加控件

### 方法 1：自动添加（推荐）
1. 在 Visual Studio 中重新加载项目（如果项目已打开）
2. 打开任意 Form 设计器（例如 Form1.cs）
3. 控件应该会自动出现在工具箱中

### 方法 2：手动添加到工具箱
如果控件没有自动出现，请按照以下步骤手动添加：

1. **打开工具箱**
   - 在 Visual Studio 中按 `Ctrl + Alt + X` 或从菜单选择 "视图" → "工具箱"

2. **创建新的工具箱选项卡**（可选）
   - 右键点击工具箱空白处
   - 选择 "添加选项卡"
   - 命名为 "HtmlEditor" 或其他您喜欢的名称

3. **添加控件**
   - 右键点击工具箱中的选项卡
   - 选择 "选择项..."
   - 在弹出的对话框中，点击 ".NET Framework 组件" 选项卡
   - 点击 "浏览..." 按钮
   - 导航到项目根目录（`E:\SourceCode\Csharp\ExcelAddIn`）
   - 选择 `HtmlEditorControl.dll` 文件
   - 点击 "打开"
   - 确保 `HtmlEditorControl` 已被勾选
   - 点击 "确定"

4. **验证控件已添加**
   - 在工具箱中应该能看到 `HtmlEditorControl` 控件
   - 控件图标应该是一个可用的组件

## 使用控件

### 通过拖拽使用
1. 打开或创建一个 Windows Form（例如 Form1.cs、Form2.cs 等）
2. 从工具箱中找到 `HtmlEditorControl` 控件
3. 将控件拖拽到 Form 设计器表面
4. 调整控件的大小和位置
5. 在属性窗口中设置控件属性

### 通过代码使用
```csharp
using HtmlEditorControl;

public partial class Form1 : Form
{
    private HtmlEditorControl.HtmlEditorControl htmlEditor;
    
    public Form1()
    {
        InitializeComponent();
        InitializeHtmlEditor();
    }
    
    private void InitializeHtmlEditor()
    {
        htmlEditor = new HtmlEditorControl.HtmlEditorControl
        {
            Dock = DockStyle.Fill
        };
        this.Controls.Add(htmlEditor);
        
        // 设置 HTML 内容
        htmlEditor.HtmlContent = "<h1>Hello, World!</h1>";
    }
}
```

## 常见问题

### Q1: 控件无法在工具箱中显示？
**解决方法：**
- 确保项目已成功编译
- 尝试关闭并重新打开 Visual Studio
- 检查 .csproj 文件中的引用是否正确
- 尝试手动添加控件到工具箱（参见上面的方法 2）

### Q2: 运行时提示找不到 WebView2 相关 DLL？
**解决方法：**
- 确保 `runtimes` 文件夹已复制到输出目录
- 检查项目配置是否正确设置了 `CopyToOutputDirectory`
- 重新生成解决方案

### Q3: 设计器显示错误或无法加载控件？
**解决方法：**
- 确保所有依赖的 DLL 都在项目根目录下
- 尝试清理并重新生成项目
- 检查 .NET Framework 版本是否兼容（本项目使用 .NET Framework 4.8）

### Q4: 需要安装 WebView2 运行时吗？
**回答：**
- 是的，最终用户需要安装 Microsoft Edge WebView2 运行时
- 下载地址：https://developer.microsoft.com/zh-cn/microsoft-edge/webview2/
- 或者使用常青版（Evergreen），它会自动更新

## 控件主要功能

`HtmlEditorControl` 提供以下主要功能：
- 富文本 HTML 编辑
- 支持 HTML 内容的显示和编辑
- 基于 Microsoft Edge WebView2 技术
- 支持现代 Web 标准
- 高性能渲染

## 相关文档

更多详细的属性和方法说明，请参考：
- `HtmlEditorControl.xml` - API 文档
- `E:\SourceCode\Csharp\HtmlEditorControl\HtmlEditorControl属性设置指南.md` - 原项目的属性设置指南
- `E:\SourceCode\Csharp\HtmlEditorControl\Visual Studio安装使用指南.md` - Visual Studio 安装使用指南

## 技术支持

如果遇到其他问题，请参考原项目文档或检查项目的 GitHub 仓库。

