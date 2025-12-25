# DynamicPic.Core 用户使用指南

## 简介

DynamicPic.Core 是一个 .NET 类库，用于从 Excel 数据生成动态图表，支持 GIF 动图和交互式 HTML 输出。

**版本：** 1.0.2  
**支持框架：** .NET Framework 4.8 / .NET 8.0 Windows

## 主要功能

- 从 Excel 文件读取数据（支持 .xlsx 和 .xls）
- 生成柱状图（横向/纵向）
- 生成 GIF 动图展示数据变化
- 创建交互式 HTML 图表（支持切片器、自动播放、视频导出）
- 自定义图表样式和颜色
- 灵活的数据排序选项
- 支持 Excel 插件开发（VSTO/COM 对象，仅 .NET Framework 4.8）

## 安装

### 方式 1：NuGet 包安装（推荐）

**使用本地 NuGet 源：**

1. 在项目目录创建 `nuget.config` 文件：
```xml
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
    <add key="LocalPackages" value="C:\path\to\nupkg\folder" />
  </packageSources>
</configuration>
```

2. 安装包：
```bash
dotnet add package DynamicPic.Core --version 1.0.2
```

### 方式 2：直接引用 DLL

在项目文件中添加：
```xml
<ItemGroup>
  <Reference Include="DynamicPic.Core">
    <HintPath>path\to\DynamicPic.Core.dll</HintPath>
  </Reference>
  <PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
  <PackageReference Include="Magick.NET-Q16-AnyCPU" Version="14.9.1" />
</ItemGroup>
```

---

## 快速开始

```csharp
using DynamicPic.Core;

// 1. 读取 Excel 数据
var data = ExcelReader.ReadValueData("data.xlsx");

// 2. 排序数据
var sortedData = ExcelReader.GetSortedValueData(data, SortOrder.Descending);

// 3. 配置图表
var config = new ChartConfig
{
    TitleTemplate = "{0}年数据",
    ChartType = ChartType.HorizontalBar
};

// 4. 生成 HTML
HtmlGenerator.GenerateInteractiveHtml(sortedData, "output.html", config);
```

---

## Excel 数据格式

Excel 文件应按以下格式组织数据：

| 项目名称 | 2021年 | 2022年 | 2023年 |
|---------|--------|--------|--------|
| 产品A    | 100    | 120    | 150    |
| 产品B    | 80     | 90     | 110    |
| 产品C    | 60     | 70     | 85     |

- 第一行：分组标题（年份/时间段）
- 第一列：类别名称
- 数据区域：数值

---

## 常用配置

### ChartConfig 主要属性

| 属性 | 说明 | 默认值 |
|------|------|--------|
| `TitleTemplate` | 标题模板，{0} 替换为分组名 | "" |
| `XAxisTitle` | X 轴标题 | "" |
| `YAxisTitle` | Y 轴标题 | "" |
| `ChartType` | 图表类型 | HorizontalBar |
| `Width` | 图表宽度 | 1200 |
| `Height` | 图表高度 | 800 |
| `DataChartColor` | 柱状图颜色 | "rgba(79, 129, 189, 0.8)" |
| `ShowDataLabels` | 显示数据标签 | true |
| `GifFrameDelay` | GIF 帧延迟（毫秒） | 1500 |
| `HtmlAutoPlay` | HTML 自动播放 | true |

### 配置示例

```csharp
var config = new ChartConfig
{
    TitleTemplate = "{0}年销售数据",
    XAxisTitle = "销售额（万元）",
    YAxisTitle = "产品",
    ChartType = ChartType.HorizontalBar,
    Width = 1200,
    Height = 800,
    DataChartColor = "rgba(54, 162, 235, 0.8)",
    ShowDataLabels = true,
    DataLabelFormat = "N1",
    GifFrameDelay = 1500,
    GifFirstLastDelay = 3000,
    HtmlAutoPlay = true,
    HtmlPlayInterval = 2000
};
```

---

## 使用示例

### 示例 1：生成交互式 HTML

```csharp
using DynamicPic.Core;

var data = ExcelReader.ReadValueData("sales.xlsx");
var sortedData = ExcelReader.GetSortedValueData(data, SortOrder.Descending);

var config = new ChartConfig
{
    TitleTemplate = "{0}年销售数据",
    XAxisTitle = "销售额（万元）",
    YAxisTitle = "产品",
    ChartType = ChartType.HorizontalBar,
    HtmlAutoPlay = true,
    HtmlPlayInterval = 2000
};

HtmlGenerator.GenerateInteractiveHtml(sortedData, "sales_chart.html", config);
```

### 示例 2：生成 GIF 动图

```csharp
using DynamicPic.Core;
using System.Collections.Generic;
using System.IO;
using System.Linq;

var data = ExcelReader.ReadValueData("sales.xlsx");
var sortedData = ExcelReader.GetSortedValueData(data, SortOrder.Descending);

var config = new ChartConfig
{
    TitleTemplate = "{0}年销售数据",
    ChartType = ChartType.HorizontalBar,
    DataChartColor = "rgba(54, 162, 235, 0.8)"
};

// 生成图表图片
string tempDir = "temp";
Directory.CreateDirectory(tempDir);
var chartPaths = new List<string>();

foreach (var group in sortedData.OrderBy(g => g.Key))
{
    string title = string.Format(config.TitleTemplate, group.Key);
    var chartData = group.Value.ToDictionary(x => x.Key, x => x.Value);
    string outputPath = Path.Combine(tempDir, $"chart_{group.Key}.png");
    
    ChartGenerator.GenerateBarChart(title, chartData, outputPath, config);
    chartPaths.Add(outputPath);
}

// 生成 GIF
GifCreator.CreateGifWithPause(chartPaths, "sales.gif", 1500, 3000);

// 清理临时文件
GifCreator.CleanupTempImages(chartPaths);
Directory.Delete(tempDir);
```

### 示例 3：Excel 插件中使用（仅 .NET Framework 4.8）

```csharp
using DynamicPic.Core;
using Excel = Microsoft.Office.Interop.Excel;

public class MyExcelAddin
{
    private Excel.Application _excelApp;

    // 从活动工作表生成图表
    public void GenerateChart()
    {
        // 通过表名读取（传 null 使用活动工作表）
        var data = ExcelReader.ReadValueDataBySheetName(_excelApp, null);
        var sortedData = ExcelReader.GetSortedValueData(data, SortOrder.Descending);

        var config = new ChartConfig
        {
            TitleTemplate = "{0}年数据",
            ChartType = ChartType.HorizontalBar
        };

        string outputPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop), 
            "chart.html");
        HtmlGenerator.GenerateInteractiveHtml(sortedData, outputPath, config);
    }

    // 从选定区域生成图表
    public void GenerateChartFromSelection()
    {
        Excel.Range selection = _excelApp.Selection as Excel.Range;
        if (selection != null)
        {
            var data = ExcelReader.ReadValueDataFromRange(selection);
            // ... 后续处理
        }
    }
}
```

---

## API 快速参考

### ExcelReader - 数据读取

```csharp
// 从文件读取
var data = ExcelReader.ReadValueData("file.xlsx", DataOrientation.ByColumn);

// 从指定工作表读取
var data = ExcelReader.ReadValueDataFromWorksheet("file.xlsx", "Sheet1");

// 获取工作表列表
var sheets = ExcelReader.GetWorksheetNames("file.xlsx");

// 排序数据
var sortedData = ExcelReader.GetSortedValueData(data, SortOrder.Descending);

// Excel 插件专用（仅 .NET Framework 4.8）
var data = ExcelReader.ReadValueDataBySheetName(excelApp, "Sheet1");
var data = ExcelReader.ReadValueDataFromActiveSheet(excelApp);
var data = ExcelReader.ReadValueDataFromRange(range);
```

### ChartGenerator - 图表生成

```csharp
// 生成柱状图
ChartGenerator.GenerateBarChart(title, data, "output.png", config);
```

### GifCreator - GIF 生成

```csharp
// 创建 GIF
GifCreator.CreateGif(imagePaths, "output.gif", delayMs: 1500);

// 创建带暂停效果的 GIF（推荐）
GifCreator.CreateGifWithPause(imagePaths, "output.gif", 1500, 3000);

// 清理临时文件
GifCreator.CleanupTempImages(imagePaths);
```

### HtmlGenerator - HTML 生成

```csharp
// 生成交互式 HTML
HtmlGenerator.GenerateInteractiveHtml(sortedData, "output.html", config);
```

---

## 枚举类型

```csharp
// 数据方向
DataOrientation.ByColumn  // 按列分组
DataOrientation.ByRow     // 按行分组

// 图表类型
ChartType.HorizontalBar   // 横向柱状图
ChartType.VerticalBar     // 纵向柱状图

// 排序方向
SortOrder.Descending      // 降序
SortOrder.Ascending       // 升序

// 输出格式
OutputFormat.Gif          // GIF 动图
OutputFormat.Html         // HTML
OutputFormat.Both         // 同时输出
```

---

## 常见问题

### 1. 如何自定义颜色？

```csharp
config.DataChartColor = "rgba(255, 99, 132, 0.8)"; // RGBA
config.DataChartColor = "#FF6384";                  // HEX
config.DataChartColor = "rgb(255, 99, 132)";        // RGB
```

### 2. 如何调整数据标签位置？

```csharp
// 横向柱状图
config.DataLabelAnchor = "end";
config.DataLabelAlign = "right";  // left/center/right

// 纵向柱状图
config.DataLabelAnchor = "end";
config.DataLabelAlign = "top";    // top/center/bottom
```

### 3. 支持 .xls 文件吗？

支持，但需要系统安装 Microsoft Excel。类库会自动将 .xls 转换为 .xlsx。

### 4. HTML 文件需要联网吗？

是的，生成的 HTML 使用 CDN 加载 Chart.js 库，需要联网才能正常显示图表。

---

## 依赖项

| 包名 | 版本 | 说明 |
|------|------|------|
| DocumentFormat.OpenXml | 3.3.0 | Excel 文件读写 |
| Magick.NET-Q16-AnyCPU | 14.9.1 | GIF 图像生成 |

---

## 许可证

MIT License
