# Skills 使用指南

## 概述

本项目采用 **Skills扩展架构**，将Excel操作功能模块化为独立的技能类。每个技能类封装一组相关的工具，便于维护和扩展。

## 架构优势

| 特性 | 说明 |
|------|------|
| 模块化设计 | 每个技能独立封装，职责单一 |
| 动态加载 | 新增技能只需添加Skill类文件，自动注册 |
| 统一管理 | SkillManager统一管理所有技能的注册和执行 |
| 参数别名 | 支持参数名称映射，兼容不同模型的参数命名 |
| 错误处理 | 统一的SkillResult返回格式 |

## 内置技能列表

### 1. ExcelBaseSkill - 基础操作

提供Excel基础信息获取功能。

| 工具名称 | 功能描述 | 必需参数 |
|---------|---------|---------|
| get_worksheet_names | 获取所有工作表名称 | 无 |

**使用示例**：
```
用户：查看当前工作簿有哪些工作表
AI将调用：get_worksheet_names
返回：Sheet1, Sheet2, Sheet3
```

---

### 2. ExcelWorkbookSkill - 工作簿操作

提供工作簿的创建、打开、保存、关闭等功能。

| 工具名称 | 功能描述 | 必需参数 | 可选参数 |
|---------|---------|---------|---------|
| create_workbook | 创建新的Excel工作簿 | fileName | sheetName |
| open_workbook | 打开工作簿文件 | fileName | - |
| close_workbook | 关闭工作簿 | - | fileName |
| save_workbook | 保存工作簿 | - | fileName |
| save_workbook_as | 另存为新文件 | newFileName | fileName |

**使用示例**：
```
用户：创建一个名为"销售数据.xlsx"的新工作簿
AI将调用：create_workbook
参数：{ "fileName": "销售数据.xlsx", "sheetName": "数据" }

用户：保存当前工作簿
AI将调用：save_workbook

用户：将当前工作簿另存为"备份.xlsx"
AI将调用：save_workbook_as
参数：{ "newFileName": "备份.xlsx" }
```

---

### 3. ExcelSheetSkill - 工作表操作

提供工作表的管理功能，包括创建、删除、重命名、复制、移动等。

| 工具名称 | 功能描述 | 必需参数 | 可选参数 |
|---------|---------|---------|---------|
| activate_worksheet | 激活/切换工作表 | sheetName | fileName |
| create_worksheet | 创建新工作表 | sheetName | fileName |
| rename_worksheet | 重命名工作表 | oldSheetName, newSheetName | fileName |
| delete_worksheet | 删除工作表 | sheetName | fileName |
| copy_worksheet | 复制工作表 | sourceSheetName, targetSheetName | fileName |
| move_worksheet | 移动工作表位置 | sheetName, position | fileName |
| set_worksheet_visible | 设置工作表可见性 | sheetName, visible | fileName |
| get_worksheet_index | 获取工作表索引 | sheetName | fileName |
| freeze_panes | 冻结窗格 | - | fileName, sheetName |
| unfreeze_panes | 取消冻结窗格 | - | fileName, sheetName |

**使用示例**：
```
用户：在当前工作簿新建2个工作表，分别命名为"一月"和"二月"
AI将调用：create_worksheet（两次）
参数1：{ "sheetName": "一月" }
参数2：{ "sheetName": "二月" }

用户：将Sheet1重命名为"数据源"
AI将调用：rename_worksheet
参数：{ "oldSheetName": "Sheet1", "newSheetName": "数据源" }

用户：切换到"汇总"工作表
AI将调用：activate_worksheet
参数：{ "sheetName": "汇总" }

用户：冻结第一行
AI将调用：freeze_panes
```

---

### 4. ExcelCellSkill - 单元格操作

提供单元格值的读取、写入、公式设置等功能。

| 工具名称 | 功能描述 | 必需参数 | 可选参数 |
|---------|---------|---------|---------|
| set_cell_value | 设置单元格值 | row, column, value | fileName, sheetName |
| get_cell_value | 获取单元格值 | row, column | fileName, sheetName |
| set_cell_formula | 设置单元格公式 | cellAddress, formula | fileName, sheetName |
| get_cell_formula | 获取单元格公式 | cellAddress | fileName, sheetName |

**使用示例**：
```
用户：在A1单元格写入"姓名"
AI将调用：set_cell_value
参数：{ "row": 1, "column": 1, "value": "姓名" }

用户：读取B2单元格的值
AI将调用：get_cell_value
参数：{ "row": 2, "column": 2 }

用户：在C10单元格设置求和公式
AI将调用：set_cell_formula
参数：{ "cellAddress": "C10", "formula": "=SUM(C1:C9)" }
```

---

### 5. ExcelRangeSkill - 区域操作

提供单元格区域的批量操作功能。

| 工具名称 | 功能描述 | 必需参数 | 可选参数 |
|---------|---------|---------|---------|
| set_range_values | 批量设置区域值 | rangeAddress, data | fileName, sheetName |
| get_range_values | 获取区域值 | rangeAddress | fileName, sheetName |
| set_formula | 设置公式 | cellAddress, formula | fileName, sheetName |
| get_formula | 获取公式 | cellAddress | fileName, sheetName |
| copy_range | 复制区域 | sourceRange, targetRange | fileName, sheetName |
| clear_range | 清除区域内容 | rangeAddress | fileName, sheetName |

**使用示例**：
```
用户：将A1:D10区域清空
AI将调用：clear_range
参数：{ "rangeAddress": "A1:D10" }

用户：复制A1:C10的数据到E1
AI将调用：copy_range
参数：{ "sourceRange": "A1:C10", "targetRange": "E1" }

用户：读取A1:D20的数据
AI将调用：get_range_values
参数：{ "rangeAddress": "A1:D20" }
```

---

### 6. ExcelFormatSkill - 格式设置

提供单元格格式化、边框设置、合并等功能。

| 工具名称 | 功能描述 | 必需参数 | 可选参数 |
|---------|---------|---------|---------|
| set_cell_format | 设置单元格格式 | rangeAddress | fontColor, backgroundColor, fontSize, bold, italic, horizontalAlignment, verticalAlignment |
| set_border | 设置边框 | rangeAddress, borderType | lineStyle |
| merge_cells | 合并单元格 | rangeAddress | - |
| unmerge_cells | 取消合并 | rangeAddress | - |
| set_cell_text_wrap | 设置自动换行 | rangeAddress | wrap |

**使用示例**：
```
用户：将A1:D1的背景色设置为黄色，字体加粗
AI将调用：set_cell_format
参数：{ "rangeAddress": "A1:D1", "backgroundColor": "黄色", "bold": true }

用户：给A1:D10添加边框
AI将调用：set_border
参数：{ "rangeAddress": "A1:D10", "borderType": "all" }

用户：合并A1:C1单元格
AI将调用：merge_cells
参数：{ "rangeAddress": "A1:C1" }

用户：将A1的内容居中对齐
AI将调用：set_cell_format
参数：{ "rangeAddress": "A1", "horizontalAlignment": "center", "verticalAlignment": "center" }
```

---

### 7. ExcelChartSkill - 图表操作

提供图表创建功能。

| 工具名称 | 功能描述 | 必需参数 | 可选参数 |
|---------|---------|---------|---------|
| create_chart | 创建图表 | dataRange | chartType, title |

**支持的图表类型**：
- `column` - 柱状图（默认）
- `line` - 折线图
- `pie` - 饼图
- `bar` - 条形图
- `area` - 面积图
- `scatter` - 散点图

**使用示例**：
```
用户：根据A1:D10的数据创建一个柱状图
AI将调用：create_chart
参数：{ "dataRange": "A1:D10", "chartType": "column", "title": "销售统计" }

用户：用B1:B10的数据画一个折线图
AI将调用：create_chart
参数：{ "dataRange": "A1:B10", "chartType": "line" }
```

---

### 8. ExcelPivotSkill - 数据透视表

提供数据透视表创建功能。

| 工具名称 | 功能描述 | 必需参数 | 可选参数 |
|---------|---------|---------|---------|
| create_pivot_table | 创建数据透视表 | sourceRange, pivotSheetName | rowFields, columnFields, valueFields |

**使用示例**：
```
用户：根据A1:E100的数据创建一个数据透视表
AI将调用：create_pivot_table
参数：{ 
  "sourceRange": "A1:E100", 
  "pivotSheetName": "透视表",
  "rowFields": "[\"产品\"]",
  "columnFields": "[\"月份\"]",
  "valueFields": "{\"销售额\":\"sum\"}"
}
```

---

### 9. ExcelAnalysisSkill - 数据分析

提供数据统计和分析功能。

| 工具名称 | 功能描述 | 必需参数 | 可选参数 |
|---------|---------|---------|---------|
| analyze_data | 分析指定范围的数据 | range | fileName, sheetName |
| get_range_statistics | 获取统计信息 | range | fileName, sheetName |

**使用示例**：
```
用户：分析A1:D100的数据
AI将调用：analyze_data
参数：{ "range": "A1:D100" }

用户：统计B列数据的最大值、最小值、平均值
AI将调用：get_range_statistics
参数：{ "range": "B:B" }
```

---

### 10. ExcelFinanceSkill - 财务分析

提供财务指标计算功能。

| 工具名称 | 功能描述 | 必需参数 | 可选参数 |
|---------|---------|---------|---------|
| calculate_financial_ratio | 计算财务比率 | revenueRange, costRange | fileName, sheetName |
| calculate_profit_margin | 计算利润率 | revenueRange, profitRange | fileName, sheetName |

**使用示例**：
```
用户：根据B列收入和C列成本计算财务比率
AI将调用：calculate_financial_ratio
参数：{ "revenueRange": "B2:B12", "costRange": "C2:C12" }

用户：计算利润率
AI将调用：calculate_profit_margin
参数：{ "revenueRange": "B2:B12", "profitRange": "D2:D12" }
```

---

## 工具分组

为提高AI选择工具的准确性，工具按功能分组：

| 分组 | 包含工具 | 触发关键词示例 |
|------|---------|--------------|
| cell_rw | 单元格读写、批量操作 | 写入、读取、单元格、公式、批量 |
| format | 格式设置 | 格式、颜色、字体、边框、合并、对齐、底纹 |
| row_col | 行列操作 | 行高、列宽、插入行、删除列、隐藏 |
| sheet | 工作表操作 | 工作表、新建表、重命名、删除表、冻结 |
| workbook | 工作簿操作 | 工作簿、新建文件、打开、保存、关闭 |
| data | 数据处理 | 排序、筛选、图表、透视表 |
| named | 命名区域 | 命名区域、定义名称 |
| link | 批注和超链接 | 批注、注释、超链接 |

---

## 扩展开发

### 新增技能步骤

1. 在 `Skills` 文件夹中创建新的技能类文件
2. 实现 `ISkill` 接口
3. 定义工具列表（`GetTools`方法）
4. 实现工具执行逻辑（`ExecuteToolAsync`方法）

### 示例代码

```csharp
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelAddIn.Skills
{
    public class MyCustomSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public MyCustomSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "MyCustom";
        public string Description => "自定义技能描述";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "my_tool",
                    Description = "工具描述",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "param1", new { type = "string", description = "参数1描述" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "param1" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "my_tool":
                        // 实现工具逻辑
                        return new SkillResult { Success = true, Content = "执行成功" };
                    default:
                        return new SkillResult { Success = false, Error = $"未知工具: {toolName}" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }
    }
}
```

---

## 注意事项

1. **参数可选性**：大多数工具的 `fileName` 和 `sheetName` 参数是可选的，不指定时使用当前活跃的工作簿/工作表
2. **行列号**：行列号从1开始，与Excel一致
3. **颜色格式**：支持颜色名称（如"红色"、"黄色"）和十六进制格式（如"#FF0000"）
4. **范围地址**：使用标准Excel格式，如"A1:D10"、"B:B"
5. **JSON数据**：批量数据使用JSON格式的二维数组

---

## 更新日志

**2026-03-23**
- 重构为Skills扩展架构
- 统一工具定义和执行逻辑
- 支持参数别名映射
- 添加详细的使用指南
