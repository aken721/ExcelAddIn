using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelFormatSkill : ISkill
{
    private readonly IExcelProvider _provider;

    public ExcelFormatSkill(IExcelProvider provider) { _provider = provider; }

    public string Name => "ExcelFormat";
    public string Description => "Excel格式设置：字体、背景、边框、合并、换行、行列操作";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "set_cell_format",
                Description = "设置单元格格式（字体颜色、背景色、字号、加粗、斜体、对齐）",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rangeAddress", new { type = "string", description = "区域地址" } },
                            { "fontColor", new { type = "string", description = "字体颜色（可选）" } },
                            { "backgroundColor", new { type = "string", description = "背景颜色（可选）" } },
                            { "fontSize", new { type = "integer", description = "字号（可选）" } },
                            { "bold", new { type = "boolean", description = "是否加粗（可选）" } },
                            { "italic", new { type = "boolean", description = "是否斜体（可选）" } },
                            { "horizontalAlignment", new { type = "string", description = "水平对齐：left/center/right（可选）" } },
                            { "verticalAlignment", new { type = "string", description = "垂直对齐：top/center/bottom（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rangeAddress" }
            },
            new()
            {
                Name = "set_border",
                Description = "设置边框",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rangeAddress", new { type = "string", description = "区域地址" } },
                            { "borderType", new { type = "string", description = "边框类型：all/outline/horizontal/vertical" } },
                            { "lineStyle", new { type = "string", description = "线条样式：continuous/dash/dot（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rangeAddress", "borderType" }
            },
            new()
            {
                Name = "merge_cells",
                Description = "合并单元格",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rangeAddress", new { type = "string", description = "区域地址" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rangeAddress" }
            },
            new()
            {
                Name = "unmerge_cells",
                Description = "取消合并单元格",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rangeAddress", new { type = "string", description = "区域地址" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rangeAddress" }
            },
            new()
            {
                Name = "set_cell_text_wrap",
                Description = "设置自动换行",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rangeAddress", new { type = "string", description = "区域地址" } },
                            { "wrap", new { type = "boolean", description = "是否自动换行" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rangeAddress", "wrap" }
            },
            new()
            {
                Name = "set_row_height",
                Description = "设置行高",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rowNumber", new { type = "integer", description = "行号" } },
                            { "height", new { type = "number", description = "行高" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rowNumber", "height" }
            },
            new()
            {
                Name = "set_column_width",
                Description = "设置列宽",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "columnNumber", new { type = "integer", description = "列号" } },
                            { "width", new { type = "number", description = "列宽" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "columnNumber", "width" }
            },
            new()
            {
                Name = "insert_rows",
                Description = "插入行",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rowIndex", new { type = "integer", description = "行索引" } },
                            { "count", new { type = "integer", description = "插入行数（可选，默认1）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rowIndex" }
            },
            new()
            {
                Name = "insert_columns",
                Description = "插入列",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "columnIndex", new { type = "integer", description = "列索引" } },
                            { "count", new { type = "integer", description = "插入列数（可选，默认1）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "columnIndex" }
            },
            new()
            {
                Name = "delete_rows",
                Description = "删除行",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rowIndex", new { type = "integer", description = "行索引" } },
                            { "count", new { type = "integer", description = "删除行数（可选，默认1）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rowIndex" }
            },
            new()
            {
                Name = "delete_columns",
                Description = "删除列",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "columnIndex", new { type = "integer", description = "列索引" } },
                            { "count", new { type = "integer", description = "删除列数（可选，默认1）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "columnIndex" }
            }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            var fn = GetStr(arguments, "fileName");
            var sn = GetStr(arguments, "sheetName");

            switch (toolName)
            {
                case "set_cell_format":
                    _provider.SetCellFormat(fn, sn, GetStr(arguments, "rangeAddress"),
                        GetStr(arguments, "fontColor"), GetStr(arguments, "backgroundColor"),
                        GetNInt(arguments, "fontSize"), GetNBool(arguments, "bold"),
                        GetNBool(arguments, "italic"), GetStr(arguments, "horizontalAlignment"),
                        GetStr(arguments, "verticalAlignment"));
                    return SkillResult.Ok("格式设置成功");
                case "set_border":
                    _provider.SetBorder(fn, sn, GetStr(arguments, "rangeAddress"),
                        GetStr(arguments, "borderType"), GetStr(arguments, "lineStyle") ?? "continuous");
                    return SkillResult.Ok("边框设置成功");
                case "merge_cells":
                    _provider.MergeCells(fn, sn, GetStr(arguments, "rangeAddress"));
                    return SkillResult.Ok("合并成功");
                case "unmerge_cells":
                    _provider.UnmergeCells(fn, sn, GetStr(arguments, "rangeAddress"));
                    return SkillResult.Ok("取消合并成功");
                case "set_cell_text_wrap":
                    _provider.SetCellTextWrap(fn, sn, GetStr(arguments, "rangeAddress"), GetBool(arguments, "wrap"));
                    return SkillResult.Ok("换行设置成功");
                case "set_row_height":
                    _provider.SetRowHeight(fn, sn, GetInt(arguments, "rowNumber"), GetDbl(arguments, "height"));
                    return SkillResult.Ok("行高设置成功");
                case "set_column_width":
                    _provider.SetColumnWidth(fn, sn, GetInt(arguments, "columnNumber"), GetDbl(arguments, "width"));
                    return SkillResult.Ok("列宽设置成功");
                case "insert_rows":
                    _provider.InsertRows(fn, sn, GetInt(arguments, "rowIndex"), GetInt(arguments, "count", 1));
                    return SkillResult.Ok("插入行成功");
                case "insert_columns":
                    _provider.InsertColumns(fn, sn, GetInt(arguments, "columnIndex"), GetInt(arguments, "count", 1));
                    return SkillResult.Ok("插入列成功");
                case "delete_rows":
                    _provider.DeleteRows(fn, sn, GetInt(arguments, "rowIndex"), GetInt(arguments, "count", 1));
                    return SkillResult.Ok("删除行成功");
                case "delete_columns":
                    _provider.DeleteColumns(fn, sn, GetInt(arguments, "columnIndex"), GetInt(arguments, "count", 1));
                    return SkillResult.Ok("删除列成功");
                default:
                    return new SkillResult { Success = false, Error = $"未知工具: {toolName}" };
            }
        }
        catch (Exception ex)
        {
            return new SkillResult { Success = false, Error = ex.Message };
        }
    }

    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
    private static int GetInt(Dictionary<string, object> a, string k, int d = 0) => a.ContainsKey(k) && int.TryParse(a[k]?.ToString(), out var v) ? v : d;
    private static double GetDbl(Dictionary<string, object> a, string k) => a.ContainsKey(k) && double.TryParse(a[k]?.ToString(), out var v) ? v : 0;
    private static bool GetBool(Dictionary<string, object> a, string k) => a.ContainsKey(k) && bool.TryParse(a[k]?.ToString(), out var v) && v;
    private static int? GetNInt(Dictionary<string, object> a, string k) => a.ContainsKey(k) && int.TryParse(a[k]?.ToString(), out var v) ? v : null;
    private static bool? GetNBool(Dictionary<string, object> a, string k) => a.ContainsKey(k) && bool.TryParse(a[k]?.ToString(), out var v) ? v : null;
}