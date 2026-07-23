using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelRangeSkill : ISkill
{
    private readonly IExcelProvider _provider;

    public ExcelRangeSkill(IExcelProvider provider) { _provider = provider; }

    public string Name => "ExcelRange";
    public string Description => "Excel区域操作：批量读写、公式、复制、清除";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "set_range_values",
                Description = "批量设置区域值",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rangeAddress", new { type = "string", description = "区域地址，如A1:D10" } },
                            { "data", new { type = "string", description = "JSON格式的二维数组数据" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rangeAddress", "data" }
            },
            new()
            {
                Name = "get_range_values",
                Description = "获取区域值",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rangeAddress", new { type = "string", description = "区域地址，如A1:D10" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rangeAddress" }
            },
            new()
            {
                Name = "copy_range",
                Description = "复制区域",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "sourceRange", new { type = "string", description = "源区域地址" } },
                            { "targetRange", new { type = "string", description = "目标区域地址" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "sourceRange", "targetRange" }
            },
            new()
            {
                Name = "clear_range",
                Description = "清除区域内容",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "rangeAddress", new { type = "string", description = "区域地址" } },
                            { "clearType", new { type = "string", description = "清除类型：all/contents/formats（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "rangeAddress" }
            },
            new()
            {
                Name = "get_used_range",
                Description = "获取工作表已使用范围",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string>()
            },
            new()
            {
                Name = "get_last_row",
                Description = "获取工作表最后使用的行号",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string>()
            }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            var fileName = GetStr(arguments, "fileName");
            var sheetName = GetStr(arguments, "sheetName");

            switch (toolName)
            {
                case "set_range_values":
                    {
                        var addr = GetStr(arguments, "rangeAddress");
                        var dataJson = GetStr(arguments, "data");
                        var data = ParseData(dataJson);
                        _provider.SetRangeValues(fileName, sheetName, addr, data);
                        return SkillResult.Ok($"成功设置区域 {addr} 的数据");
                    }
                case "get_range_values":
                    {
                        var addr = GetStr(arguments, "rangeAddress");
                        var values = _provider.GetRangeValues(fileName, sheetName, addr);
                        return SkillResult.Ok(FormatValues(values));
                    }
                case "copy_range":
                    {
                        _provider.CopyRange(fileName, sheetName, GetStr(arguments, "sourceRange"), GetStr(arguments, "targetRange"));
                        return SkillResult.Ok("复制成功");
                    }
                case "clear_range":
                    {
                        var clearType = GetStr(arguments, "clearType") ?? "all";
                        _provider.ClearRange(fileName, sheetName, GetStr(arguments, "rangeAddress"), clearType);
                        return SkillResult.Ok("清除成功");
                    }
                case "get_used_range":
                    return SkillResult.Ok(_provider.GetUsedRange(fileName, sheetName));
                case "get_last_row":
                    return SkillResult.Ok($"最后行号: {_provider.GetLastRow(fileName, sheetName)}");
                default:
                    return new SkillResult { Success = false, Error = $"未知工具: {toolName}" };
            }
        }
        catch (Exception ex)
        {
            return new SkillResult { Success = false, Error = ex.Message };
        }
    }

    private static object[,] ParseData(string json)
    {
        var doc = System.Text.Json.JsonDocument.Parse(json);
        var arr = doc.RootElement.EnumerateArray().ToList();
        var rows = arr.Count;
        var cols = arr.Count > 0 ? arr[0].EnumerateArray().Count() : 0;
        var data = new object[rows, cols];
        for (int r = 0; r < rows; r++)
        {
            var rowArr = arr[r].EnumerateArray().ToList();
            for (int c = 0; c < cols; c++)
                data[r, c] = rowArr[c].ValueKind == System.Text.Json.JsonValueKind.Number
                    ? rowArr[c].GetDouble()
                    : rowArr[c].ToString();
        }
        return data;
    }

    private static string FormatValues(object[,] values)
    {
        var sb = new System.Text.StringBuilder();
        int rows = values.GetLength(0);
        int cols = values.GetLength(1);
        for (int r = 0; r < rows; r++)
        {
            var row = new string[cols];
            for (int c = 0; c < cols; c++)
                row[c] = values[r, c]?.ToString() ?? "";
            sb.AppendLine(string.Join("\t", row));
        }
        return sb.ToString();
    }

    private static string GetStr(Dictionary<string, object> args, string key) => (args.ContainsKey(key) ? args[key]?.ToString() : null)!;
}