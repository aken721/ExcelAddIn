using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelCellSkill : ISkill
{
    private readonly IExcelProvider _provider;

    public ExcelCellSkill(IExcelProvider provider) { _provider = provider; }

    public string Name => "ExcelCell";
    public string Description => "Excel单元格操作：读写值、公式设置获取";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "set_cell_value",
                Description = "设置单元格的值",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "row", new { type = "integer", description = "行号（从1开始）" } },
                            { "column", new { type = "integer", description = "列号（从1开始）" } },
                            { "value", new { type = "string", description = "要设置的值" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "row", "column", "value" }
            },
            new()
            {
                Name = "get_cell_value",
                Description = "获取单元格的值",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "row", new { type = "integer", description = "行号（从1开始）" } },
                            { "column", new { type = "integer", description = "列号（从1开始）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "row", "column" }
            },
            new()
            {
                Name = "set_cell_formula",
                Description = "设置单元格的公式",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "cellAddress", new { type = "string", description = "单元格地址，如A1" } },
                            { "formula", new { type = "string", description = "公式，如=SUM(A1:A10)" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "cellAddress", "formula" }
            },
            new()
            {
                Name = "get_cell_formula",
                Description = "获取单元格的公式",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "cellAddress", new { type = "string", description = "单元格地址，如A1" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "cellAddress" }
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
                case "set_cell_value":
                    {
                        var row = GetInt(arguments, "row");
                        var col = GetInt(arguments, "column");
                        var value = arguments["value"];
                        _provider.SetCellValue(fileName, sheetName, row, col, value);
                        return SkillResult.Ok($"成功设置单元格 ({row},{col}) 的值为: {value}");
                    }
                case "get_cell_value":
                    {
                        var row = GetInt(arguments, "row");
                        var col = GetInt(arguments, "column");
                        var val = _provider.GetCellValue(fileName, sheetName, row, col);
                        return SkillResult.Ok($"单元格 ({row},{col}) 的值为: {val}");
                    }
                case "set_cell_formula":
                    {
                        var addr = GetStr(arguments, "cellAddress");
                        var formula = GetStr(arguments, "formula");
                        _provider.SetFormula(fileName, sheetName, addr, formula);
                        return SkillResult.Ok($"成功设置单元格 {addr} 的公式为: {formula}");
                    }
                case "get_cell_formula":
                    {
                        var addr = GetStr(arguments, "cellAddress");
                        var formula = _provider.GetFormula(fileName, sheetName, addr);
                        return SkillResult.Ok($"单元格 {addr} 的公式为: {formula}");
                    }
                default:
                    return new SkillResult { Success = false, Error = $"未知工具: {toolName}" };
            }
        }
        catch (Exception ex)
        {
            return new SkillResult { Success = false, Error = ex.Message };
        }
    }

    private static string GetStr(Dictionary<string, object> args, string key) => (args.ContainsKey(key) ? args[key]?.ToString() : null)!;
    private static int GetInt(Dictionary<string, object> args, string key) => args.ContainsKey(key) && int.TryParse(args[key]?.ToString(), out var v) ? v : 0;
}