using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelFinanceSkill : ISkill
{
    private readonly IExcelProvider _provider;

    public ExcelFinanceSkill(IExcelProvider provider) { _provider = provider; }

    public string Name => "ExcelFinance";
    public string Description => "Excel财务分析：财务比率、利润率计算";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "calculate_financial_ratio",
                Description = "计算财务比率",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "revenueRange", new { type = "string", description = "收入数据范围" } },
                            { "costRange", new { type = "string", description = "成本数据范围" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "revenueRange", "costRange" }
            },
            new()
            {
                Name = "calculate_profit_margin",
                Description = "计算利润率",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "revenueRange", new { type = "string", description = "收入数据范围" } },
                            { "profitRange", new { type = "string", description = "利润数据范围" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "revenueRange", "profitRange" }
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
                case "calculate_financial_ratio":
                    {
                        var revStats = _provider.GetRangeStatistics(fn, sn, GetStr(arguments, "revenueRange"));
                        var costStats = _provider.GetRangeStatistics(fn, sn, GetStr(arguments, "costRange"));
                        return SkillResult.Ok($"财务比率分析:\n收入统计:\n{revStats}\n成本统计:\n{costStats}");
                    }
                case "calculate_profit_margin":
                    {
                        var revStats = _provider.GetRangeStatistics(fn, sn, GetStr(arguments, "revenueRange"));
                        var profitStats = _provider.GetRangeStatistics(fn, sn, GetStr(arguments, "profitRange"));
                        return SkillResult.Ok($"利润率分析:\n收入统计:\n{revStats}\n利润统计:\n{profitStats}");
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

    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
}