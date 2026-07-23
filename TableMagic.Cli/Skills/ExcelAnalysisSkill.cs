using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelAnalysisSkill : ISkill
{
    private readonly IExcelProvider _provider;

    public ExcelAnalysisSkill(IExcelProvider provider) { _provider = provider; }

    public string Name => "ExcelAnalysis";
    public string Description => "Excel数据分析：统计、分析";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "analyze_data",
                Description = "分析指定范围的数据",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "range", new { type = "string", description = "数据范围" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "range" }
            },
            new()
            {
                Name = "get_range_statistics",
                Description = "获取统计信息（最小值、最大值、平均值、总和等）",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "range", new { type = "string", description = "数据范围" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "range" }
            }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            var fn = GetStr(arguments, "fileName");
            var sn = GetStr(arguments, "sheetName");
            var range = GetStr(arguments, "range");

            return toolName switch
            {
                "analyze_data" => SkillResult.Ok(_provider.AnalyzeData(fn, sn, range)),
                "get_range_statistics" => SkillResult.Ok(_provider.GetRangeStatistics(fn, sn, range)),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex)
        {
            return new SkillResult { Success = false, Error = ex.Message };
        }
    }

    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
}