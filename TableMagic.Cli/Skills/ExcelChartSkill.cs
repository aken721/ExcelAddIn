using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelChartSkill : ISkill
{
    private readonly IExcelProvider _provider;

    public ExcelChartSkill(IExcelProvider provider) { _provider = provider; }

    public string Name => "ExcelChart";
    public string Description => "Excel图表操作：创建柱状图、折线图、饼图等";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "create_chart",
                Description = "创建图表（支持column/line/pie/bar/area/scatter）",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "dataRange", new { type = "string", description = "数据范围，如A1:D10" } },
                            { "chartType", new { type = "string", description = "图表类型：column/line/pie/bar/area/scatter（可选，默认column）" } },
                            { "title", new { type = "string", description = "图表标题（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "dataRange" }
            }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            var fn = GetStr(arguments, "fileName");
            var sn = GetStr(arguments, "sheetName");

            return toolName switch
            {
                "create_chart" => SkillResult.Ok(_provider.CreateChart(fn, sn,
                    GetStr(arguments, "dataRange"),
                    GetStr(arguments, "chartType") ?? "column",
                    GetStr(arguments, "title") ?? "")),
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