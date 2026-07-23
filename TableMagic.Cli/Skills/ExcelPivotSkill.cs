using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelPivotSkill : ISkill
{
    private readonly IExcelProvider _provider;

    public ExcelPivotSkill(IExcelProvider provider) { _provider = provider; }

    public string Name => "ExcelPivot";
    public string Description => "Excel数据透视表操作";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "create_pivot_table",
                Description = "创建数据透视表",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "sourceRange", new { type = "string", description = "数据源范围" } },
                            { "pivotSheetName", new { type = "string", description = "透视表工作表名称" } },
                            { "rowFields", new { type = "string", description = "行字段（JSON数组，可选）" } },
                            { "columnFields", new { type = "string", description = "列字段（JSON数组，可选）" } },
                            { "valueFields", new { type = "string", description = "值字段（JSON对象，可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "sourceRange", "pivotSheetName" }
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
                "create_pivot_table" => SkillResult.Ok(_provider.CreatePivotTable(fn, sn,
                    GetStr(arguments, "sourceRange"), GetStr(arguments, "pivotSheetName"),
                    GetStr(arguments, "rowFields"), GetStr(arguments, "columnFields"),
                    GetStr(arguments, "valueFields"))),
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