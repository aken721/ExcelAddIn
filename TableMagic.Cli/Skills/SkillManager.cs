using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TableMagic.Cli.Skills;

public class SkillManager
{
    private readonly List<ISkill> _skills = new();
    private readonly Dictionary<string, ISkill> _toolToSkillMap = new();
    private readonly Dictionary<string, SkillTool> _toolDefs = new();

    public void LoadSkill(ISkill skill)
    {
        _skills.Add(skill);
        foreach (var tool in skill.GetTools())
        {
            _toolToSkillMap[tool.Name] = skill;
            _toolDefs[tool.Name] = tool;
        }
    }

    public List<SkillTool> GetAllTools()
    {
        var allTools = new List<SkillTool>();
        foreach (var skill in _skills)
            allTools.AddRange(skill.GetTools());
        return allTools;
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        if (!_toolToSkillMap.TryGetValue(toolName, out var skill))
            return new SkillResult { Success = false, Error = $"工具 {toolName} 未找到" };

        var missing = ValidateRequiredParams(toolName, arguments);
        if (missing.Count > 0)
            return SkillResult.MissingParamsResult(toolName, missing);

        return await skill.ExecuteToolAsync(toolName, arguments);
    }

    public List<ISkill> GetLoadedSkills() => _skills;

    public bool IsToolAvailable(string toolName) => _toolToSkillMap.ContainsKey(toolName);

    public SkillTool? GetToolDef(string toolName)
    {
        return _toolDefs.TryGetValue(toolName, out var def) ? def : null;
    }

    private List<MissingParam> ValidateRequiredParams(string toolName, Dictionary<string, object> arguments)
    {
        var missing = new List<MissingParam>();
        if (!_toolDefs.TryGetValue(toolName, out var toolDef))
            return missing;

        if (toolDef.RequiredParameters == null || toolDef.RequiredParameters.Count == 0)
            return missing;

        var props = toolDef.Parameters?.GetValueOrDefault("properties") as Dictionary<string, object>;

        foreach (var reqParam in toolDef.RequiredParameters)
        {
            if (!arguments.ContainsKey(reqParam) || arguments[reqParam] == null)
            {
                var mp = new MissingParam { Name = reqParam };

                if (props != null && props.TryGetValue(reqParam, out var propObj))
                {
                    var propType = propObj.GetType();
                    var typeProp = propType.GetProperty("type");
                    var descProp = propType.GetProperty("description");
                    mp.Type = typeProp?.GetValue(propObj)?.ToString() ?? "string";
                    mp.Description = descProp?.GetValue(propObj)?.ToString() ?? reqParam;
                }
                else
                {
                    mp.Description = reqParam;
                }

                mp.PromptHint = BuildPromptHint(toolName, reqParam, mp.Description);
                missing.Add(mp);
            }
            else
            {
                var val = arguments[reqParam];
                if (val is string s && string.IsNullOrWhiteSpace(s))
                {
                    var mp = new MissingParam { Name = reqParam, Description = reqParam };
                    mp.PromptHint = BuildPromptHint(toolName, reqParam, reqParam);
                    missing.Add(mp);
                }
            }
        }

        return missing;
    }

    private static string BuildPromptHint(string toolName, string paramName, string paramDesc)
    {
        return toolName switch
        {
            "create_workbook" when paramName == "fileName" => "请提供工作簿文件名，例如：销售数据.xlsx",
            "open_workbook" when paramName == "fileName" => "请提供要打开的工作簿文件名",
            "set_cell_value" when paramName == "value" => "请提供要写入单元格的值",
            "set_cell_value" when paramName == "row" => "请提供行号（从1开始）",
            "set_cell_value" when paramName == "column" => "请提供列号（从1开始）",
            "create_worksheet" when paramName == "sheetName" => "请提供新工作表名称",
            "rename_worksheet" when paramName == "oldSheetName" => "请提供原工作表名称",
            "rename_worksheet" when paramName == "newSheetName" => "请提供新工作表名称",
            "set_range_values" when paramName == "rangeAddress" => "请提供区域地址，例如：A1:D10",
            "set_range_values" when paramName == "data" => "请提供数据，JSON二维数组格式，例如：[[\"姓名\",\"年龄\"],[\"张三\",25]]",
            "set_cell_format" when paramName == "rangeAddress" => "请提供要设置格式的区域地址",
            "set_border" when paramName == "borderType" => "请提供边框类型：all（全部）/outline（外框）/horizontal（水平）/vertical（垂直）",
            "create_chart" when paramName == "dataRange" => "请提供图表数据范围，例如：A1:D10",
            "create_pivot_table" when paramName == "sourceRange" => "请提供数据源范围",
            "create_pivot_table" when paramName == "pivotSheetName" => "请提供透视表工作表名称",
            "analyze_data" or "get_range_statistics" when paramName == "range" => "请提供要分析的数据范围",
            "calculate_financial_ratio" when paramName == "revenueRange" => "请提供收入数据范围",
            "calculate_financial_ratio" when paramName == "costRange" => "请提供成本数据范围",
            "calculate_profit_margin" when paramName == "revenueRange" => "请提供收入数据范围",
            "calculate_profit_margin" when paramName == "profitRange" => "请提供利润数据范围",
            _ => $"请提供 {paramName}：{paramDesc}"
        };
    }
}