using System.Collections.Generic;
using System.Threading.Tasks;

namespace TableMagic.Cli.Skills;

public interface ISkill
{
    string Name { get; }
    string Description { get; }
    List<SkillTool> GetTools();
    Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments);
}

public class SkillTool
{
    public string Name { get; set; } = "";
    public string Description { get; set; } = "";
    public Dictionary<string, object> Parameters { get; set; } = new();
    public List<string> RequiredParameters { get; set; } = new();
}

public class SkillResult
{
    public bool Success { get; set; }
    public string Content { get; set; } = "";
    public string Error { get; set; } = "";
    public List<string> Suggestions { get; set; } = new();
    public bool RequiresUserDecision { get; set; }
    public bool MissingRequiredParams { get; set; }
    public List<MissingParam> MissingParams { get; set; } = new();

    public static SkillResult FromError(string error, List<string>? suggestions = null, bool requiresUserDecision = false)
    {
        return new SkillResult
        {
            Success = false,
            Error = error,
            Suggestions = suggestions ?? new List<string>(),
            RequiresUserDecision = requiresUserDecision
        };
    }

    public static SkillResult Ok(string content)
    {
        return new SkillResult { Success = true, Content = content };
    }

    public static SkillResult NotSupported(string toolName)
    {
        return new SkillResult
        {
            Success = false,
            Error = $"工具 {toolName} 在当前模式下不支持，请使用Excel COM模式"
        };
    }

    public static SkillResult MissingParamsResult(string toolName, List<MissingParam> missingParams)
    {
        var paramDesc = string.Join("、", missingParams.ConvertAll(p => $"「{p.Name}」（{p.Description}）"));
        var promptHints = missingParams.ConvertAll(p => p.PromptHint ?? $"请提供 {p.Name}：{p.Description}");
        return new SkillResult
        {
            Success = false,
            Error = $"调用工具 {toolName} 缺少必需参数：{paramDesc}。请向用户追问这些参数后再调用。",
            MissingRequiredParams = true,
            MissingParams = missingParams,
            Suggestions = promptHints
        };
    }
}

public class MissingParam
{
    public string Name { get; set; } = "";
    public string Type { get; set; } = "string";
    public string Description { get; set; } = "";
    public string PromptHint { get; set; } = "";
}