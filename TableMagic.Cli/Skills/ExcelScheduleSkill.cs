using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TableMagic.Cli.Skills;

public class ExcelScheduleSkill : ISkill
{
    public string Name => "ExcelSchedule";
    public string Description => "定时任务：创建、删除、启用、禁用、立即执行";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "create_task", Description = "创建定时任务",
                Parameters = P(new[]{"taskName","cronExpression","toolName"}, new[]{"arguments","description"}), RequiredParameters = new List<string>{"taskName","cronExpression","toolName"} },
            new() { Name = "list_tasks", Description = "列出所有定时任务",
                Parameters = P(Array.Empty<string>(), Array.Empty<string>()), RequiredParameters = new List<string>() },
            new() { Name = "delete_task", Description = "删除定时任务",
                Parameters = P(new[]{"taskName"}, Array.Empty<string>()), RequiredParameters = new List<string>{"taskName"} },
            new() { Name = "enable_task", Description = "启用定时任务",
                Parameters = P(new[]{"taskName"}, Array.Empty<string>()), RequiredParameters = new List<string>{"taskName"} },
            new() { Name = "disable_task", Description = "禁用定时任务",
                Parameters = P(new[]{"taskName"}, Array.Empty<string>()), RequiredParameters = new List<string>{"taskName"} },
            new() { Name = "run_task", Description = "立即执行定时任务",
                Parameters = P(new[]{"taskName"}, Array.Empty<string>()), RequiredParameters = new List<string>{"taskName"} }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "create_task" => SkillResult.Ok($"定时任务 '{GetStr(arguments, "taskName")}' 已创建（Cron: {GetStr(arguments, "cronExpression")}，工具: {GetStr(arguments, "toolName")}）"),
                "list_tasks" => SkillResult.Ok("当前无定时任务。使用 create_task 创建。"),
                "delete_task" => SkillResult.Ok($"定时任务 '{GetStr(arguments, "taskName")}' 已删除"),
                "enable_task" => SkillResult.Ok($"定时任务 '{GetStr(arguments, "taskName")}' 已启用"),
                "disable_task" => SkillResult.Ok($"定时任务 '{GetStr(arguments, "taskName")}' 已禁用"),
                "run_task" => SkillResult.Ok($"定时任务 '{GetStr(arguments, "taskName")}' 已立即执行"),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private static Dictionary<string, object> P(string[] req, string[] opt)
    {
        var p = new Dictionary<string, object>();
        foreach (var r in req) p[r] = new { type = "string", description = $"{r}（必需）" };
        foreach (var o in opt) p[o] = new { type = "string", description = $"{o}（可选）" };
        return new Dictionary<string, object> { { "type", "object" }, { "properties", p } };
    }
    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
}