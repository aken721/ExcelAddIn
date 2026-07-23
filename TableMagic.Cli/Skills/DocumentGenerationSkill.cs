using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TableMagic.Cli.Skills;

public class DocumentGenerationSkill : ISkill
{
    public string Name => "DocumentGeneration";
    public string Description => "Word文档批量生成：模板+占位符+图片插入";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "generate_documents", Description = "根据Excel数据和Word模板批量生成文档",
                Parameters = P(new[]{"templatePath","outputFolder"}, new[]{"fileName","sheetName","nameColumn","format"}), RequiredParameters = new List<string>{"templatePath","outputFolder"} },
            new() { Name = "preview_document", Description = "预览文档生成效果（生成第一条数据）",
                Parameters = P(new[]{"templatePath"}, new[]{"fileName","sheetName"}), RequiredParameters = new List<string>{"templatePath"} }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "generate_documents" => GenerateDocuments(arguments),
                "preview_document" => PreviewDocument(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult GenerateDocuments(Dictionary<string, object> args)
    {
        var templatePath = GetStr(args, "templatePath");
        var outputFolder = GetStr(args, "outputFolder");
        if (!System.IO.File.Exists(templatePath)) return new SkillResult { Success = false, Error = $"模板文件不存在: {templatePath}" };
        if (!System.IO.Directory.Exists(outputFolder)) System.IO.Directory.CreateDirectory(outputFolder);
        return SkillResult.Ok($"文档批量生成功能需要OpenXML SDK。请安装DocumentFormat.OpenXml NuGet包后使用，或使用VSTO插件模式获取完整功能。模板: {templatePath}, 输出: {outputFolder}");
    }

    private SkillResult PreviewDocument(Dictionary<string, object> args)
    {
        var templatePath = GetStr(args, "templatePath");
        if (!System.IO.File.Exists(templatePath)) return new SkillResult { Success = false, Error = $"模板文件不存在: {templatePath}" };
        return SkillResult.Ok($"文档预览功能需要OpenXML SDK。模板: {templatePath}");
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