using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelWorkbookSkill : ISkill
{
    private readonly IExcelProvider _provider;

    public ExcelWorkbookSkill(IExcelProvider provider) { _provider = provider; }

    public string Name => "ExcelWorkbook";
    public string Description => "Excel工作簿操作：创建、打开、保存、关闭、另存为";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "create_workbook",
                Description = "创建新的Excel工作簿文件",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "新工作簿文件名（必需）" } },
                            { "sheetName", new { type = "string", description = "初始工作表名称（可选，默认Sheet1）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "fileName" }
            },
            new()
            {
                Name = "open_workbook",
                Description = "打开工作簿文件",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "要打开的工作簿文件名（必需）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "fileName" }
            },
            new()
            {
                Name = "close_workbook",
                Description = "关闭工作簿",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选，默认关闭当前活跃工作簿）" } }
                        }
                    }
                },
                RequiredParameters = new List<string>()
            },
            new()
            {
                Name = "save_workbook",
                Description = "保存工作簿",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string>()
            },
            new()
            {
                Name = "save_workbook_as",
                Description = "将工作簿另存为新文件",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "原工作簿文件名（可选）" } },
                            { "newFileName", new { type = "string", description = "新文件名（必需）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "newFileName" }
            },
            new()
            {
                Name = "delete_workbook",
                Description = "删除工作簿文件",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "要删除的工作簿文件名（必需）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "fileName" }
            }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"]?.ToString() : null;

            return toolName switch
            {
                "create_workbook" => ExecuteCreate(fileName!, arguments),
                "open_workbook" => ExecuteOpen(fileName!),
                "close_workbook" => ExecuteClose(fileName!),
                "save_workbook" => ExecuteSave(fileName!),
                "save_workbook_as" => ExecuteSaveAs(fileName!, arguments),
                "delete_workbook" => ExecuteDelete(fileName!),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex)
        {
            return new SkillResult { Success = false, Error = ex.Message };
        }
    }

    private SkillResult ExecuteCreate(string fileName, Dictionary<string, object> arguments)
    {
        if (string.IsNullOrEmpty(fileName))
            return new SkillResult { Success = false, Error = "必须指定文件名" };
        var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"]?.ToString() ?? "Sheet1" : "Sheet1";
        var result = _provider.CreateWorkbook(fileName, sheetName);
        return SkillResult.Ok($"成功创建工作簿: {result}");
    }

    private SkillResult ExecuteOpen(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            return new SkillResult { Success = false, Error = "必须指定文件名" };
        _provider.OpenWorkbook(fileName);
        return SkillResult.Ok($"成功打开工作簿: {fileName}");
    }

    private SkillResult ExecuteClose(string fileName)
    {
        _provider.CloseWorkbook(fileName);
        return SkillResult.Ok($"成功关闭工作簿: {fileName ?? "当前活跃工作簿"}");
    }

    private SkillResult ExecuteSave(string fileName)
    {
        _provider.SaveWorkbook(fileName);
        return SkillResult.Ok($"成功保存工作簿: {fileName ?? "当前活跃工作簿"}");
    }

    private SkillResult ExecuteSaveAs(string fileName, Dictionary<string, object> arguments)
    {
        var newFileName = arguments.ContainsKey("newFileName") ? arguments["newFileName"]?.ToString() : null;
        if (string.IsNullOrEmpty(newFileName))
            return new SkillResult { Success = false, Error = "必须指定新文件名" };
        _provider.SaveWorkbookAs(fileName, newFileName);
        return SkillResult.Ok($"成功另存为: {newFileName}");
    }

    private SkillResult ExecuteDelete(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
            return new SkillResult { Success = false, Error = "必须指定文件名" };
        _provider.DeleteWorkbook(fileName);
        return SkillResult.Ok($"成功删除工作簿: {fileName}");
    }
}