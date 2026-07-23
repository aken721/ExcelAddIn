using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelSheetSkill : ISkill
{
    private readonly IExcelProvider _provider;

    public ExcelSheetSkill(IExcelProvider provider) { _provider = provider; }

    public string Name => "ExcelSheet";
    public string Description => "Excel工作表操作：激活、创建、重命名、删除、复制、移动、冻结";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "activate_worksheet",
                Description = "激活/切换工作表",
                Parameters = ParamDef(new[] { "sheetName" }, new[] { "fileName" }),
                RequiredParameters = new List<string> { "sheetName" }
            },
            new()
            {
                Name = "create_worksheet",
                Description = "创建新工作表",
                Parameters = ParamDef(new[] { "sheetName" }, new[] { "fileName" }),
                RequiredParameters = new List<string> { "sheetName" }
            },
            new()
            {
                Name = "rename_worksheet",
                Description = "重命名工作表",
                Parameters = ParamDef(new[] { "oldSheetName", "newSheetName" }, new[] { "fileName" }),
                RequiredParameters = new List<string> { "oldSheetName", "newSheetName" }
            },
            new()
            {
                Name = "delete_worksheet",
                Description = "删除工作表",
                Parameters = ParamDef(new[] { "sheetName" }, new[] { "fileName" }),
                RequiredParameters = new List<string> { "sheetName" }
            },
            new()
            {
                Name = "copy_worksheet",
                Description = "复制工作表",
                Parameters = ParamDef(new[] { "sourceSheetName", "targetSheetName" }, new[] { "fileName" }),
                RequiredParameters = new List<string> { "sourceSheetName", "targetSheetName" }
            },
            new()
            {
                Name = "move_worksheet",
                Description = "移动工作表位置",
                Parameters = ParamDef(new[] { "sheetName", "position" }, new[] { "fileName" }),
                RequiredParameters = new List<string> { "sheetName", "position" }
            },
            new()
            {
                Name = "set_worksheet_visible",
                Description = "设置工作表可见性",
                Parameters = ParamDef(new[] { "sheetName", "visible" }, new[] { "fileName" }),
                RequiredParameters = new List<string> { "sheetName", "visible" }
            },
            new()
            {
                Name = "get_worksheet_index",
                Description = "获取工作表索引",
                Parameters = ParamDef(new[] { "sheetName" }, new[] { "fileName" }),
                RequiredParameters = new List<string> { "sheetName" }
            },
            new()
            {
                Name = "freeze_panes",
                Description = "冻结窗格",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "row", new { type = "integer", description = "冻结行数（可选，默认1）" } },
                            { "column", new { type = "integer", description = "冻结列数（可选，默认1）" } }
                        }
                    }
                },
                RequiredParameters = new List<string>()
            },
            new()
            {
                Name = "unfreeze_panes",
                Description = "取消冻结窗格",
                Parameters = ParamDef(Array.Empty<string>(), new[] { "fileName", "sheetName" }),
                RequiredParameters = new List<string>()
            }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            var fileName = GetStr(arguments, "fileName");
            var sheetName = GetStr(arguments, "sheetName");

            return toolName switch
            {
                "activate_worksheet" => Exec(() => _provider.ActivateWorksheet(fileName, sheetName), r => $"已激活工作表: {r}"),
                "create_worksheet" => Exec(() => _provider.CreateWorksheet(fileName, sheetName), r => $"已创建工作表: {r}"),
                "rename_worksheet" => Exec(() => _provider.RenameWorksheet(fileName, GetStr(arguments, "oldSheetName"), GetStr(arguments, "newSheetName")), r => "重命名成功"),
                "delete_worksheet" => Exec(() => { _provider.DeleteWorksheet(fileName, sheetName); }, r => $"已删除工作表: {sheetName}"),
                "copy_worksheet" => Exec(() => _provider.CopyWorksheet(fileName, GetStr(arguments, "sourceSheetName"), GetStr(arguments, "targetSheetName")), r => "复制成功"),
                "move_worksheet" => Exec(() => _provider.MoveWorksheet(fileName, sheetName, GetInt(arguments, "position")), r => "移动成功"),
                "set_worksheet_visible" => Exec(() => _provider.SetWorksheetVisible(fileName, sheetName, GetBool(arguments, "visible")), r => $"已设置工作表可见性"),
                "get_worksheet_index" => Exec(() => _provider.GetWorksheetIndex(fileName, sheetName), r => $"工作表索引: {r}"),
                "freeze_panes" => Exec(() => _provider.FreezePanes(fileName, sheetName, GetInt(arguments, "row", 1), GetInt(arguments, "column", 1)), r => "已冻结窗格"),
                "unfreeze_panes" => Exec(() => _provider.UnfreezePanes(fileName, sheetName), r => "已取消冻结窗格"),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex)
        {
            return new SkillResult { Success = false, Error = ex.Message };
        }
    }

    private static Dictionary<string, object> ParamDef(string[] required, string[] optional)
    {
        var props = new Dictionary<string, object>();
        foreach (var r in required) props[r] = new { type = "string", description = $"{r}（必需）" };
        foreach (var o in optional) props[o] = new { type = "string", description = $"{o}（可选）" };
        return new Dictionary<string, object>
        {
            { "type", "object" },
            { "properties", props }
        };
    }

    private static string GetStr(Dictionary<string, object> args, string key) => (args.ContainsKey(key) ? args[key]?.ToString() : null)!;
    private static int GetInt(Dictionary<string, object> args, string key, int def = 0) => args.ContainsKey(key) && int.TryParse(args[key]?.ToString(), out var v) ? v : def;
    private static bool GetBool(Dictionary<string, object> args, string key) => args.ContainsKey(key) && bool.TryParse(args[key]?.ToString(), out var v) && v;

    private static SkillResult Exec(Action action, Func<object, string> successMsg)
    {
        try { action(); return SkillResult.Ok(successMsg(null!)); }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private static SkillResult Exec<T>(Func<T> func, Func<T, string> successMsg)
    {
        try { var r = func(); return SkillResult.Ok(successMsg(r)); }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }
}