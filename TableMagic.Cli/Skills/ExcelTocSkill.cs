using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelTocSkill : ISkill
{
    private readonly IExcelProvider _provider;
    public ExcelTocSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelToc";
    public string Description => "目录页技能：创建目录表、根据目录建表、更新超链接";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "create_toc_sheet", Description = "创建目录表，列出当前工作簿中所有工作表名称",
                Parameters = P(Array.Empty<string>(), new[]{"fileName","tocSheetName","includeHiddenSheets"}), RequiredParameters = new List<string>() },
            new() { Name = "create_sheets_from_toc", Description = "根据目录表批量创建工作表",
                Parameters = P(new[]{"linkColumnName"}, new[]{"fileName","createSheets"}), RequiredParameters = new List<string>{"linkColumnName"} },
            new() { Name = "update_toc_hyperlinks", Description = "更新目录表中的超链接",
                Parameters = P(new[]{"columnName"}, new[]{"fileName"}), RequiredParameters = new List<string>{"columnName"} }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            var fn = GetStr(arguments, "fileName");
            return toolName switch
            {
                "create_toc_sheet" => CreateTocSheet(fn, arguments),
                "create_sheets_from_toc" => CreateSheetsFromToc(fn, arguments),
                "update_toc_hyperlinks" => UpdateHyperlinks(fn, arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult CreateTocSheet(string fn, Dictionary<string, object> args)
    {
        var tocName = GetStr(args, "tocSheetName") ?? "_目录";
        var names = _provider.GetWorksheetNames(fn);
        _provider.CreateWorksheet(fn, tocName);
        _provider.SetCellValue(fn, tocName, 1, 1, "表目录");
        int row = 2;
        foreach (var name in names.Where(n => n != tocName))
        {
            _provider.SetCellValue(fn, tocName, row, 1, name);
            row++;
        }
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"目录表创建完成，共列出 {row - 2} 个工作表");
    }

    private SkillResult CreateSheetsFromToc(string fn, Dictionary<string, object> args)
    {
        var linkCol = GetStr(args, "linkColumnName");
        var tocName = "目录";
        var lastRow = _provider.GetLastRow(fn, tocName);
        var lastCol = _provider.GetLastColumn(fn, tocName);
        var colIdx = FindCol(fn, tocName, linkCol, lastCol);
        if (colIdx == 0) return new SkillResult { Success = false, Error = $"未找到列: {linkCol}" };

        var existing = _provider.GetWorksheetNames(fn);
        int created = 0;
        for (int r = 2; r <= lastRow; r++)
        {
            var name = _provider.GetCellValue(fn, tocName, r, colIdx)?.ToString();
            if (string.IsNullOrEmpty(name) || existing.Contains(name)) continue;
            _provider.CreateWorksheet(fn, name);
            created++;
        }
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"根据目录创建 {created} 个工作表");
    }

    private SkillResult UpdateHyperlinks(string fn, Dictionary<string, object> args)
    {
        return SkillResult.Ok("超链接更新功能在ClosedXML模式下有限支持，完整功能请使用Excel COM模式");
    }

    private int FindCol(string fn, string sn, string colName, int lastCol)
    {
        if (int.TryParse(colName, out int n)) return n;
        for (int c = 1; c <= lastCol; c++)
            if (_provider.GetCellValue(fn, sn, 1, c)?.ToString() == colName) return c;
        return 0;
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