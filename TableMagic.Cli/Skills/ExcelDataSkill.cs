using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelDataSkill : ISkill
{
    private readonly IExcelProvider _provider;
    public ExcelDataSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelData";
    public string Description => "Excel数据处理：分表、并表、批量导删、转置、工资条";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "split_sheet_by_column", Description = "根据指定列的值将工作表拆分为多个工作表",
                Parameters = ParamDef(new[]{"columnName"}, new[]{"fileName","sheetName","dataStartRow"}), RequiredParameters = new List<string>{"columnName"} },
            new() { Name = "split_and_export", Description = "根据指定列分表并导出为独立文件",
                Parameters = ParamDef(new[]{"columnName","outputFolder"}, new[]{"fileName","sheetName","fileFormat"}), RequiredParameters = new List<string>{"columnName","outputFolder"} },
            new() { Name = "merge_sheets", Description = "将多个工作表合并为一个工作表",
                Parameters = ParamDef(Array.Empty<string>(), new[]{"fileName","sheetNames","outputSheetName","includeHeader"}), RequiredParameters = new List<string>() },
            new() { Name = "merge_workbooks", Description = "将指定目录下所有工作簿的工作表合并到当前工作簿",
                Parameters = ParamDef(new[]{"folderPath"}, new[]{"fileName","includeSubfolders","skipEmptySheets"}), RequiredParameters = new List<string>{"folderPath"} },
            new() { Name = "export_sheets", Description = "批量导出工作表为独立文件",
                Parameters = ParamDef(new[]{"outputFolder"}, new[]{"fileName","sheetNames","fileFormat"}), RequiredParameters = new List<string>{"outputFolder"} },
            new() { Name = "delete_sheets", Description = "批量删除工作表",
                Parameters = ParamDef(new[]{"sheetNames"}, new[]{"fileName"}), RequiredParameters = new List<string>{"sheetNames"} },
            new() { Name = "transpose_columns", Description = "将列名称转置为字段内数据（宽表转长表）",
                Parameters = ParamDef(new[]{"startColumn","fieldName"}, new[]{"fileName","sheetName"}), RequiredParameters = new List<string>{"startColumn","fieldName"} },
            new() { Name = "create_multiple_sheets", Description = "批量创建指定数量的工作表",
                Parameters = ParamDef(new[]{"count"}, new[]{"fileName","baseName"}), RequiredParameters = new List<string>{"count"} },
            new() { Name = "generate_payslips", Description = "将工资表转换为工资条格式（每行数据前插入标题行）",
                Parameters = ParamDef(Array.Empty<string>(), new[]{"fileName","sheetName","outputSheetName"}), RequiredParameters = new List<string>() }
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
                "split_sheet_by_column" => SplitSheetByColumn(fn, sn, arguments),
                "split_and_export" => SplitAndExport(fn, sn, arguments),
                "merge_sheets" => MergeSheets(fn, arguments),
                "merge_workbooks" => MergeWorkbooks(fn, arguments),
                "export_sheets" => ExportSheets(fn, arguments),
                "delete_sheets" => DeleteSheets(fn, arguments),
                "transpose_columns" => TransposeColumns(fn, sn, arguments),
                "create_multiple_sheets" => CreateMultipleSheets(fn, arguments),
                "generate_payslips" => GeneratePayslips(fn, sn, arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult SplitSheetByColumn(string fn, string sn, Dictionary<string, object> args)
    {
        var columnName = GetStr(args, "columnName");
        var dataStartRow = GetInt(args, "dataStartRow", 2);
        var lastRow = _provider.GetLastRow(fn, sn);
        var lastCol = _provider.GetLastColumn(fn, sn);
        var colIndex = FindColumn(fn, sn, columnName, lastCol);
        if (colIndex == 0) return new SkillResult { Success = false, Error = $"未找到列: {columnName}" };

        var uniqueValues = new HashSet<string>();
        for (int r = dataStartRow; r <= lastRow; r++)
        {
            var val = _provider.GetCellValue(fn, sn, r, colIndex)?.ToString() ?? "";
            if (!string.IsNullOrEmpty(val)) uniqueValues.Add(val);
        }

        int created = 0;
        foreach (var val in uniqueValues)
        {
            var newSheetName = val.Length > 31 ? val[..31] : val;
            _provider.CreateWorksheet(fn, newSheetName);
            for (int c = 1; c <= lastCol; c++)
                _provider.SetCellValue(fn, newSheetName, 1, c, _provider.GetCellValue(fn, sn, 1, c));
            int destRow = 2;
            for (int r = dataStartRow; r <= lastRow; r++)
            {
                if (_provider.GetCellValue(fn, sn, r, colIndex)?.ToString() == val)
                {
                    for (int c = 1; c <= lastCol; c++)
                        _provider.SetCellValue(fn, newSheetName, destRow, c, _provider.GetCellValue(fn, sn, r, c));
                    destRow++;
                }
            }
            created++;
        }
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"分表完成，共创建 {created} 个工作表");
    }

    private SkillResult SplitAndExport(string fn, string sn, Dictionary<string, object> args)
    {
        var outputFolder = GetStr(args, "outputFolder");
        if (!Directory.Exists(outputFolder)) Directory.CreateDirectory(outputFolder);

        var existingSheets = new HashSet<string>(_provider.GetWorksheetNames(fn));
        var splitResult = SplitSheetByColumn(fn, sn, args);
        if (!splitResult.Success) return splitResult;

        var allSheets = _provider.GetWorksheetNames(fn);
        var newSheets = allSheets.Where(n => !existingSheets.Contains(n)).ToList();
        int exported = 0;
        foreach (var name in newSheets)
        {
            var newFn = Path.Combine(outputFolder, $"{name}.xlsx");
            _provider.CreateWorkbook(newFn, name);
            var lastRow = _provider.GetLastRow(fn, name);
            var lastCol = _provider.GetLastColumn(fn, name);
            for (int r = 1; r <= lastRow; r++)
                for (int c = 1; c <= lastCol; c++)
                    _provider.SetCellValue(newFn, name, r, c, _provider.GetCellValue(fn, name, r, c));
            _provider.SaveWorkbook(newFn);
            exported++;
        }
        return SkillResult.Ok($"分表并导出完成，共导出 {exported} 个文件到 {outputFolder}");
    }

    private SkillResult MergeSheets(string fn, Dictionary<string, object> args)
    {
        var outputSn = GetStr(args, "outputSheetName") ?? "合并表";
        var includeHeader = GetBool(args, "includeHeader", true);
        var sheetNamesJson = GetStr(args, "sheetNames");
        var allNames = _provider.GetWorksheetNames(fn);
        var targetSheets = string.IsNullOrEmpty(sheetNamesJson)
            ? allNames.Where(n => n != outputSn).ToList()
            : System.Text.Json.JsonSerializer.Deserialize<List<string>>(sheetNamesJson)!;

        _provider.CreateWorksheet(fn, outputSn);
        int destRow = 1;
        int colCount = 0;
        foreach (var srcName in targetSheets)
        {
            var lastRow = _provider.GetLastRow(fn, srcName);
            var lastCol = _provider.GetLastColumn(fn, srcName);
            if (lastRow <= 0) continue;
            colCount = Math.Max(colCount, lastCol);
            int startRow = (destRow == 1 || includeHeader) ? 1 : 2;
            for (int r = startRow; r <= lastRow; r++)
                for (int c = 1; c <= lastCol; c++)
                    _provider.SetCellValue(fn, outputSn, destRow + r - startRow, c, _provider.GetCellValue(fn, srcName, r, c));
            destRow += lastRow - startRow + 1;
        }
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"合并完成，共合并 {targetSheets.Count} 个工作表到 '{outputSn}'");
    }

    private SkillResult MergeWorkbooks(string fn, Dictionary<string, object> args)
    {
        var folderPath = GetStr(args, "folderPath");
        if (!Directory.Exists(folderPath)) return new SkillResult { Success = false, Error = $"文件夹不存在: {folderPath}" };
        var files = Directory.GetFiles(folderPath, "*.xlsx");
        int merged = 0;
        foreach (var file in files)
        {
            var srcFn = Path.GetFileName(file);
            if (srcFn == fn) continue;
            _provider.OpenWorkbook(srcFn);
            foreach (var sn in _provider.GetWorksheetNames(srcFn))
            {
                var newSn = $"{Path.GetFileNameWithoutExtension(srcFn)}_{sn}";
                if (newSn.Length > 31) newSn = newSn[..31];
                _provider.CreateWorksheet(fn, newSn);
                var lastRow = _provider.GetLastRow(srcFn, sn);
                var lastCol = _provider.GetLastColumn(srcFn, sn);
                for (int r = 1; r <= lastRow; r++)
                    for (int c = 1; c <= lastCol; c++)
                        _provider.SetCellValue(fn, newSn, r, c, _provider.GetCellValue(srcFn, sn, r, c));
                merged++;
            }
            _provider.CloseWorkbook(srcFn);
        }
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"合并工作簿完成，共合并 {files.Length} 个文件中的 {merged} 个工作表");
    }

    private SkillResult ExportSheets(string fn, Dictionary<string, object> args)
    {
        var outputFolder = GetStr(args, "outputFolder");
        if (!Directory.Exists(outputFolder)) Directory.CreateDirectory(outputFolder);
        var sheetNamesJson = GetStr(args, "sheetNames");
        var allNames = _provider.GetWorksheetNames(fn);
        var targetSheets = string.IsNullOrEmpty(sheetNamesJson) ? allNames : System.Text.Json.JsonSerializer.Deserialize<List<string>>(sheetNamesJson)!;
        int exported = 0;
        foreach (var sn in targetSheets)
        {
            var newFn = $"{sn}.xlsx";
            _provider.CreateWorkbook(newFn, sn);
            var lastRow = _provider.GetLastRow(fn, sn);
            var lastCol = _provider.GetLastColumn(fn, sn);
            for (int r = 1; r <= lastRow; r++)
                for (int c = 1; c <= lastCol; c++)
                    _provider.SetCellValue(newFn, sn, r, c, _provider.GetCellValue(fn, sn, r, c));
            _provider.SaveWorkbook(newFn);
            exported++;
        }
        return SkillResult.Ok($"导出完成，共导出 {exported} 个工作表到 {outputFolder}");
    }

    private SkillResult DeleteSheets(string fn, Dictionary<string, object> args)
    {
        var sheetNamesJson = GetStr(args, "sheetNames");
        var names = System.Text.Json.JsonSerializer.Deserialize<List<string>>(sheetNamesJson)!;
        foreach (var name in names) _provider.DeleteWorksheet(fn, name);
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"已删除 {names.Count} 个工作表: {string.Join(", ", names)}");
    }

    private SkillResult TransposeColumns(string fn, string sn, Dictionary<string, object> args)
    {
        var startCol = GetInt(args, "startColumn", 2);
        var fieldName = GetStr(args, "fieldName");
        var lastRow = _provider.GetLastRow(fn, sn);
        var lastCol = _provider.GetLastColumn(fn, sn);
        var outputSn = $"{sn}_转置";
        _provider.CreateWorksheet(fn, outputSn);
        int destRow = 1;
        for (int r = 2; r <= lastRow; r++)
        {
            for (int c = startCol; c <= lastCol; c++)
            {
                var header = _provider.GetCellValue(fn, sn, 1, c)?.ToString() ?? "";
                var keyVal = _provider.GetCellValue(fn, sn, r, 1)?.ToString() ?? "";
                var dataVal = _provider.GetCellValue(fn, sn, r, c)?.ToString() ?? "";
                for (int fc = 1; fc < startCol; fc++)
                    _provider.SetCellValue(fn, outputSn, destRow, fc, _provider.GetCellValue(fn, sn, r, fc));
                _provider.SetCellValue(fn, outputSn, destRow, startCol, header);
                _provider.SetCellValue(fn, outputSn, destRow, startCol + 1, dataVal);
                destRow++;
            }
        }
        _provider.SetCellValue(fn, outputSn, 1, startCol, fieldName);
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"转置完成，结果在工作表 '{outputSn}'");
    }

    private SkillResult CreateMultipleSheets(string fn, Dictionary<string, object> args)
    {
        var count = Math.Min(GetInt(args, "count", 1), 15);
        var baseName = GetStr(args, "baseName") ?? "新建表";
        for (int i = 1; i <= count; i++)
            _provider.CreateWorksheet(fn, $"{baseName}{i}");
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"已创建 {count} 个工作表");
    }

    private SkillResult GeneratePayslips(string fn, string sn, Dictionary<string, object> args)
    {
        var outputSn = GetStr(args, "outputSheetName") ?? "工资条";
        var lastRow = _provider.GetLastRow(fn, sn);
        var lastCol = _provider.GetLastColumn(fn, sn);
        _provider.CreateWorksheet(fn, outputSn);
        int destRow = 1;
        for (int r = 2; r <= lastRow; r++)
        {
            for (int c = 1; c <= lastCol; c++)
                _provider.SetCellValue(fn, outputSn, destRow, c, _provider.GetCellValue(fn, sn, 1, c));
            destRow++;
            for (int c = 1; c <= lastCol; c++)
                _provider.SetCellValue(fn, outputSn, destRow, c, _provider.GetCellValue(fn, sn, r, c));
            destRow++;
        }
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"工资条生成完成，共 {lastRow - 1} 条，在工作表 '{outputSn}'");
    }

    private int FindColumn(string fn, string sn, string columnName, int lastCol)
    {
        if (int.TryParse(columnName, out int colNum)) return colNum;
        for (int c = 1; c <= lastCol; c++)
            if (_provider.GetCellValue(fn, sn, 1, c)?.ToString() == columnName) return c;
        return 0;
    }

    private static Dictionary<string, object> ParamDef(string[] required, string[] optional)
    {
        var props = new Dictionary<string, object>();
        foreach (var r in required) props[r] = new { type = "string", description = $"{r}（必需）" };
        foreach (var o in optional) props[o] = new { type = "string", description = $"{o}（可选）" };
        return new Dictionary<string, object> { { "type", "object" }, { "properties", props } };
    }
    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
    private static int GetInt(Dictionary<string, object> a, string k, int d = 0) => a.ContainsKey(k) && int.TryParse(a[k]?.ToString(), out var v) ? v : d;
    private static bool GetBool(Dictionary<string, object> a, string k, bool d = false) => a.ContainsKey(k) && bool.TryParse(a[k]?.ToString(), out var v) ? v : d;
}