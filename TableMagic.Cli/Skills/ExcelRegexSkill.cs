using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelRegexSkill : ISkill
{
    private readonly IExcelProvider _provider;
    public ExcelRegexSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelRegex";
    public string Description => "正则表达式技能，从单元格内容中提取指定格式的内容";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "extract_by_regex",
                Description = "使用正则表达式从指定列提取内容到新列。支持预定义模式：number/english/chinese/url/idcard/email/phone/ip/custom",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                            { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                            { "columnName", new { type = "string", description = "要提取内容的列名" } },
                            { "patternType", new { type = "string", description = "预定义模式：number/english/chinese/url/idcard/email/phone/ip/custom" } },
                            { "pattern", new { type = "string", description = "自定义正则表达式（patternType为custom时需要）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "columnName" }
            },
            new()
            {
                Name = "get_regex_patterns",
                Description = "获取预定义的正则表达式模式列表",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>() }
                },
                RequiredParameters = new List<string>()
            },
            new()
            {
                Name = "validate_regex",
                Description = "验证正则表达式是否有效",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "pattern", new { type = "string", description = "要验证的正则表达式" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "pattern" }
            }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "extract_by_regex" => ExtractByRegex(arguments),
                "get_regex_patterns" => GetRegexPatterns(),
                "validate_regex" => ValidateRegex(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult ExtractByRegex(Dictionary<string, object> args)
    {
        var fn = GetStr(args, "fileName");
        var sn = GetStr(args, "sheetName");
        var columnName = GetStr(args, "columnName");
        var patternType = GetStr(args, "patternType") ?? "number";
        var customPattern = GetStr(args, "pattern");

        var pattern = GetPattern(patternType, customPattern);
        if (pattern == null) return new SkillResult { Success = false, Error = "无效的正则表达式模式" };

        var regex = new Regex(pattern);
        var lastRow = _provider.GetLastRow(fn, sn);
        var lastCol = _provider.GetLastColumn(fn, sn);
        var usedRange = _provider.GetUsedRange(fn, sn);

        var colIndex = FindColumnIndex(fn, sn, columnName, lastCol);
        if (colIndex == 0) return new SkillResult { Success = false, Error = $"未找到列: {columnName}" };

        int matchCount = 0;
        for (int r = 2; r <= lastRow; r++)
        {
            var cellValue = _provider.GetCellValue(fn, sn, r, colIndex)?.ToString() ?? "";
            if (!string.IsNullOrEmpty(cellValue))
            {
                var matches = regex.Matches(cellValue);
                if (matches.Count > 0)
                {
                    var result = string.Join("|", matches.Cast<Match>().Select(m => m.Value));
                    _provider.SetCellValue(fn, sn, r, lastCol + 1, result);
                    matchCount++;
                }
            }
        }

        _provider.SetCellValue(fn, sn, 1, lastCol + 1, $"{columnName}_提取结果");
        _provider.SaveWorkbook(fn);
        return SkillResult.Ok($"提取完成，共在 {matchCount} 行中找到匹配内容，结果已写入第 {lastCol + 1} 列");
    }

    private SkillResult GetRegexPatterns()
    {
        var patterns = new Dictionary<string, string>
        {
            { "number", "数字 - \\d+\\.?\\d*" },
            { "english", "英文 - [A-Za-z]+" },
            { "chinese", "中文 - [^\\x00-\\xff]+" },
            { "url", "网址 - ((http|https):\\/\\/)?[\\w-]+(\\.[\\w-]+)+.*" },
            { "idcard", "身份证号 - \\d{15}$|\\d{17}([0-9]|X|x)" },
            { "email", "电子邮箱 - \\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Z|a-z]{2,}\\b" },
            { "phone", "电话号码 - (?:(?:\\+|00)86)?1[3-9]\\d{9}|(?:0[1-9]\\d{1,2}-)?\\d{7,8}" },
            { "ip", "IP地址 - \\b\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\b" }
        };
        return SkillResult.Ok("预定义正则表达式模式：\n" + string.Join("\n", patterns.Select(p => $"  - {p.Value}")));
    }

    private SkillResult ValidateRegex(Dictionary<string, object> args)
    {
        var pattern = GetStr(args, "pattern");
        try { Regex.IsMatch("", pattern); return SkillResult.Ok("正则表达式有效"); }
        catch (Exception ex) { return new SkillResult { Success = false, Error = $"正则表达式无效：{ex.Message}" }; }
    }

    private int FindColumnIndex(string fn, string sn, string columnName, int lastCol)
    {
        if (int.TryParse(columnName, out int colNum)) return colNum;
        for (int c = 1; c <= lastCol; c++)
        {
            var header = _provider.GetCellValue(fn, sn, 1, c)?.ToString();
            if (header == columnName) return c;
        }
        return 0;
    }

    private static string GetPattern(string patternType, string customPattern) => patternType switch
    {
        "number" => @"\d+\.?\d*",
        "english" => @"[A-Za-z]+",
        "chinese" => @"[^\x00-\xff]+",
        "url" => @"((http|https):\/\/)?[\w-]+(\.[\w-]+)+.*",
        "idcard" => @"\d{15}$|\d{17}([0-9]|X|x)",
        "email" => @"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b",
        "phone" => @"(?:(?:\+|00)86)?1[3-9]\d{9}|(?:0[1-9]\d{1,2}-)?\d{7,8}",
        "ip" => @"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b",
        "custom" => customPattern,
        _ => customPattern
    };

    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
}