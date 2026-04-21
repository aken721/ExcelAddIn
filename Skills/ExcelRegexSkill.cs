using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelRegexSkill : ISkill
    {
        public string Name => "ExcelRegex";
        public string Description => "正则表达式技能，从单元格内容中提取指定格式的内容";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "extract_by_regex",
                    Description = "使用正则表达式从指定列提取内容到新列。当用户要求提取数字、邮箱、电话、身份证等特定格式内容时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "columnName", new { type = "string", description = "要提取内容的列名" } },
                                { "patternType", new { type = "string", description = "预定义模式：number(数字)/english(英文)/chinese(中文)/url(网址)/idcard(身份证号)/email(邮箱)/phone(电话)/ip(IP地址)/custom(自定义)" } },
                                { "pattern", new { type = "string", description = "自定义正则表达式（patternType为custom时需要）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "columnName" }
                },
                new SkillTool
                {
                    Name = "get_regex_patterns",
                    Description = "获取预定义的正则表达式模式列表。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "validate_regex",
                    Description = "验证正则表达式是否有效。",
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
                switch (toolName)
                {
                    case "extract_by_regex":
                        return await ExtractByRegexAsync(arguments);
                    case "get_regex_patterns":
                        return GetRegexPatterns();
                    case "validate_regex":
                        return ValidateRegex(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelRegexSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private async Task<SkillResult> ExtractByRegexAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var columnName = arguments["columnName"].ToString();
                var patternType = arguments.ContainsKey("patternType") 
                    ? arguments["patternType"].ToString().ToLower() 
                    : "number";
                var customPattern = arguments.ContainsKey("pattern") 
                    ? arguments["pattern"].ToString() 
                    : null;
                var sheetName = arguments.ContainsKey("sheetName") 
                    ? arguments["sheetName"].ToString() 
                    : null;

                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) 
                    ? workbook.ActiveSheet 
                    : workbook.Worksheets[sheetName];

                var usedRange = sheet.UsedRange;
                int lastRow = usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;

                int colIndex = GetColumnIndex(sheet, columnName);
                if (colIndex == 0)
                    return new SkillResult { Success = false, Error = $"未找到列: {columnName}" };

                var pattern = GetPattern(patternType, customPattern);
                if (pattern == null)
                    return new SkillResult { Success = false, Error = "无效的正则表达式模式" };

                ThisAddIn.app.ScreenUpdating = false;

                var regex = new Regex(pattern);
                int matchCount = 0;

                for (int r = 2; r <= lastRow; r++)
                {
                    var cellValue = sheet.Cells[r, colIndex].Text?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        var matches = regex.Matches(cellValue);
                        if (matches.Count > 0)
                        {
                            var matchList = new List<string>();
                            foreach (System.Text.RegularExpressions.Match m in matches)
                            {
                                matchList.Add(m.Value);
                            }
                            var result = string.Join("|", matchList);
                            sheet.Cells[r, lastCol + 1].Value = result;
                            matchCount++;
                        }
                    }
                }

                sheet.Cells[1, lastCol + 1].Value = $"{columnName}_提取结果";
                sheet.Columns[lastCol + 1].AutoFit();

                ThisAddIn.app.ScreenUpdating = true;

                return new SkillResult 
                { 
                    Success = true, 
                    Content = $"提取完成，共在 {matchCount} 行中找到匹配内容，结果已写入第 {lastCol + 1} 列" 
                };
            });
        }

        private SkillResult GetRegexPatterns()
        {
            var patterns = new Dictionary<string, string>
            {
                { "number", "数字 - \\d+\\.?\\d*" },
                { "english", "英文 - [A-Za-z]+" },
                { "chinese", "中文 - [^\\x00-\\xff]+" },
                { "url", "网址 - ((http|https):\\/\\/)?[\\w-]+(\\.[\\w-]+)+([\\w.,@?^=%&amp;:/~+#-]*[\\w@?^=%&amp;/~+#-])?" },
                { "idcard", "身份证号 - \\d{15}$|\\d{17}([0-9]|X|x)" },
                { "email", "电子邮箱 - \\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Z|a-z]{2,}\\b" },
                { "phone", "电话号码 - (?:(?:\\+|00)86)?1[3-9]\\d{9}|(?:0[1-9]\\d{1,2}-)?\\d{7,8}" },
                { "ip", "IP地址 - \\b\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\b" }
            };

            var content = "预定义正则表达式模式：\n" + string.Join("\n", patterns.Select(p => $"  - {p.Value}"));
            return new SkillResult { Success = true, Content = content };
        }

        private SkillResult ValidateRegex(Dictionary<string, object> arguments)
        {
            var pattern = arguments["pattern"].ToString();

            try
            {
                Regex.IsMatch("", pattern);
                return new SkillResult { Success = true, Content = "正则表达式有效" };
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = $"正则表达式无效：{ex.Message}" };
            }
        }

        private string GetPattern(string patternType, string customPattern)
        {
            return patternType switch
            {
                "number" => @"\d+\.?\d*",
                "english" => @"[A-Za-z]+",
                "chinese" => @"[^\x00-\xff]+",
                "url" => @"((http|https):\/\/)?[\w-]+(\.[\w-]+)+([\w.,@?^=%&amp;:/~+#-]*[\w@?^=%&amp;/~+#-])?",
                "idcard" => @"\d{15}$|\d{17}([0-9]|X|x)",
                "email" => @"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b",
                "phone" => @"(?:(?:\+|00)86)?1[3-9]\d{9}|(?:0[1-9]\d{1,2}-)?\d{7,8}",
                "ip" => @"\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b",
                "custom" => customPattern,
                _ => customPattern
            };
        }

        private int GetColumnIndex(Excel.Worksheet sheet, string columnName)
        {
            if (int.TryParse(columnName, out int colNum))
                return colNum;

            var usedRange = sheet.UsedRange;
            for (int c = 1; c <= usedRange.Columns.Count; c++)
            {
                if (sheet.Cells[1, c].Text?.ToString() == columnName)
                    return c;
            }
            return 0;
        }
    }
}
