using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TableMagic.Skills
{
    public class ExcelRangeSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public ExcelRangeSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "ExcelRange";
        public string Description => "Excel范围操作技能，提供范围数据处理和公式功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "set_range_values",
                    Description = "批量设置单元格区域的值。当用户要求批量写入数据、填充区域、设置多行多列数据时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "rangeAddress", new { type = "string", description = "范围地址（必需），如A1:D10" } },
                                { "data", new { type = "string", description = "数据（必需，JSON格式的二维数组）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "rangeAddress", "data" }
                },
                new SkillTool
                {
                    Name = "get_range_values",
                    Description = "获取单元格区域的值。当用户要求读取区域数据、获取多行多列数据、查看范围内容时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "rangeAddress", new { type = "string", description = "范围地址（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "rangeAddress" }
                },
                new SkillTool
                {
                    Name = "set_formula",
                    Description = "设置单元格公式。当用户要求设置公式、添加计算公式、使用函数（如SUM、AVERAGE等）时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "cellAddress", new { type = "string", description = "单元格地址（必需），如A1" } },
                                { "formula", new { type = "string", description = "公式（必需），如=SUM(A1:A10)" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "cellAddress", "formula" }
                },
                new SkillTool
                {
                    Name = "get_formula",
                    Description = "获取单元格公式。当用户要求查看公式、获取公式内容时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "cellAddress", new { type = "string", description = "单元格地址（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "cellAddress" }
                },
                new SkillTool
                {
                    Name = "copy_range",
                    Description = "复制单元格区域。当用户要求复制数据、复制区域、复制粘贴时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "sourceRange", new { type = "string", description = "源范围地址（必需）" } },
                                { "targetRange", new { type = "string", description = "目标范围地址（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sourceRange", "targetRange" }
                },
                new SkillTool
                {
                    Name = "clear_range",
                    Description = "清除单元格区域。当用户要求清空数据、删除内容、清除格式时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "rangeAddress", new { type = "string", description = "范围地址（必需）" } },
                                { "clearType", new { type = "string", description = "清除类型（可选，默认all）：all/values/formats/comments" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "rangeAddress" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "set_range_values":
                        {
                            var rangeAddress = arguments["rangeAddress"].ToString();
                            var dataStr = arguments["data"].ToString();
                            // 默认期望 JSON 格式的二维数组字符串，尝试解析为 object[,]
                            object[,] data = null;
                            try
                            {
                                var options = new System.Text.Json.JsonSerializerOptions();
                                var parsed = System.Text.Json.JsonSerializer.Deserialize<object[][]>(dataStr, options);
                                if (parsed != null)
                                {
                                    int r = parsed.Length;
                                    int c = parsed[0].Length;
                                    data = new object[r, c];
                                    for (int i = 0; i < r; i++)
                                        for (int j = 0; j < c; j++)
                                            data[i, j] = parsed[i][j];
                                }
                            }
                            catch
                            {
                                // ignore parse error, leave data as null
                            }
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            _excelMcp.SetRangeValues(fileName, sheetName, rangeAddress, data);
                            return new SkillResult { Success = true, Content = "设置范围值成功" };
                        }
                    case "get_range_values":
                        {
                            var rangeAddress = arguments["rangeAddress"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            var result = _excelMcp.GetRangeValues(fileName, sheetName, rangeAddress);
                            return new SkillResult { Success = true, Content = ConvertRangeToString(result) };
                        }
                    case "set_formula":
                        {
                            var cellAddress = arguments["cellAddress"].ToString();
                            var formula = arguments["formula"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            _excelMcp.SetFormula(fileName, sheetName, cellAddress, formula);
                            return new SkillResult { Success = true, Content = "设置公式成功" };
                        }
                    case "get_formula":
                        {
                            var cellAddress = arguments["cellAddress"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            var formula = _excelMcp.GetFormula(fileName, sheetName, cellAddress);
                            return new SkillResult { Success = true, Content = formula };
                        }
                    case "copy_range":
                        {
                            var sourceRange = arguments["sourceRange"].ToString();
                            var targetRange = arguments["targetRange"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            _excelMcp.CopyRange(fileName, sheetName, sourceRange, targetRange);
                            return new SkillResult { Success = true, Content = "复制范围成功" };
                        }
                    case "clear_range":
                        {
                            var rangeAddress = arguments["rangeAddress"].ToString();
                            var clearType = arguments.ContainsKey("clearType") ? arguments["clearType"].ToString() : "all";
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            _excelMcp.ClearRange(fileName, sheetName, rangeAddress, clearType);
                            return new SkillResult { Success = true, Content = "清除范围成功" };
                        }
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelRangeSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private string ConvertRangeToString(object[,] data)
        {
            if (data == null)
                return "[]";

            var sb = new System.Text.StringBuilder();
            sb.Append("[");

            for (int i = 0; i < data.GetLength(0); i++)
            {
                sb.Append("[");
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    var value = data[i, j];
                    if (value is string s)
                        sb.Append($"\"{s}\"");
                    else
                        sb.Append(value?.ToString() ?? "null");

                    if (j < data.GetLength(1) - 1)
                        sb.Append(", ");
                }
                sb.Append("]");
                if (i < data.GetLength(0) - 1)
                    sb.Append(", ");
            }

            sb.Append("]");
            return sb.ToString();
        }
    }
}