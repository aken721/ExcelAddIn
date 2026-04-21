using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TableMagic.Skills
{
    public class ExcelAnalysisSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public ExcelAnalysisSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "ExcelAnalysis";
        public string Description => "Excel数据分析技能，提供数据统计、分析等功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "analyze_data",
                    Description = "分析指定范围的数据",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "range", new { type = "string", description = "数据范围" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "range" }
                },
                new SkillTool
                {
                    Name = "get_range_statistics",
                    Description = "获取指定范围的统计信息",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "range", new { type = "string", description = "数据范围" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "range" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "analyze_data":
                        {
                            var range = arguments["range"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            var data = _excelMcp.GetRangeValues(fileName, sheetName, range);
                            var analysis = AnalyzeData(data);
                            return new SkillResult { Success = true, Content = analysis };
                        }
                    case "get_range_statistics":
                        {
                            var range = arguments["range"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            var stats = _excelMcp.GetRangeStatistics(fileName, sheetName, range);
                            return new SkillResult { Success = true, Content = stats?.ToString() ?? string.Empty };
                        }
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelAnalysisSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private string AnalyzeData(object[,] data)
        {
            if (data == null || data.GetLength(0) == 0 || data.GetLength(1) == 0)
            {
                return "数据为空";
            }

            int rows = data.GetLength(0);
            int cols = data.GetLength(1);
            int numericCount = 0;
            double sum = 0;

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    if (data[i, j] is double d)
                    {
                        sum += d;
                        numericCount++;
                    }
                    else if (data[i, j] is int n)
                    {
                        sum += n;
                        numericCount++;
                    }
                }
            }

            var analysis = new System.Text.StringBuilder();
            analysis.AppendLine($"数据范围: {rows}行 × {cols}列");
            analysis.AppendLine($"数值单元格数量: {numericCount}");
            if (numericCount > 0)
            {
                analysis.AppendLine($"总和: {sum}");
                analysis.AppendLine($"平均值: {sum / numericCount}");
            }

            return analysis.ToString();
        }
    }
}