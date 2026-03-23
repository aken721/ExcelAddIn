using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelAddIn.Skills
{
    public class ExcelChartSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public ExcelChartSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "ExcelChart";
        public string Description => "Excel图表技能，提供图表创建和管理功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "create_chart",
                    Description = "创建图表。当用户要求创建图表、画图、生成柱状图、折线图、饼图、条形图、面积图、散点图时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "dataRange", new { type = "string", description = "数据范围（必需），如A1:D10" } },
                                { "chartType", new { type = "string", description = "图表类型（可选，默认column）：column(柱状图)/line(折线图)/pie(饼图)/bar(条形图)/area(面积图)/scatter(散点图)" } },
                                { "title", new { type = "string", description = "图表标题（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "dataRange" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "create_chart":
                        {
                            var dataRange = arguments["dataRange"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                            var chartType = arguments.ContainsKey("chartType") ? arguments["chartType"].ToString() : "column";
                            var title = arguments.ContainsKey("title") ? arguments["title"].ToString() : "";
                            var xAxisTitle = arguments.ContainsKey("xAxisTitle") ? arguments["xAxisTitle"].ToString() : "";
                            var yAxisTitle = arguments.ContainsKey("yAxisTitle") ? arguments["yAxisTitle"].ToString() : "";

                            // ExcelMcp.CreateChart signature expects chartPosition and numeric width/height.
                            // Use default position and sizes for compatibility with this wrapper.
                            _excelMcp.CreateChart(fileName, sheetName, chartType, dataRange, "A1", title, 400, 300);
                            return new SkillResult { Success = true, Content = "创建图表成功" };
                        }
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelChartSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }
    }
}