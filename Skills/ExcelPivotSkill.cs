using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TableMagic.Skills
{
    public class ExcelPivotSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public ExcelPivotSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "ExcelPivot";
        public string Description => "Excel数据透视表技能，提供数据透视表创建和管理功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "create_pivot_table",
                    Description = "创建数据透视表",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "sourceRange", new { type = "string", description = "数据源范围" } },
                                { "pivotSheetName", new { type = "string", description = "数据透视表工作表名称" } },
                                { "rowFields", new { type = "string", description = "行字段（JSON数组，可选）" } },
                                { "columnFields", new { type = "string", description = "列字段（JSON数组，可选）" } },
                                { "valueFields", new { type = "string", description = "值字段（JSON对象，可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sourceRange", "pivotSheetName" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "create_pivot_table":
                        {
                            var sourceRange = arguments["sourceRange"].ToString();
                            var pivotSheetName = arguments["pivotSheetName"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                            // 解析字段参数为集合
                            List<string> rowFieldsList = null;
                            List<string> columnFieldsList = null;
                            Dictionary<string, string> valueFieldsDict = null;

                            try { rowFieldsList = System.Text.Json.JsonSerializer.Deserialize<List<string>>(arguments.ContainsKey("rowFields") ? arguments["rowFields"].ToString() : "[]"); } catch { }
                            try { columnFieldsList = System.Text.Json.JsonSerializer.Deserialize<List<string>>(arguments.ContainsKey("columnFields") ? arguments["columnFields"].ToString() : "[]"); } catch { }
                            try { valueFieldsDict = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string,string>>(arguments.ContainsKey("valueFields") ? arguments["valueFields"].ToString() : "{}"); } catch { }

                            _excelMcp.CreatePivotTable(fileName, sheetName, sourceRange, pivotSheetName, "A1", "PivotTable1", rowFieldsList, columnFieldsList, valueFieldsDict);
                            return new SkillResult { Success = true, Content = "创建数据透视表成功" };
                        }
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelPivotSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }
    }
}