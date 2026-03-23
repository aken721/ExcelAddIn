using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Text.Json;

namespace ExcelAddIn.Skills
{
    public class ExcelCellSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public ExcelCellSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "ExcelCell";
        public string Description => "Excel单元格技能，提供单元格值的读取、写入、公式设置等功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "set_cell_value",
                    Description = "设置单元格的值。当用户要求写入数据、填写单元格、设置单元格内容时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "row", new { type = "integer", description = "行号（从1开始）" } },
                                { "column", new { type = "integer", description = "列号（从1开始）" } },
                                { "value", new { type = "string", description = "要设置的值（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "row", "column", "value" }
                },
                new SkillTool
                {
                    Name = "get_cell_value",
                    Description = "获取单元格的值。当用户要求读取数据、查看单元格内容、获取单元格值时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "row", new { type = "integer", description = "行号（从1开始）" } },
                                { "column", new { type = "integer", description = "列号（从1开始）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "row", "column" }
                },
                new SkillTool
                {
                    Name = "set_cell_formula",
                    Description = "设置单元格的公式。当用户要求设置公式、添加计算公式、使用函数时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "cellAddress", new { type = "string", description = "单元格地址，如A1" } },
                                { "formula", new { type = "string", description = "公式，如=SUM(A1:A10)（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "cellAddress", "formula" }
                },
                new SkillTool
                {
                    Name = "get_cell_formula",
                    Description = "获取单元格的公式。当用户要求查看公式、获取公式内容时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "cellAddress", new { type = "string", description = "单元格地址，如A1（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "cellAddress" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                string fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                string sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                
                var workbook = GetWorkbook(fileName);
                var worksheet = GetWorksheet(workbook, sheetName);

                switch (toolName)
                {
                    case "set_cell_value":
                        {
                            var row = Convert.ToInt32(arguments["row"]);
                            var column = Convert.ToInt32(arguments["column"]);
                            var value = arguments["value"];
                            
                            object cellValue;
                            if (value is JsonElement je)
                            {
                                if (je.ValueKind == JsonValueKind.Number)
                                    cellValue = je.GetDouble();
                                else
                                    cellValue = je.ToString();
                            }
                            else
                            {
                                cellValue = value.ToString();
                            }
                            
                            worksheet.Cells[row, column].Value = cellValue;
                            return new SkillResult { Success = true, Content = $"成功设置单元格 ({row},{column}) 的值为: {cellValue}" };
                        }
                    case "get_cell_value":
                        {
                            var row = Convert.ToInt32(arguments["row"]);
                            var column = Convert.ToInt32(arguments["column"]);
                            var cellValue = worksheet.Cells[row, column].Value?.ToString() ?? "";
                            return new SkillResult { Success = true, Content = $"单元格 ({row},{column}) 的值为: {cellValue}" };
                        }
                    case "set_cell_formula":
                        {
                            var cellAddress = arguments["cellAddress"].ToString();
                            var formula = arguments["formula"].ToString();
                            worksheet.Range[cellAddress].Formula = formula;
                            return new SkillResult { Success = true, Content = $"成功设置单元格 {cellAddress} 的公式为: {formula}" };
                        }
                    case "get_cell_formula":
                        {
                            var cellAddress = arguments["cellAddress"].ToString();
                            var formula = worksheet.Range[cellAddress].Formula?.ToString() ?? "";
                            return new SkillResult { Success = true, Content = $"单元格 {cellAddress} 的公式为: {formula}" };
                        }
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelCellSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private Microsoft.Office.Interop.Excel.Workbook GetWorkbook(string fileName)
        {
            if (ThisAddIn.app == null)
                throw new Exception("Excel应用程序未初始化");

            if (string.IsNullOrEmpty(fileName))
            {
                if (ThisAddIn.app.ActiveWorkbook != null)
                    return ThisAddIn.app.ActiveWorkbook;
                throw new Exception("未指定工作簿且没有活跃工作簿");
            }

            foreach (Microsoft.Office.Interop.Excel.Workbook wb in ThisAddIn.app.Workbooks)
            {
                if (wb.Name == fileName)
                    return wb;
            }

            throw new Exception($"未找到工作簿: {fileName}");
        }

        private Microsoft.Office.Interop.Excel.Worksheet GetWorksheet(Microsoft.Office.Interop.Excel.Workbook workbook, string sheetName)
        {
            string targetSheetName = sheetName;
            
            if (string.IsNullOrEmpty(targetSheetName))
            {
                if (workbook.ActiveSheet != null)
                    return workbook.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;
                throw new Exception("未指定工作表且没有活跃工作表");
            }

            foreach (Microsoft.Office.Interop.Excel.Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == targetSheetName)
                    return ws;
            }

            throw new Exception($"未找到工作表: {targetSheetName}");
        }
    }
}

