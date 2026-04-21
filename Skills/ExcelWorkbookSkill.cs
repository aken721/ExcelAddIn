using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace TableMagic.Skills
{
    public class ExcelWorkbookSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public ExcelWorkbookSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "ExcelWorkbook";
        public string Description => "Excel工作簿技能，提供工作簿创建、打开、保存、关闭等功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "create_workbook",
                    Description = "创建新的Excel工作簿文件。当用户要求新建文件、创建工作簿、新建Excel时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "新工作簿文件名（必需）" } },
                                { "sheetName", new { type = "string", description = "初始工作表名称（可选，默认Sheet1）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "fileName" }
                },
                new SkillTool
                {
                    Name = "open_workbook",
                    Description = "打开工作簿文件。当用户要求打开文件、打开Excel文件时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "要打开的工作簿文件名（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "fileName" }
                },
                new SkillTool
                {
                    Name = "close_workbook",
                    Description = "关闭工作簿。当用户要求关闭文件、关闭Excel时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认关闭当前活跃工作簿）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "save_workbook",
                    Description = "保存工作簿。当用户要求保存文件、保存Excel时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认保存当前活跃工作簿）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "save_workbook_as",
                    Description = "将工作簿另存为新文件。当用户要求另存为、保存副本时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "原工作簿文件名（可选）" } },
                                { "newFileName", new { type = "string", description = "新文件名（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "newFileName" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "create_workbook":
                        {
                            var fileName = arguments["fileName"].ToString();
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : "Sheet1";
                            var result = _excelMcp.CreateWorkbook(fileName, sheetName);
                            return new SkillResult { Success = true, Content = $"成功创建工作簿文件: {result}（保存在excel_files目录）" };
                        }
                    case "open_workbook":
                        {
                            var fileName = arguments["fileName"].ToString();
                            var filePath = System.IO.Path.Combine("./excel_files", fileName);
                            
                            if (!System.IO.File.Exists(filePath))
                                return new SkillResult { Success = false, Error = $"文件不存在: {filePath}" };

                            var wb = ThisAddIn.app.Workbooks.Open(filePath);
                            return new SkillResult { Success = true, Content = $"成功打开工作簿: {fileName}" };
                        }
                    case "close_workbook":
                        {
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            if (string.IsNullOrEmpty(fileName))
                            {
                                if (ThisAddIn.app.ActiveWorkbook != null)
                                {
                                    ThisAddIn.app.ActiveWorkbook.Close(true);
                                    return new SkillResult { Success = true, Content = "成功关闭当前活跃工作簿" };
                                }
                                return new SkillResult { Success = false, Error = "没有活跃的工作簿" };
                            }
                            
                            foreach (Microsoft.Office.Interop.Excel.Workbook wb in ThisAddIn.app.Workbooks)
                            {
                                if (wb.Name == fileName)
                                {
                                    wb.Close(true);
                                    return new SkillResult { Success = true, Content = $"成功关闭工作簿: {fileName}" };
                                }
                            }
                            return new SkillResult { Success = false, Error = $"未找到工作簿: {fileName}" };
                        }
                    case "save_workbook":
                        {
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            Microsoft.Office.Interop.Excel.Workbook wb = null;
                            
                            if (string.IsNullOrEmpty(fileName))
                            {
                                wb = ThisAddIn.app.ActiveWorkbook;
                                if (wb == null)
                                    return new SkillResult { Success = false, Error = "没有活跃的工作簿" };
                            }
                            else
                            {
                                foreach (Microsoft.Office.Interop.Excel.Workbook workbook in ThisAddIn.app.Workbooks)
                                {
                                    if (workbook.Name == fileName)
                                    {
                                        wb = workbook;
                                        break;
                                    }
                                }
                                if (wb == null)
                                    return new SkillResult { Success = false, Error = $"未找到工作簿: {fileName}" };
                            }
                            
                            wb.Save();
                            return new SkillResult { Success = true, Content = $"成功保存工作簿: {wb.Name}" };
                        }
                    case "save_workbook_as":
                        {
                            var newFileName = arguments["newFileName"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            Microsoft.Office.Interop.Excel.Workbook wb = null;
                            
                            if (string.IsNullOrEmpty(fileName))
                            {
                                wb = ThisAddIn.app.ActiveWorkbook;
                                if (wb == null)
                                    return new SkillResult { Success = false, Error = "没有活跃的工作簿" };
                            }
                            else
                            {
                                foreach (Microsoft.Office.Interop.Excel.Workbook workbook in ThisAddIn.app.Workbooks)
                                {
                                    if (workbook.Name == fileName)
                                    {
                                        wb = workbook;
                                        break;
                                    }
                                }
                                if (wb == null)
                                    return new SkillResult { Success = false, Error = $"未找到工作簿: {fileName}" };
                            }
                            
                            var newFilePath = System.IO.Path.Combine(
                                System.IO.Path.GetDirectoryName(wb.FullName), newFileName);
                            wb.SaveAs(newFilePath);
                            
                            return new SkillResult { Success = true, Content = $"成功将工作簿另存为: {newFileName}" };
                        }
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelWorkbookSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }
    }
}
