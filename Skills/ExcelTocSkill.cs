using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelTocSkill : ISkill
    {
        public string Name => "ExcelToc";
        public string Description => "目录页技能，根据目录表批量创建工作表并添加超链接";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "create_sheets_from_toc",
                    Description = "根据目录表批量创建工作表并添加超链接。目录表需命名为'目录'。当用户要求根据目录创建表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "linkColumnName", new { type = "string", description = "包含工作表名称的列名" } },
                                { "createSheets", new { type = "boolean", description = "是否创建工作表（默认true）" } },
                                { "addHyperlinks", new { type = "boolean", description = "是否添加超链接（默认true）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "linkColumnName" }
                },
                new SkillTool
                {
                    Name = "create_toc_sheet",
                    Description = "创建目录表，列出当前工作簿中所有工作表名称。当用户要求创建目录页、生成目录时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "tocSheetName", new { type = "string", description = "目录表名称（默认'_目录'）" } },
                                { "addHyperlinks", new { type = "boolean", description = "是否添加超链接（默认true）" } },
                                { "includeHiddenSheets", new { type = "boolean", description = "是否包含隐藏的工作表（默认false）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "update_toc_hyperlinks",
                    Description = "更新目录表中的超链接。当用户要求更新目录链接时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "columnName", new { type = "string", description = "包含工作表名称的列名" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "columnName" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "create_sheets_from_toc":
                        return await CreateSheetsFromTocAsync(arguments);
                    case "create_toc_sheet":
                        return await CreateTocSheetAsync(arguments);
                    case "update_toc_hyperlinks":
                        return await UpdateTocHyperlinksAsync(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelTocSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private Task<SkillResult> CreateSheetsFromTocAsync(Dictionary<string, object> arguments)
        {
            var linkColumnName = arguments["linkColumnName"].ToString();
            var createSheets = !arguments.ContainsKey("createSheets") || Convert.ToBoolean(arguments["createSheets"]);
            var addHyperlinks = !arguments.ContainsKey("addHyperlinks") || Convert.ToBoolean(arguments["addHyperlinks"]);

            var workbook = ThisAddIn.app.ActiveWorkbook;

            Excel.Worksheet tocSheet = null;
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == "目录")
                {
                    tocSheet = sheet;
                    break;
                }
            }

            if (tocSheet == null)
                return Task.FromResult(new SkillResult { Success = false, Error = "未找到名为'目录'的工作表" });

            int colIndex = GetColumnIndex(tocSheet, linkColumnName);
            if (colIndex == 0)
                return Task.FromResult(new SkillResult { Success = false, Error = $"未找到列: {linkColumnName}" });

            var existingSheets = new List<string>();
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
                existingSheets.Add(sheet.Name);

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            int lastRow = tocSheet.UsedRange.Rows.Count;
            int createdCount = 0;
            int linkedCount = 0;

            for (int r = 2; r <= lastRow; r++)
            {
                var sheetName = tocSheet.Cells[r, colIndex].Text?.ToString();
                if (string.IsNullOrEmpty(sheetName)) continue;

                if (createSheets && !existingSheets.Any(s => string.Equals(s, sheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    try
                    {
                        var newSheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                        newSheet.Name = sheetName;
                        existingSheets.Add(sheetName);
                        createdCount++;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"创建工作表 '{sheetName}' 失败: {ex.Message}");
                    }
                }

                if (addHyperlinks)
                {
                    try
                    {
                        tocSheet.Hyperlinks.Add(
                            tocSheet.Cells[r, colIndex],
                            "",
                            $"'{sheetName}'!A1",
                            sheetName,
                            sheetName
                        );
                        tocSheet.Cells[r, colIndex].Font.Name = "微软雅黑";
                        tocSheet.Cells[r, colIndex].Font.Size = 12;
                        tocSheet.Cells[r, colIndex].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        tocSheet.Cells[r, colIndex].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                        linkedCount++;
                    }
                    catch { }
                }
            }

            tocSheet.Activate();
            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;

            return Task.FromResult(new SkillResult 
            { 
                Success = true, 
                Content = $"目录页处理完成：创建 {createdCount} 个工作表，添加 {linkedCount} 个超链接" 
            });
        }

        private Task<SkillResult> CreateTocSheetAsync(Dictionary<string, object> arguments)
        {
            var tocSheetName = arguments.ContainsKey("tocSheetName") 
                ? arguments["tocSheetName"].ToString() 
                : "_目录";
            var addHyperlinks = !arguments.ContainsKey("addHyperlinks") || Convert.ToBoolean(arguments["addHyperlinks"]);
            var includeHiddenSheets = arguments.ContainsKey("includeHiddenSheets") && Convert.ToBoolean(arguments["includeHiddenSheets"]);

            var workbook = ThisAddIn.app.ActiveWorkbook;

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            var existingSheets = new List<string>();
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
                existingSheets.Add(sheet.Name);

            if (existingSheets.Any(s => string.Equals(s, tocSheetName, StringComparison.OrdinalIgnoreCase)))
            {
                var oldSheet = workbook.Worksheets[tocSheetName];
                oldSheet.Name = tocSheetName + DateTime.Now.ToString("yyyyMMddHHmmss");
            }

            var tocSheet = workbook.Worksheets.Add(Before: workbook.Worksheets[1]);
            tocSheet.Name = tocSheetName;
            tocSheet.Cells[1, 1].Value = "表目录";
            tocSheet.Cells[1, 1].Font.Bold = true;

            int row = 2;
            int visibleCount = 0;
            int hiddenCount = 0;
            var hiddenSheets = new List<string>();

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == tocSheetName) continue;
                
                bool isHidden = sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible;
                if (isHidden)
                {
                    hiddenCount++;
                    hiddenSheets.Add(sheet.Name);
                }
                
                if (!includeHiddenSheets && isHidden) continue;

                tocSheet.Cells[row, 1].Value = sheet.Name;

                if (addHyperlinks)
                {
                    tocSheet.Hyperlinks.Add(
                        tocSheet.Cells[row, 1],
                        "",
                        $"'{sheet.Name}'!A1",
                        sheet.Name,
                        sheet.Name
                    );
                }

                tocSheet.Cells[row, 1].Font.Name = "微软雅黑";
                tocSheet.Cells[row, 1].Font.Size = 12;
                tocSheet.Cells[row, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                tocSheet.Cells[row, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                row++;
                visibleCount++;
            }

            tocSheet.UsedRange.Columns.AutoFit();
            tocSheet.Activate();

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;

            var contentBuilder = new System.Text.StringBuilder();
            contentBuilder.AppendLine($"目录表创建完成，共列出 {row - 2} 个工作表");
            
            if (hiddenCount > 0 && !includeHiddenSheets)
            {
                contentBuilder.AppendLine($"⚠️ 检测到 {hiddenCount} 个隐藏工作表未列入目录：{string.Join("、", hiddenSheets)}");
                contentBuilder.AppendLine("如需包含隐藏工作表，请重新执行并设置 includeHiddenSheets=true");
            }

            return Task.FromResult(new SkillResult 
            { 
                Success = true, 
                Content = contentBuilder.ToString().TrimEnd()
            });
        }

        private async Task<SkillResult> UpdateTocHyperlinksAsync(Dictionary<string, object> arguments)
        {
            return await CreateSheetsFromTocAsync(new Dictionary<string, object>
            {
                { "linkColumnName", arguments["columnName"].ToString() },
                { "createSheets", false },
                { "addHyperlinks", true }
            });
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
