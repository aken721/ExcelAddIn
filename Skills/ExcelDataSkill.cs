using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelDataSkill : ISkill
    {
        public string Name => "ExcelData";
        public string Description => "Excel数据处理技能，提供分表、并表、批量导删、转置、工资条等功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "split_sheet_by_column",
                    Description = "根据指定列的值将工作表拆分为多个工作表。当用户要求按某列分表、拆分工作表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "columnName", new { type = "string", description = "按哪列分表（列名或列号）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选，默认当前表）" } },
                                { "dataStartRow", new { type = "integer", description = "数据起始行（默认2）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "columnName" }
                },
                new SkillTool
                {
                    Name = "split_and_export",
                    Description = "根据指定列分表并导出为独立文件。当用户要求分表并保存为单独文件时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "columnName", new { type = "string", description = "按哪列分表" } },
                                { "outputFolder", new { type = "string", description = "输出文件夹路径" } },
                                { "fileFormat", new { type = "string", description = "文件格式：xlsx/xls/csv，默认xlsx" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "columnName", "outputFolder" }
                },
                new SkillTool
                {
                    Name = "merge_sheets",
                    Description = "将多个工作表合并为一个工作表。当用户要求合并表、并表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "sheetNames", new { type = "string", description = "要合并的工作表名称列表（JSON数组格式，留空则合并所有）" } },
                                { "outputSheetName", new { type = "string", description = "合并后的表名（默认'合并表'）" } },
                                { "includeHeader", new { type = "boolean", description = "是否包含标题行（默认true）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "merge_workbooks",
                    Description = "将指定目录下所有工作簿的工作表合并到当前工作簿。当用户要求合并多个工作簿时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "folderPath", new { type = "string", description = "工作簿所在文件夹路径" } },
                                { "includeSubfolders", new { type = "boolean", description = "是否包含子目录（默认true）" } },
                                { "skipEmptySheets", new { type = "boolean", description = "是否跳过空表（默认true）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "folderPath" }
                },
                new SkillTool
                {
                    Name = "export_sheets",
                    Description = "批量导出工作表为独立文件。当用户要求导出工作表、批量导出时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "outputFolder", new { type = "string", description = "输出文件夹路径" } },
                                { "sheetNames", new { type = "string", description = "要导出的工作表名称列表（JSON数组格式，留空则导出所有）" } },
                                { "fileFormat", new { type = "string", description = "文件格式：xlsx/xls/csv，默认xlsx" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "outputFolder" }
                },
                new SkillTool
                {
                    Name = "delete_sheets",
                    Description = "批量删除工作表。当用户要求批量删除表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "sheetNames", new { type = "string", description = "要删除的工作表名称列表（JSON数组格式）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sheetNames" }
                },
                new SkillTool
                {
                    Name = "transpose_columns",
                    Description = "将列名称转置为字段内数据（宽表转长表）。当用户要求转置、宽表转长表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "startColumn", new { type = "integer", description = "从第几列开始转置（不小于2）" } },
                                { "fieldName", new { type = "string", description = "转置列的字段名称" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "startColumn", "fieldName" }
                },
                new SkillTool
                {
                    Name = "import_sheets_from_folder",
                    Description = "将指定文件夹下所有工作簿的工作表导入到当前工作簿。当用户要求多工作簿表转同工作簿时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "folderPath", new { type = "string", description = "工作簿所在文件夹路径" } },
                                { "includeSubfolders", new { type = "boolean", description = "是否包含子目录（默认true）" } },
                                { "skipEmptySheets", new { type = "boolean", description = "是否跳过空表（默认true）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "folderPath" }
                },
                new SkillTool
                {
                    Name = "create_multiple_sheets",
                    Description = "批量创建指定数量的工作表。当用户要求一键建立多个工作表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "count", new { type = "integer", description = "创建数量（最多15张）" } },
                                { "baseName", new { type = "string", description = "工作表基础名称（默认'新建表'）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "count" }
                },
                new SkillTool
                {
                    Name = "generate_payslips",
                    Description = "将工资表转换为工资条格式。在每一行数据前插入标题行，便于打印分发。当用户要求生成工资条时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "sheetName", new { type = "string", description = "工资表名称（可选，默认当前表）" } },
                                { "outputSheetName", new { type = "string", description = "工资条表名称（默认'工资条'）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "split_sheet_by_column":
                        return await SplitSheetByColumnAsync(arguments);
                    case "split_and_export":
                        return await SplitAndExportAsync(arguments);
                    case "merge_sheets":
                        return await MergeSheetsAsync(arguments);
                    case "merge_workbooks":
                        return await MergeWorkbooksAsync(arguments);
                    case "export_sheets":
                        return await ExportSheetsAsync(arguments);
                    case "delete_sheets":
                        return await DeleteSheetsAsync(arguments);
                    case "transpose_columns":
                        return await TransposeColumnsAsync(arguments);
                    case "import_sheets_from_folder":
                        return await ImportSheetsFromFolderAsync(arguments);
                    case "create_multiple_sheets":
                        return await CreateMultipleSheetsAsync(arguments);
                    case "generate_payslips":
                        return await GeneratePayslipsAsync(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelDataSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private Task<SkillResult> SplitSheetByColumnAsync(Dictionary<string, object> arguments)
        {
            var columnName = arguments["columnName"].ToString();
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
            var dataStartRow = arguments.ContainsKey("dataStartRow") ? Convert.ToInt32(arguments["dataStartRow"]) : 2;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sourceSheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];
            
            int colIndex = GetColumnIndex(sourceSheet, columnName);
            if (colIndex == 0)
                return Task.FromResult(new SkillResult { Success = false, Error = $"未找到列: {columnName}" });

            var usedRange = sourceSheet.UsedRange;
            int lastRow = usedRange.Rows.Count;
            int lastCol = usedRange.Columns.Count;

            var uniqueValues = new HashSet<string>();
            for (int r = dataStartRow; r <= lastRow; r++)
            {
                var val = sourceSheet.Cells[r, colIndex].Text?.ToString() ?? "";
                uniqueValues.Add(val);
            }

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            int createdCount = 0;
            foreach (var val in uniqueValues)
            {
                if (string.IsNullOrEmpty(val)) continue;

                string newSheetName = val.Length > 31 ? val.Substring(0, 31) : val;
                newSheetName = CleanSheetName(newSheetName);

                var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
                if (existingNames.Contains(newSheetName))
                {
                    int suffix = 2;
                    while (existingNames.Contains($"{newSheetName}_{suffix}"))
                        suffix++;
                    newSheetName = $"{newSheetName}_{suffix}";
                }

                var newSheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                newSheet.Name = newSheetName;

                sourceSheet.UsedRange.Rows[1].Copy(newSheet.Rows[1]);

                int destRow = 2;
                for (int r = dataStartRow; r <= lastRow; r++)
                {
                    if (sourceSheet.Cells[r, colIndex].Text?.ToString() == val)
                    {
                        sourceSheet.UsedRange.Rows[r].Copy(newSheet.Rows[destRow]);
                        destRow++;
                    }
                }
                createdCount++;
            }

            sourceSheet.Activate();
            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;

            return Task.FromResult(new SkillResult { Success = true, Content = $"分表完成，共创建 {createdCount} 个工作表" });
        }

        private async Task<SkillResult> SplitAndExportAsync(Dictionary<string, object> arguments)
        {
            var columnName = arguments["columnName"].ToString();
            var outputFolder = arguments["outputFolder"].ToString();
            var fileFormat = arguments.ContainsKey("fileFormat") ? arguments["fileFormat"].ToString() : "xlsx";
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

            var splitResult = await SplitSheetByColumnAsync(new Dictionary<string, object>
            {
                { "columnName", columnName },
                { "sheetName", sheetName }
            });

            if (!splitResult.Success) return splitResult;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sourceSheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];
            
            var sheetsToExport = new List<string>();
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name != sourceSheet.Name)
                    sheetsToExport.Add(sheet.Name);
            }

            return await ExportSheetsAsync(new Dictionary<string, object>
            {
                { "outputFolder", outputFolder },
                { "sheetNames", Newtonsoft.Json.JsonConvert.SerializeObject(sheetsToExport) },
                { "fileFormat", fileFormat }
            });
        }

        private Task<SkillResult> MergeSheetsAsync(Dictionary<string, object> arguments)
        {
            var workbook = ThisAddIn.app.ActiveWorkbook;
            var outputSheetName = arguments.ContainsKey("outputSheetName") ? arguments["outputSheetName"].ToString() : "合并表";
            var includeHeader = !arguments.ContainsKey("includeHeader") || Convert.ToBoolean(arguments["includeHeader"]);

            List<string> sheetNames;
            if (arguments.ContainsKey("sheetNames") && !string.IsNullOrEmpty(arguments["sheetNames"]?.ToString()))
            {
                sheetNames = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["sheetNames"].ToString());
            }
            else
            {
                sheetNames = new List<string>();
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    sheetNames.Add(sheet.Name);
            }

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            string actualMergeName = outputSheetName;
            if (existingNames.Contains(actualMergeName))
            {
                int suffix = 2;
                while (existingNames.Contains($"{actualMergeName}_{suffix}"))
                    suffix++;
                actualMergeName = $"{actualMergeName}_{suffix}";
            }

            var mergeSheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            mergeSheet.Name = actualMergeName;

            int destRow = 1;
            bool firstSheet = true;

            foreach (var name in sheetNames)
            {
                try
                {
                    var sourceSheet = workbook.Worksheets[name];
                    var usedRange = sourceSheet.UsedRange;
                    int lastRow = usedRange.Rows.Count;
                    int lastCol = usedRange.Columns.Count;

                    if (lastRow <= 1 && usedRange.Cells[1, 1].Value == null) continue;

                    int startRow = (firstSheet || includeHeader) ? 1 : 2;
                    
                    for (int r = startRow; r <= lastRow; r++)
                    {
                        for (int c = 1; c <= lastCol; c++)
                        {
                            mergeSheet.Cells[destRow, c].Value = sourceSheet.Cells[r, c].Value;
                        }
                        destRow++;
                    }

                    firstSheet = false;
                }
                catch { }
            }

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;

            return Task.FromResult(new SkillResult { Success = true, Content = $"合并完成，共合并 {sheetNames.Count} 个工作表到 '{actualMergeName}'" });
        }

        private Task<SkillResult> MergeWorkbooksAsync(Dictionary<string, object> arguments)
        {
            var folderPath = arguments["folderPath"].ToString();
            var includeSubfolders = !arguments.ContainsKey("includeSubfolders") || Convert.ToBoolean(arguments["includeSubfolders"]);
            var skipEmptySheets = !arguments.ContainsKey("skipEmptySheets") || Convert.ToBoolean(arguments["skipEmptySheets"]);

            var destWorkbook = ThisAddIn.app.ActiveWorkbook;
            var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var files = Directory.GetFiles(folderPath, "*.xls*", searchOption)
                .Where(f => !f.Contains("~$")).ToList();

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            int importedCount = 0;

            foreach (var file in files)
            {
                try
                {
                    var fileName = Path.GetFileNameWithoutExtension(file);
                    var sourceWorkbook = ThisAddIn.app.Workbooks.Open(file);

                    foreach (Excel.Worksheet sheet in sourceWorkbook.Worksheets)
                    {
                        try
                        {
                            var usedRange = sheet.UsedRange;
                            if (skipEmptySheets && usedRange.Cells.Count == 1 && usedRange.Cells[1, 1].Value == null)
                                continue;

                            var newSheetName = $"{fileName}_{sheet.Name}";
                            if (newSheetName.Length > 31)
                                newSheetName = newSheetName.Substring(0, 31);

                            var newSheet = destWorkbook.Worksheets.Add(After: destWorkbook.Worksheets[destWorkbook.Worksheets.Count]);
                            newSheet.Name = CleanSheetName(newSheetName);
                            usedRange.Copy(newSheet.Range["A1"]);
                            importedCount++;
                        }
                        catch { }
                    }

                    sourceWorkbook.Close(false);
                }
                catch { }
            }

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;

            return Task.FromResult(new SkillResult { Success = true, Content = $"合并完成，共从 {files.Count} 个工作簿导入 {importedCount} 个工作表" });
        }

        private Task<SkillResult> ExportSheetsAsync(Dictionary<string, object> arguments)
        {
            var outputFolder = arguments["outputFolder"].ToString();
            var fileFormat = arguments.ContainsKey("fileFormat") ? arguments["fileFormat"].ToString() : "xlsx";

            var workbook = ThisAddIn.app.ActiveWorkbook;

            List<string> sheetNames;
            if (arguments.ContainsKey("sheetNames") && !string.IsNullOrEmpty(arguments["sheetNames"]?.ToString()))
            {
                sheetNames = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["sheetNames"].ToString());
            }
            else
            {
                sheetNames = new List<string>();
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    sheetNames.Add(sheet.Name);
            }

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            int exportedCount = 0;
            foreach (var name in sheetNames)
            {
                try
                {
                    var sheet = workbook.Worksheets[name];
                    var newWorkbook = ThisAddIn.app.Workbooks.Add();
                    sheet.Copy(newWorkbook.Worksheets[1]);
                    
                    var extension = fileFormat.ToLower() switch
                    {
                        "xls" => ".xls",
                        "csv" => ".csv",
                        _ => ".xlsx"
                    };

                    var xlFileFormat = fileFormat.ToLower() switch
                    {
                        "xls" => Excel.XlFileFormat.xlExcel8,
                        "csv" => Excel.XlFileFormat.xlCSV,
                        _ => Excel.XlFileFormat.xlOpenXMLWorkbook
                    };

                    var filePath = Path.Combine(outputFolder, name + extension);
                    newWorkbook.SaveAs(filePath, xlFileFormat);
                    newWorkbook.Close(false);
                    exportedCount++;
                }
                catch { }
            }

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;

            return Task.FromResult(new SkillResult { Success = true, Content = $"导出完成，共导出 {exportedCount} 个工作表到 {outputFolder}" });
        }

        private Task<SkillResult> DeleteSheetsAsync(Dictionary<string, object> arguments)
        {
            var sheetNames = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["sheetNames"].ToString());
            var workbook = ThisAddIn.app.ActiveWorkbook;

            ThisAddIn.app.DisplayAlerts = false;

            int deletedCount = 0;
            foreach (var name in sheetNames)
            {
                try
                {
                    var sheet = workbook.Worksheets[name];
                    sheet.Delete();
                    deletedCount++;
                }
                catch { }
            }

            ThisAddIn.app.DisplayAlerts = true;

            return Task.FromResult(new SkillResult { Success = true, Content = $"删除完成，共删除 {deletedCount} 个工作表" });
        }

        private Task<SkillResult> TransposeColumnsAsync(Dictionary<string, object> arguments)
        {
            var startColumn = Convert.ToInt32(arguments["startColumn"]);
            var fieldName = arguments["fieldName"].ToString();
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sourceSheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            var usedRange = sourceSheet.UsedRange;
            int lastRow = usedRange.Rows.Count;
            int lastCol = usedRange.Columns.Count;

            var transSheetName = sourceSheet.Name + "_转置表";
            var transSheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            transSheet.Name = transSheetName;

            for (int c = 1; c < startColumn; c++)
            {
                transSheet.Cells[1, c].Value = sourceSheet.Cells[1, c].Value;
            }
            transSheet.Cells[1, startColumn].Value = "值";
            transSheet.Cells[1, startColumn + 1].Value = fieldName;

            int destRow = 2;
            for (int r = 2; r <= lastRow; r++)
            {
                for (int c = startColumn; c <= lastCol; c++)
                {
                    for (int fc = 1; fc < startColumn; fc++)
                    {
                        transSheet.Cells[destRow, fc].Value = sourceSheet.Cells[r, fc].Value;
                    }
                    transSheet.Cells[destRow, startColumn].Value = sourceSheet.Cells[r, c].Value;
                    transSheet.Cells[destRow, startColumn + 1].Value = sourceSheet.Cells[1, c].Value;
                    destRow++;
                }
            }

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;

            return Task.FromResult(new SkillResult { Success = true, Content = $"转置完成，新表名称: {transSheetName}" });
        }

        private Task<SkillResult> ImportSheetsFromFolderAsync(Dictionary<string, object> arguments)
        {
            return MergeWorkbooksAsync(arguments);
        }

        private Task<SkillResult> CreateMultipleSheetsAsync(Dictionary<string, object> arguments)
        {
            var count = Convert.ToInt32(arguments["count"]);
            var baseName = arguments.ContainsKey("baseName") ? arguments["baseName"].ToString() : "新建表";

            if (count > 15) count = 15;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Excel.Worksheet ws in workbook.Worksheets)
                existingNames.Add(ws.Name);

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            int createdCount = 0;
            var skippedNames = new List<string>();

            for (int i = 1; i <= count; i++)
            {
                string targetName = baseName + i.ToString();
                if (existingNames.Contains(targetName))
                {
                    skippedNames.Add(targetName);
                    continue;
                }
                var sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                sheet.Name = targetName;
                existingNames.Add(targetName);
                createdCount++;
            }

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;

            if (skippedNames.Count > 0)
            {
                var suggestions = new List<string>();
                foreach (var name in skippedNames)
                {
                    string newName = name;
                    int suffix = 2;
                    while (existingNames.Contains(newName))
                    {
                        newName = $"{name}_{suffix}";
                        suffix++;
                    }
                    suggestions.Add($"工作表 '{name}' 已存在，可使用 '{newName}' 代替");
                }
                return Task.FromResult(new SkillResult
                {
                    Success = true,
                    Content = $"创建完成，成功创建 {createdCount} 个工作表，跳过 {skippedNames.Count} 个同名工作表",
                    Suggestions = suggestions
                });
            }

            return Task.FromResult(new SkillResult { Success = true, Content = $"创建完成，共创建 {count} 个工作表" });
        }

        private Task<SkillResult> GeneratePayslipsAsync(Dictionary<string, object> arguments)
        {
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
            var outputSheetName = arguments.ContainsKey("outputSheetName") ? arguments["outputSheetName"].ToString() : "工资条";

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sourceSheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            var existingNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            string actualOutputName = outputSheetName;
            if (existingNames.Contains(actualOutputName))
            {
                int suffix = 2;
                while (existingNames.Contains($"{actualOutputName}_{suffix}"))
                    suffix++;
                actualOutputName = $"{actualOutputName}_{suffix}";
            }

            var usedRange = sourceSheet.UsedRange;
            int lastRow = usedRange.Rows.Count;

            var payslipSheet = workbook.Worksheets.Add(Before: sourceSheet);
            payslipSheet.Name = actualOutputName;

            usedRange.Copy(payslipSheet.Range["A1"]);

            for (int n = lastRow; n >= 3; n--)
            {
                payslipSheet.Rows[1].Copy();
                payslipSheet.Rows[n].Insert(Excel.XlDirection.xlDown);
                payslipSheet.Rows[n].Insert(Excel.XlDirection.xlDown);
            }

            payslipSheet.Activate();
            payslipSheet.Range["A1"].Select();

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;

            return Task.FromResult(new SkillResult { Success = true, Content = $"工资条生成完成，新表名称: {actualOutputName}" });
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

        private string CleanSheetName(string name)
        {
            var invalidChars = new[] { '\\', '/', '*', '?', ':', '[', ']' };
            foreach (var c in invalidChars)
                name = name.Replace(c, '_');
            return name;
        }
    }
}
