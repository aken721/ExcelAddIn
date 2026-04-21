using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelFileSkill : ISkill
    {
        public string Name => "ExcelFile";
        public string Description => "文件/文件夹技能，支持批量重命名、复制、移动、删除等文件操作";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "list_files",
                    Description = "列出文件夹中的文件信息。当用户要求查看文件列表、列出文件时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "folderPath", new { type = "string", description = "文件夹路径" } },
                                { "pattern", new { type = "string", description = "文件筛选模式（如*.xlsx，可选）" } },
                                { "includeSubfolders", new { type = "boolean", description = "是否包含子目录（默认false）" } },
                                { "outputSheetName", new { type = "string", description = "输出工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "folderPath" }
                },
                new SkillTool
                {
                    Name = "batch_rename",
                    Description = "根据Excel数据批量重命名文件。Excel中需包含原文件名和新文件名列。当用户要求批量重命名文件时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "oldNameColumn", new { type = "string", description = "原文件名列名" } },
                                { "newNameColumn", new { type = "string", description = "新文件名列名" } },
                                { "folderPath", new { type = "string", description = "文件所在文件夹路径（可选，不填则使用完整路径）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "oldNameColumn", "newNameColumn" }
                },
                new SkillTool
                {
                    Name = "batch_copy",
                    Description = "批量复制文件到目标文件夹。当用户要求批量复制文件时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileNameColumn", new { type = "string", description = "文件名/路径列名" } },
                                { "targetFolder", new { type = "string", description = "目标文件夹路径" } },
                                { "sourceFolder", new { type = "string", description = "源文件夹路径（可选，不填则使用完整路径）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "fileNameColumn", "targetFolder" }
                },
                new SkillTool
                {
                    Name = "batch_move",
                    Description = "批量移动文件到目标文件夹。当用户要求批量移动文件时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileNameColumn", new { type = "string", description = "文件名/路径列名" } },
                                { "targetFolder", new { type = "string", description = "目标文件夹路径" } },
                                { "sourceFolder", new { type = "string", description = "源文件夹路径（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "fileNameColumn", "targetFolder" }
                },
                new SkillTool
                {
                    Name = "batch_delete",
                    Description = "批量删除文件。当用户要求批量删除文件时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileNameColumn", new { type = "string", description = "文件名/路径列名" } },
                                { "folderPath", new { type = "string", description = "文件所在文件夹路径（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "fileNameColumn" }
                },
                new SkillTool
                {
                    Name = "create_folder",
                    Description = "创建文件夹。当用户要求创建文件夹、新建目录时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "folderPath", new { type = "string", description = "文件夹路径" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "folderPath" }
                },
                new SkillTool
                {
                    Name = "get_file_info",
                    Description = "获取文件详细信息（大小、创建时间、修改时间等）。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "filePath", new { type = "string", description = "文件路径" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "filePath" }
                },
                new SkillTool
                {
                    Name = "open_folder",
                    Description = "在资源管理器中打开文件夹。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "folderPath", new { type = "string", description = "文件夹路径" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "folderPath" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "list_files":
                        return await ListFilesAsync(arguments);
                    case "batch_rename":
                        return await BatchRenameAsync(arguments);
                    case "batch_copy":
                        return await BatchCopyAsync(arguments);
                    case "batch_move":
                        return await BatchMoveAsync(arguments);
                    case "batch_delete":
                        return await BatchDeleteAsync(arguments);
                    case "create_folder":
                        return CreateFolder(arguments);
                    case "get_file_info":
                        return GetFileInfo(arguments);
                    case "open_folder":
                        return OpenFolder(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelFileSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private async Task<SkillResult> ListFilesAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var folderPath = arguments["folderPath"].ToString();
                var pattern = arguments.ContainsKey("pattern") ? arguments["pattern"].ToString() : "*.*";
                var includeSubfolders = arguments.ContainsKey("includeSubfolders") && Convert.ToBoolean(arguments["includeSubfolders"]);
                var outputSheetName = arguments.ContainsKey("outputSheetName") 
                    ? arguments["outputSheetName"].ToString() 
                    : "文件列表";

                if (!Directory.Exists(folderPath))
                    return new SkillResult { Success = false, Error = $"文件夹不存在: {folderPath}" };

                var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                var files = Directory.GetFiles(folderPath, pattern, searchOption);

                var workbook = ThisAddIn.app.ActiveWorkbook;

                ThisAddIn.app.ScreenUpdating = false;
                ThisAddIn.app.DisplayAlerts = false;

                Excel.Worksheet sheet;
                
                var existingNames = new List<string>();
                foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
                
                string actualSheetName = outputSheetName;
                if (existingNames.Any(n => string.Equals(n, outputSheetName, StringComparison.OrdinalIgnoreCase)))
                {
                    int suffix = 2;
                    while (existingNames.Any(n => string.Equals(n, $"{outputSheetName}_{suffix}", StringComparison.OrdinalIgnoreCase)))
                    {
                        suffix++;
                    }
                    actualSheetName = $"{outputSheetName}_{suffix}";
                }
                
                sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                sheet.Name = actualSheetName;

                var headers = new[] { "文件名", "扩展名", "大小(KB)", "创建时间", "修改时间", "完整路径" };
                for (int c = 0; c < headers.Length; c++)
                {
                    sheet.Cells[1, c + 1].Value = headers[c];
                    sheet.Cells[1, c + 1].Font.Bold = true;
                }

                for (int i = 0; i < files.Length; i++)
                {
                    var file = files[i];
                    var fileInfo = new FileInfo(file);

                    sheet.Cells[i + 2, 1].Value = fileInfo.Name;
                    sheet.Cells[i + 2, 2].Value = fileInfo.Extension;
                    sheet.Cells[i + 2, 3].Value = Math.Round(fileInfo.Length / 1024.0, 2);
                    sheet.Cells[i + 2, 4].Value = fileInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss");
                    sheet.Cells[i + 2, 5].Value = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss");
                    sheet.Cells[i + 2, 6].Value = file;
                }

                sheet.UsedRange.Columns.AutoFit();
                sheet.Activate();

                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;

                return new SkillResult { Success = true, Content = $"文件列表已导出，共 {files.Length} 个文件" };
            });
        }

        private async Task<SkillResult> BatchRenameAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var oldNameColumn = arguments["oldNameColumn"].ToString();
                var newNameColumn = arguments["newNameColumn"].ToString();
                var folderPath = arguments.ContainsKey("folderPath") ? arguments["folderPath"].ToString() : null;
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                var usedRange = sheet.UsedRange;
                int lastRow = usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;

                var columnMap = new Dictionary<string, int>();
                for (int c = 1; c <= lastCol; c++)
                {
                    var colName = sheet.Cells[1, c].Text?.ToString();
                    if (!string.IsNullOrEmpty(colName))
                        columnMap[colName] = c;
                }

                if (!columnMap.ContainsKey(oldNameColumn) || !columnMap.ContainsKey(newNameColumn))
                    return new SkillResult { Success = false, Error = "未找到指定的列" };

                int oldColIdx = columnMap[oldNameColumn];
                int newColIdx = columnMap[newNameColumn];

                int successCount = 0;
                int failCount = 0;
                var errors = new List<string>();

                for (int r = 2; r <= lastRow; r++)
                {
                    var oldName = sheet.Cells[r, oldColIdx].Text?.ToString();
                    var newName = sheet.Cells[r, newColIdx].Text?.ToString();

                    if (string.IsNullOrEmpty(oldName) || string.IsNullOrEmpty(newName)) continue;

                    string oldPath, newPath;

                    if (!string.IsNullOrEmpty(folderPath))
                    {
                        oldPath = Path.Combine(folderPath, oldName);
                        newPath = Path.Combine(folderPath, newName);
                    }
                    else
                    {
                        oldPath = oldName;
                        newPath = newName;
                    }

                    try
                    {
                        if (File.Exists(oldPath))
                        {
                            File.Move(oldPath, newPath);
                            successCount++;
                        }
                        else if (Directory.Exists(oldPath))
                        {
                            Directory.Move(oldPath, newPath);
                            successCount++;
                        }
                        else
                        {
                            failCount++;
                            errors.Add($"行{r}: 文件不存在 {oldPath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        errors.Add($"行{r}: {ex.Message}");
                    }
                }

                var result = $"批量重命名完成\n成功: {successCount} 个\n失败: {failCount} 个";
                if (errors.Count > 0)
                    result += $"\n失败详情:\n{string.Join("\n", errors.Take(10))}";

                return new SkillResult { Success = failCount == 0, Content = result };
            });
        }

        private async Task<SkillResult> BatchCopyAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var fileNameColumn = arguments["fileNameColumn"].ToString();
                var targetFolder = arguments["targetFolder"].ToString();
                var sourceFolder = arguments.ContainsKey("sourceFolder") ? arguments["sourceFolder"].ToString() : null;
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                var usedRange = sheet.UsedRange;
                int lastRow = usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;

                var columnMap = new Dictionary<string, int>();
                for (int c = 1; c <= lastCol; c++)
                {
                    var colName = sheet.Cells[1, c].Text?.ToString();
                    if (!string.IsNullOrEmpty(colName))
                        columnMap[colName] = c;
                }

                if (!columnMap.ContainsKey(fileNameColumn))
                    return new SkillResult { Success = false, Error = $"未找到列: {fileNameColumn}" };

                int fileColIdx = columnMap[fileNameColumn];

                int successCount = 0;
                int failCount = 0;

                for (int r = 2; r <= lastRow; r++)
                {
                    var fileName = sheet.Cells[r, fileColIdx].Text?.ToString();
                    if (string.IsNullOrEmpty(fileName)) continue;

                    string sourcePath = !string.IsNullOrEmpty(sourceFolder) 
                        ? Path.Combine(sourceFolder, fileName) 
                        : fileName;

                    var targetPath = Path.Combine(targetFolder, Path.GetFileName(fileName));

                    try
                    {
                        if (File.Exists(sourcePath))
                        {
                            File.Copy(sourcePath, targetPath, true);
                            successCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }
                    catch
                    {
                        failCount++;
                    }
                }

                return new SkillResult { Success = true, Content = $"批量复制完成\n成功: {successCount} 个\n失败: {failCount} 个" };
            });
        }

        private async Task<SkillResult> BatchMoveAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var fileNameColumn = arguments["fileNameColumn"].ToString();
                var targetFolder = arguments["targetFolder"].ToString();
                var sourceFolder = arguments.ContainsKey("sourceFolder") ? arguments["sourceFolder"].ToString() : null;
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                if (!Directory.Exists(targetFolder))
                    Directory.CreateDirectory(targetFolder);

                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                var usedRange = sheet.UsedRange;
                int lastRow = usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;

                var columnMap = new Dictionary<string, int>();
                for (int c = 1; c <= lastCol; c++)
                {
                    var colName = sheet.Cells[1, c].Text?.ToString();
                    if (!string.IsNullOrEmpty(colName))
                        columnMap[colName] = c;
                }

                if (!columnMap.ContainsKey(fileNameColumn))
                    return new SkillResult { Success = false, Error = $"未找到列: {fileNameColumn}" };

                int fileColIdx = columnMap[fileNameColumn];

                int successCount = 0;
                int failCount = 0;

                for (int r = 2; r <= lastRow; r++)
                {
                    var fileName = sheet.Cells[r, fileColIdx].Text?.ToString();
                    if (string.IsNullOrEmpty(fileName)) continue;

                    string sourcePath = !string.IsNullOrEmpty(sourceFolder) 
                        ? Path.Combine(sourceFolder, fileName) 
                        : fileName;

                    var targetPath = Path.Combine(targetFolder, Path.GetFileName(fileName));

                    try
                    {
                        if (File.Exists(sourcePath))
                        {
                            File.Move(sourcePath, targetPath);
                            successCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }
                    catch
                    {
                        failCount++;
                    }
                }

                return new SkillResult { Success = true, Content = $"批量移动完成\n成功: {successCount} 个\n失败: {failCount} 个" };
            });
        }

        private async Task<SkillResult> BatchDeleteAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var fileNameColumn = arguments["fileNameColumn"].ToString();
                var folderPath = arguments.ContainsKey("folderPath") ? arguments["folderPath"].ToString() : null;
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                var usedRange = sheet.UsedRange;
                int lastRow = usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;

                var columnMap = new Dictionary<string, int>();
                for (int c = 1; c <= lastCol; c++)
                {
                    var colName = sheet.Cells[1, c].Text?.ToString();
                    if (!string.IsNullOrEmpty(colName))
                        columnMap[colName] = c;
                }

                if (!columnMap.ContainsKey(fileNameColumn))
                    return new SkillResult { Success = false, Error = $"未找到列: {fileNameColumn}" };

                int fileColIdx = columnMap[fileNameColumn];

                int successCount = 0;
                int failCount = 0;

                for (int r = 2; r <= lastRow; r++)
                {
                    var fileName = sheet.Cells[r, fileColIdx].Text?.ToString();
                    if (string.IsNullOrEmpty(fileName)) continue;

                    string filePath = !string.IsNullOrEmpty(folderPath) 
                        ? Path.Combine(folderPath, fileName) 
                        : fileName;

                    try
                    {
                        if (File.Exists(filePath))
                        {
                            File.Delete(filePath);
                            successCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }
                    catch
                    {
                        failCount++;
                    }
                }

                return new SkillResult { Success = true, Content = $"批量删除完成\n成功: {successCount} 个\n失败: {failCount} 个" };
            });
        }

        private SkillResult CreateFolder(Dictionary<string, object> arguments)
        {
            var folderPath = arguments["folderPath"].ToString();

            try
            {
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                    return new SkillResult { Success = true, Content = $"文件夹创建成功: {folderPath}" };
                }
                else
                {
                    return new SkillResult { Success = true, Content = $"文件夹已存在: {folderPath}" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = $"创建文件夹失败: {ex.Message}" };
            }
        }

        private SkillResult GetFileInfo(Dictionary<string, object> arguments)
        {
            var filePath = arguments["filePath"].ToString();

            if (!File.Exists(filePath))
                return new SkillResult { Success = false, Error = $"文件不存在: {filePath}" };

            var fileInfo = new FileInfo(filePath);

            return new SkillResult 
            { 
                Success = true, 
                Content = $"文件信息:\n名称: {fileInfo.Name}\n大小: {fileInfo.Length / 1024.0:F2} KB\n创建时间: {fileInfo.CreationTime}\n修改时间: {fileInfo.LastWriteTime}\n路径: {fileInfo.FullName}" 
            };
        }

        private SkillResult OpenFolder(Dictionary<string, object> arguments)
        {
            var folderPath = arguments["folderPath"].ToString();

            if (!Directory.Exists(folderPath))
                return new SkillResult { Success = false, Error = $"文件夹不存在: {folderPath}" };

            try
            {
                System.Diagnostics.Process.Start("explorer.exe", folderPath);
                return new SkillResult { Success = true, Content = $"已打开文件夹: {folderPath}" };
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = $"打开文件夹失败: {ex.Message}" };
            }
        }
    }
}
