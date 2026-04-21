using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelPdfSkill : ISkill
    {
        public string Name => "ExcelPdf";
        public string Description => "PDF转换技能，支持Excel工作表导出为PDF";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "export_sheet_to_pdf",
                    Description = "将当前工作表导出为PDF文件。当用户要求将Excel转PDF、导出PDF时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "outputPath", new { type = "string", description = "PDF输出路径（可选，默认与工作簿同目录）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选，默认当前表）" } },
                                { "orientation", new { type = "string", description = "页面方向：横向/纵向，默认横向" } },
                                { "paperSize", new { type = "string", description = "纸张大小：A4/A3/B5，默认A4" } },
                                { "zoom", new { type = "string", description = "缩放方式：无缩放/表自适应/行自适应/列自适应，默认无缩放" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "export_workbook_to_pdf",
                    Description = "将整个工作簿导出为一个PDF文件。当用户要求将整个Excel文件转PDF时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "outputPath", new { type = "string", description = "PDF输出路径（可选）" } },
                                { "sheetNames", new { type = "string", description = "要导出的工作表名称列表（JSON数组格式，可选，默认全部）" } },
                                { "orientation", new { type = "string", description = "页面方向：横向/纵向，默认横向" } },
                                { "paperSize", new { type = "string", description = "纸张大小：A4/A3/B5，默认A4" } },
                                { "zoom", new { type = "string", description = "缩放方式：无缩放/表自适应/行自适应/列自适应，默认无缩放" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "batch_export_sheets_to_pdf",
                    Description = "批量将多个工作表分别导出为独立的PDF文件。当用户要求批量导出PDF时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "outputFolder", new { type = "string", description = "输出文件夹路径（可选，默认与工作簿同目录）" } },
                                { "sheetNames", new { type = "string", description = "工作表名称列表（JSON数组格式，可选，默认全部）" } },
                                { "orientation", new { type = "string", description = "页面方向：横向/纵向，默认横向" } },
                                { "paperSize", new { type = "string", description = "纸张大小：A4/A3/B5，默认A4" } },
                                { "zoom", new { type = "string", description = "缩放方式：无缩放/表自适应/行自适应/列自适应，默认无缩放" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "export_range_to_pdf",
                    Description = "将指定区域导出为PDF。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "range", new { type = "string", description = "要导出的区域（如A1:D20）" } },
                                { "outputPath", new { type = "string", description = "PDF输出路径（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "orientation", new { type = "string", description = "页面方向：横向/纵向，默认横向" } },
                                { "paperSize", new { type = "string", description = "纸张大小：A4/A3/B5，默认A4" } }
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
                    case "export_sheet_to_pdf":
                        return await ExportSheetToPdfAsync(arguments);
                    case "export_workbook_to_pdf":
                        return await ExportWorkbookToPdfAsync(arguments);
                    case "batch_export_sheets_to_pdf":
                        return await BatchExportSheetsToPdfAsync(arguments);
                    case "export_range_to_pdf":
                        return await ExportRangeToPdfAsync(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelPdfSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private void ApplyPageSetup(Excel.Worksheet sheet, Dictionary<string, object> arguments)
        {
            var orientation = arguments.ContainsKey("orientation") ? arguments["orientation"].ToString() : "横向";
            var paperSize = arguments.ContainsKey("paperSize") ? arguments["paperSize"].ToString() : "A4";
            var zoom = arguments.ContainsKey("zoom") ? arguments["zoom"].ToString() : "无缩放";

            var pageSetup = sheet.PageSetup;
            pageSetup.PrintArea = sheet.UsedRange.Address;
            pageSetup.LeftMargin = 0.8;
            pageSetup.RightMargin = 0.8;
            pageSetup.TopMargin = 0.8;
            pageSetup.BottomMargin = 0.8;
            pageSetup.CenterHorizontally = true;
            pageSetup.CenterVertically = true;

            switch (orientation)
            {
                case "纵向":
                    pageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                    break;
                default:
                    pageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                    break;
            }

            switch (paperSize)
            {
                case "A3":
                    pageSetup.PaperSize = Excel.XlPaperSize.xlPaperA3;
                    break;
                case "B5":
                    pageSetup.PaperSize = Excel.XlPaperSize.xlPaperB5;
                    break;
                default:
                    pageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                    break;
            }

            switch (zoom)
            {
                case "表自适应":
                    pageSetup.Zoom = false;
                    pageSetup.FitToPagesWide = 1;
                    pageSetup.FitToPagesTall = 1;
                    break;
                case "行自适应":
                    pageSetup.Zoom = false;
                    pageSetup.FitToPagesTall = 1;
                    break;
                case "列自适应":
                    pageSetup.Zoom = false;
                    pageSetup.FitToPagesTall = 1;
                    break;
                default:
                    pageSetup.Zoom = 100;
                    break;
            }
        }

        private async Task<SkillResult> ExportSheetToPdfAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                var outputPath = arguments.ContainsKey("outputPath") ? arguments["outputPath"].ToString() : null;

                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                if (string.IsNullOrEmpty(outputPath))
                {
                    var workbookPath = workbook.Path;
                    var fileName = Path.GetFileNameWithoutExtension(workbook.Name);
                    outputPath = Path.Combine(workbookPath, $"{fileName}_{sheet.Name}.pdf");
                }

                ThisAddIn.app.ScreenUpdating = false;
                ThisAddIn.app.DisplayAlerts = false;

                try
                {
                    ApplyPageSetup(sheet, arguments);

                    sheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputPath, Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false);

                    return new SkillResult { Success = true, Content = $"工作表 '{sheet.Name}' 已导出为PDF:\n{outputPath}" };
                }
                catch (Exception ex)
                {
                    var suggestions = new List<string>();
                    if (ex.Message.Contains("路径") || ex.Message.Contains("path") || ex.Message.Contains("目录"))
                    {
                        suggestions.Add("1. 检查输出路径是否存在，如果不存在请先创建目录");
                        suggestions.Add("2. 尝试使用其他输出路径，例如桌面或文档文件夹");
                    }
                    else if (ex.Message.Contains("权限") || ex.Message.Contains("permission") || ex.Message.Contains("拒绝"))
                    {
                        suggestions.Add("1. 检查输出路径是否有写入权限");
                        suggestions.Add("2. 尝试以管理员身份运行Excel");
                        suggestions.Add("3. 选择其他有写入权限的目录");
                    }
                    else if (ex.Message.Contains("打印") || ex.Message.Contains("Print") || ex.Message.Contains("printer"))
                    {
                        suggestions.Add("1. 检查是否安装了打印机驱动");
                        suggestions.Add("2. 尝试安装Microsoft Print to PDF虚拟打印机");
                    }
                    else
                    {
                        suggestions.Add("1. 检查工作表是否有数据");
                        suggestions.Add("2. 尝试手动导出PDF确认功能是否正常");
                        suggestions.Add("3. 检查输出路径是否有效");
                    }
                    return SkillResult.FromError($"PDF导出失败: {ex.Message}", suggestions, requiresUserDecision: true);
                }
                finally
                {
                    ThisAddIn.app.ScreenUpdating = true;
                    ThisAddIn.app.DisplayAlerts = true;
                }
            });
        }

        private async Task<SkillResult> ExportWorkbookToPdfAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var workbook = ThisAddIn.app.ActiveWorkbook;
                var outputPath = arguments.ContainsKey("outputPath") ? arguments["outputPath"].ToString() : null;
                var sheetNames = arguments.ContainsKey("sheetNames")
                    ? Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["sheetNames"].ToString())
                    : null;

                if (string.IsNullOrEmpty(outputPath))
                {
                    var workbookPath = workbook.Path;
                    var fileName = Path.GetFileNameWithoutExtension(workbook.Name);
                    outputPath = Path.Combine(workbookPath, $"{fileName}.pdf");
                }

                ThisAddIn.app.ScreenUpdating = false;
                ThisAddIn.app.DisplayAlerts = false;

                try
                {
                    if (sheetNames != null && sheetNames.Count > 0)
                    {
                        var sheets = new List<Excel.Worksheet>();
                        foreach (var name in sheetNames)
                        {
                            try { sheets.Add(workbook.Worksheets[name]); }
                            catch { }
                        }

                        foreach (var s in sheets)
                        {
                            ApplyPageSetup(s, arguments);
                        }

                        if (sheets.Count > 0)
                        {
                            sheets[0].Select();
                            for (int i = 1; i < sheets.Count; i++)
                            {
                                sheets[i].Select(false);
                            }
                        }

                        ThisAddIn.app.Selection.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputPath, Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false);
                    }
                    else
                    {
                        foreach (Excel.Worksheet s in workbook.Worksheets)
                        {
                            ApplyPageSetup(s, arguments);
                        }

                        workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputPath, Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false);
                    }

                    return new SkillResult { Success = true, Content = $"工作簿已导出为PDF:\n{outputPath}" };
                }
                catch (Exception ex)
                {
                    return SkillResult.FromError($"PDF导出失败: {ex.Message}",
                        new List<string>
                        {
                            "1. 检查输出路径是否存在且有写入权限",
                            "2. 检查工作簿中是否有数据",
                            "3. 尝试使用其他输出路径"
                        },
                        requiresUserDecision: true);
                }
                finally
                {
                    ThisAddIn.app.ScreenUpdating = true;
                    ThisAddIn.app.DisplayAlerts = true;
                }
            });
        }

        private async Task<SkillResult> BatchExportSheetsToPdfAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var workbook = ThisAddIn.app.ActiveWorkbook;
                var outputFolder = arguments.ContainsKey("outputFolder")
                    ? arguments["outputFolder"].ToString()
                    : workbook.Path;
                var sheetNames = arguments.ContainsKey("sheetNames")
                    ? Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["sheetNames"].ToString())
                    : null;

                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                var sheetsToExport = new List<Excel.Worksheet>();

                if (sheetNames != null && sheetNames.Count > 0)
                {
                    foreach (var name in sheetNames)
                    {
                        try { sheetsToExport.Add(workbook.Worksheets[name]); }
                        catch { }
                    }
                }
                else
                {
                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        sheetsToExport.Add(sheet);
                    }
                }

                ThisAddIn.app.ScreenUpdating = false;
                ThisAddIn.app.DisplayAlerts = false;

                int successCount = 0;
                int failCount = 0;
                var errors = new List<string>();

                foreach (var sheet in sheetsToExport)
                {
                    var outputPath = Path.Combine(outputFolder, $"{sheet.Name}.pdf");

                    try
                    {
                        ApplyPageSetup(sheet, arguments);

                        sheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputPath, Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false);
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        errors.Add($"{sheet.Name}: {ex.Message}");
                    }
                }

                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;

                var result = $"批量PDF导出完成\n成功: {successCount} 个\n失败: {failCount} 个\n输出目录: {outputFolder}";
                if (errors.Count > 0)
                    result += $"\n失败详情:\n{string.Join("\n", errors.Take(5))}";

                return new SkillResult { Success = failCount == 0, Content = result };
            });
        }

        private async Task<SkillResult> ExportRangeToPdfAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var range = arguments["range"].ToString();
                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                var outputPath = arguments.ContainsKey("outputPath") ? arguments["outputPath"].ToString() : null;

                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                if (string.IsNullOrEmpty(outputPath))
                {
                    var workbookPath = workbook.Path;
                    var fileName = Path.GetFileNameWithoutExtension(workbook.Name);
                    outputPath = Path.Combine(workbookPath, $"{fileName}_{sheet.Name}_range.pdf");
                }

                ThisAddIn.app.ScreenUpdating = false;
                ThisAddIn.app.DisplayAlerts = false;

                try
                {
                    var orientation = arguments.ContainsKey("orientation") ? arguments["orientation"].ToString() : "横向";
                    var paperSize = arguments.ContainsKey("paperSize") ? arguments["paperSize"].ToString() : "A4";

                    var pageSetup = sheet.PageSetup;
                    pageSetup.PrintArea = range;
                    pageSetup.LeftMargin = 0.8;
                    pageSetup.RightMargin = 0.8;
                    pageSetup.TopMargin = 0.8;
                    pageSetup.BottomMargin = 0.8;
                    pageSetup.CenterHorizontally = true;
                    pageSetup.CenterVertically = true;

                    pageSetup.Orientation = orientation == "纵向"
                        ? Excel.XlPageOrientation.xlPortrait
                        : Excel.XlPageOrientation.xlLandscape;

                    pageSetup.PaperSize = paperSize switch
                    {
                        "A3" => Excel.XlPaperSize.xlPaperA3,
                        "B5" => Excel.XlPaperSize.xlPaperB5,
                        _ => Excel.XlPaperSize.xlPaperA4
                    };

                    var exportRange = sheet.Range[range];

                    exportRange.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, outputPath, Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false);

                    return new SkillResult { Success = true, Content = $"区域 '{range}' 已导出为PDF:\n{outputPath}" };
                }
                catch (Exception ex)
                {
                    return SkillResult.FromError($"PDF导出失败: {ex.Message}",
                        new List<string>
                        {
                            "1. 检查输出路径是否存在且有写入权限",
                            "2. 检查指定区域是否有效",
                            "3. 尝试使用其他输出路径"
                        },
                        requiresUserDecision: true);
                }
                finally
                {
                    ThisAddIn.app.ScreenUpdating = true;
                    ThisAddIn.app.DisplayAlerts = true;
                }
            });
        }
    }
}
