using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ZXing;
using ZXing.QrCode.Internal;

namespace TableMagic.Skills
{
    public class ExcelQRSkill : ISkill
    {
        public string Name => "ExcelQR";
        public string Description => "二维码技能，支持生成二维码、条形码以及扫描识别二维码";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "generate_qr_code",
                    Description = "为指定列的数据生成二维码。当用户要求生成二维码时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "columnNames", new { type = "string", description = "要生成二维码的列名列表（JSON数组格式）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "size", new { type = "integer", description = "二维码尺寸（像素，默认100）" } },
                                { "foregroundColor", new { type = "string", description = "前景色（十六进制，默认黑色000000）" } },
                                { "backgroundColor", new { type = "string", description = "背景色（十六进制，默认白色FFFFFF）" } },
                                { "logoPath", new { type = "string", description = "Logo图片路径（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "columnNames" }
                },
                new SkillTool
                {
                    Name = "generate_barcode",
                    Description = "为指定列的数据生成条形码（Code128格式）。当用户要求生成条形码时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "columnName", new { type = "string", description = "要生成条形码的列名" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "width", new { type = "integer", description = "条形码宽度（默认150）" } },
                                { "height", new { type = "integer", description = "条形码高度（默认50）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "columnName" }
                },
                new SkillTool
                {
                    Name = "scan_qr_code",
                    Description = "扫描图片文件中的二维码并返回内容。当用户要求识别二维码时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "imagePaths", new { type = "string", description = "图片文件路径列表（JSON数组格式）" } },
                                { "outputSheetName", new { type = "string", description = "输出工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "imagePaths" }
                },
                new SkillTool
                {
                    Name = "scan_qr_code_folder",
                    Description = "批量扫描文件夹中所有图片的二维码。当用户要求批量识别二维码时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "folderPath", new { type = "string", description = "图片文件夹路径" } },
                                { "includeSubfolders", new { type = "boolean", description = "是否包含子目录（默认true）" } },
                                { "outputSheetName", new { type = "string", description = "输出工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "folderPath" }
                },
                new SkillTool
                {
                    Name = "decode_qr_code_from_range",
                    Description = "从Excel中嵌入的二维码图片解码内容。当用户要求读取Excel中的二维码图片时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
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
                    case "generate_qr_code":
                        return await GenerateQrCodeAsync(arguments);
                    case "generate_barcode":
                        return await GenerateBarcodeAsync(arguments);
                    case "scan_qr_code":
                        return await ScanQrCodeAsync(arguments);
                    case "scan_qr_code_folder":
                        return await ScanQrCodeFolderAsync(arguments);
                    case "decode_qr_code_from_range":
                        return await DecodeQrCodeFromRangeAsync(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelQRSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private async Task<SkillResult> GenerateQrCodeAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var columnNames = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["columnNames"].ToString());
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                var size = arguments.ContainsKey("size") ? Convert.ToInt32(arguments["size"]) : 100;
                var foregroundColor = arguments.ContainsKey("foregroundColor") 
                    ? ColorTranslator.FromHtml("#" + arguments["foregroundColor"].ToString()) 
                    : Color.Black;
                var backgroundColor = arguments.ContainsKey("backgroundColor") 
                    ? ColorTranslator.FromHtml("#" + arguments["backgroundColor"].ToString()) 
                    : Color.White;
                var logoPath = arguments.ContainsKey("logoPath") ? arguments["logoPath"].ToString() : null;

                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                var usedRange = sheet.UsedRange;
                int lastRow = usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;

                var colIndices = new List<int>();
                foreach (var colName in columnNames)
                {
                    int idx = GetColumnIndex(sheet, colName);
                    if (idx > 0) colIndices.Add(idx);
                }

                if (colIndices.Count == 0)
                    return new SkillResult { Success = false, Error = "未找到指定的列" };

                ThisAddIn.app.ScreenUpdating = false;

                int qrCol = lastCol + 1;
                sheet.Cells[1, qrCol].Value = "二维码";

                var writer = new BarcodeWriter
                {
                    Format = BarcodeFormat.QR_CODE,
                    Options = new ZXing.QrCode.QrCodeEncodingOptions
                    {
                        Height = size,
                        Width = size,
                        CharacterSet = "UTF-8",
                        ErrorCorrection = ErrorCorrectionLevel.H,
                        Margin = 0
                    }
                };

                int generatedCount = 0;
                for (int r = 2; r <= lastRow; r++)
                {
                    string data = "";
                    if (colIndices.Count > 1)
                    {
                        foreach (var idx in colIndices)
                        {
                            var key = sheet.Cells[1, idx].Text?.ToString();
                            var value = sheet.Cells[r, idx].Text?.ToString();
                            data += $"{key}:{value};";
                        }
                    }
                    else
                    {
                        data = sheet.Cells[r, colIndices[0]].Text?.ToString() ?? "";
                    }

                    if (string.IsNullOrEmpty(data)) continue;

                    try
                    {
                        var qrBitmap = writer.Write(data);

                        if (!string.IsNullOrEmpty(logoPath) && File.Exists(logoPath))
                        {
                            using (var logo = new Bitmap(logoPath))
                            {
                                int logoSize = size / 5;
                                using (var resizedLogo = new Bitmap(logo, new Size(logoSize, logoSize)))
                                using (var g = Graphics.FromImage(qrBitmap))
                                {
                                    float x = (size - logoSize) / 2f;
                                    float y = (size - logoSize) / 2f;
                                    g.DrawImage(resizedLogo, x, y);
                                }
                            }
                        }

                        var tempPath = Path.Combine(Path.GetTempPath(), $"qr_{Guid.NewGuid()}.png");
                        qrBitmap.Save(tempPath, ImageFormat.Png);

                        var cell = sheet.Cells[r, qrCol];
                        cell.RowHeight = size;
                        cell.ColumnWidth = size / 7;

                        sheet.Shapes.AddPicture(tempPath, 
                            Microsoft.Office.Core.MsoTriState.msoFalse, 
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            cell.Left, cell.Top, size, size);

                        File.Delete(tempPath);
                        generatedCount++;
                    }
                    catch { }
                }

                ThisAddIn.app.ScreenUpdating = true;

                return new SkillResult { Success = true, Content = $"二维码生成完成，共生成 {generatedCount} 个二维码" };
            });
        }

        private async Task<SkillResult> GenerateBarcodeAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var columnName = arguments["columnName"].ToString();
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                var width = arguments.ContainsKey("width") ? Convert.ToInt32(arguments["width"]) : 150;
                var height = arguments.ContainsKey("height") ? Convert.ToInt32(arguments["height"]) : 50;

                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                int colIndex = GetColumnIndex(sheet, columnName);
                if (colIndex == 0)
                    return new SkillResult { Success = false, Error = $"未找到列: {columnName}" };

                var usedRange = sheet.UsedRange;
                int lastRow = usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;

                ThisAddIn.app.ScreenUpdating = false;

                int barcodeCol = lastCol + 1;
                sheet.Cells[1, barcodeCol].Value = "条形码";

                var writer = new BarcodeWriter
                {
                    Format = BarcodeFormat.CODE_128,
                    Options = new ZXing.QrCode.QrCodeEncodingOptions
                    {
                        Height = height,
                        Width = width,
                        Margin = 1,
                        PureBarcode = true
                    }
                };

                int generatedCount = 0;
                var asciiPattern = new System.Text.RegularExpressions.Regex(@"^[\x00-\x7F]*$");

                for (int r = 2; r <= lastRow; r++)
                {
                    var value = sheet.Cells[r, colIndex].Text?.ToString();
                    if (string.IsNullOrEmpty(value)) continue;
                    if (!asciiPattern.IsMatch(value)) continue;

                    try
                    {
                        var barcodeBitmap = writer.Write(value);
                        var tempPath = Path.Combine(Path.GetTempPath(), $"bc_{Guid.NewGuid()}.png");
                        barcodeBitmap.Save(tempPath, ImageFormat.Png);

                        var cell = sheet.Cells[r, barcodeCol];
                        cell.RowHeight = height;
                        cell.ColumnWidth = width / 7;

                        sheet.Shapes.AddPicture(tempPath,
                            Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue,
                            cell.Left, cell.Top, width, height);

                        File.Delete(tempPath);
                        generatedCount++;
                    }
                    catch { }
                }

                ThisAddIn.app.ScreenUpdating = true;

                return new SkillResult { Success = true, Content = $"条形码生成完成，共生成 {generatedCount} 个条形码" };
            });
        }

        private async Task<SkillResult> ScanQrCodeAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var imagePaths = Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["imagePaths"].ToString());
                var outputSheetName = arguments.ContainsKey("outputSheetName") 
                    ? arguments["outputSheetName"].ToString() 
                    : "二维码识别结果";

                var results = new List<(string Path, string Content)>();
                var reader = new BarcodeReader();

                foreach (var path in imagePaths)
                {
                    if (!File.Exists(path)) continue;

                    try
                    {
                        using (var bitmap = new Bitmap(path))
                        {
                            var result = reader.Decode(bitmap);
                            results.Add((path, result?.Text ?? "无法识别"));
                        }
                    }
                    catch (Exception ex)
                    {
                        results.Add((path, $"错误: {ex.Message}"));
                    }
                }

                WriteScanResultsToExcel(results, outputSheetName);

                return new SkillResult 
                { 
                    Success = true, 
                    Content = $"二维码识别完成，共处理 {imagePaths.Count} 个文件，成功识别 {results.Count(r => r.Content != "无法识别" && !r.Content.StartsWith("错误"))} 个" 
                };
            });
        }

        private async Task<SkillResult> ScanQrCodeFolderAsync(Dictionary<string, object> arguments)
        {
            var folderPath = arguments["folderPath"].ToString();
            var includeSubfolders = !arguments.ContainsKey("includeSubfolders") || Convert.ToBoolean(arguments["includeSubfolders"]);
            var outputSheetName = arguments.ContainsKey("outputSheetName") 
                ? arguments["outputSheetName"].ToString() 
                : "二维码识别结果";

            var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            var extensions = new[] { "*.jpg", "*.jpeg", "*.png", "*.bmp" };
            var imagePaths = new List<string>();

            foreach (var ext in extensions)
            {
                imagePaths.AddRange(Directory.GetFiles(folderPath, ext, searchOption));
            }

            return await ScanQrCodeAsync(new Dictionary<string, object>
            {
                { "imagePaths", Newtonsoft.Json.JsonConvert.SerializeObject(imagePaths) },
                { "outputSheetName", outputSheetName }
            });
        }

        private async Task<SkillResult> DecodeQrCodeFromRangeAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                var results = new List<(string Name, string Content)>();
                var reader = new BarcodeReader();

                foreach (Excel.Shape shape in sheet.Shapes)
                {
                    try
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                        {
                            shape.Copy();
                            using (var img = (Bitmap)Clipboard.GetImage())
                            {
                                if (img != null)
                                {
                                    var result = reader.Decode(img);
                                    results.Add((shape.Name, result?.Text ?? "无法识别"));
                                }
                            }
                        }
                    }
                    catch { }
                }

                return new SkillResult 
                { 
                    Success = true, 
                    Content = results.Count > 0 
                        ? $"识别结果：\n{string.Join("\n", results.Select(r => $"{r.Name}: {r.Content}"))}" 
                        : "未找到可识别的二维码图片" 
                };
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

        private void WriteScanResultsToExcel(List<(string Path, string Content)> results, string sheetName)
        {
            var workbook = ThisAddIn.app.ActiveWorkbook;

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            Excel.Worksheet sheet;
            
            var existingNames = new List<string>();
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            
            string actualSheetName = sheetName;
            if (existingNames.Any(n => string.Equals(n, sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                int suffix = 2;
                while (existingNames.Any(n => string.Equals(n, $"{sheetName}_{suffix}", StringComparison.OrdinalIgnoreCase)))
                {
                    suffix++;
                }
                actualSheetName = $"{sheetName}_{suffix}";
            }
            
            sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            sheet.Name = actualSheetName;

            sheet.Cells[1, 1].Value = "文件路径";
            sheet.Cells[1, 2].Value = "识别内容";
            sheet.Cells[1, 1].Font.Bold = true;
            sheet.Cells[1, 2].Font.Bold = true;

            for (int i = 0; i < results.Count; i++)
            {
                sheet.Cells[i + 2, 1].Value = results[i].Path;
                sheet.Cells[i + 2, 2].Value = results[i].Content;
            }

            sheet.UsedRange.Columns.AutoFit();
            sheet.Activate();

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;
        }
    }
}
