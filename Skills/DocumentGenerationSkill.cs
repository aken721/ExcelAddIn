using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace TableMagic.Skills
{
    public class DocumentGenerationSkill : ISkill
    {
        private float? _imageHeightCm = null;
        private float? _imageWidthCm = null;
        private bool _maintainAspectRatio = true;
        private int _pictureInsertMode = 0;

        public string Name => "DocumentGeneration";
        public string Description => "文档生成技能，支持根据Excel数据和Word模板批量生成文档";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "generate_documents",
                    Description = "根据Excel数据和Word模板批量生成文档。模板中使用占位符（如{列名}）作为替换标记，第一列为输出文件名。当用户要求批量生成文档、批量生成Word时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "templatePath", new { type = "string", description = "Word模板文件路径" } },
                                { "outputFolder", new { type = "string", description = "输出文件夹路径" } },
                                { "sheetName", new { type = "string", description = "数据工作表名称（可选，默认当前表）" } },
                                { "placeholderPattern", new { type = "string", description = "占位符格式：{列名}/[列名]/(列名)/【列名】/（列名）/**列名**///列名///##列名##，默认{列名}" } },
                                { "startRow", new { type = "integer", description = "起始行（默认2）" } },
                                { "endRow", new { type = "integer", description = "结束行（可选，默认到最后一行）" } },
                                { "imageHeight", new { type = "number", description = "图片高度（厘米，可选）" } },
                                { "imageWidth", new { type = "number", description = "图片宽度（厘米，可选）" } },
                                { "maintainAspectRatio", new { type = "boolean", description = "是否保持图片宽高比（默认true）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "templatePath", "outputFolder" }
                },
                new SkillTool
                {
                    Name = "preview_document",
                    Description = "预览生成的文档内容（不保存文件）。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "templatePath", new { type = "string", description = "Word模板文件路径" } },
                                { "previewRow", new { type = "integer", description = "预览行号（默认2）" } },
                                { "sheetName", new { type = "string", description = "数据工作表名称（可选）" } },
                                { "placeholderPattern", new { type = "string", description = "占位符格式（默认{列名}）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "templatePath" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "generate_documents":
                        return await GenerateDocumentsAsync(arguments);
                    case "preview_document":
                        return await PreviewDocumentAsync(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in DocumentGenerationSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private async Task<SkillResult> GenerateDocumentsAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var templatePath = arguments["templatePath"].ToString();
                var outputFolder = arguments["outputFolder"].ToString();
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                var placeholderPattern = arguments.ContainsKey("placeholderPattern") ? arguments["placeholderPattern"].ToString() : "{列名}";
                var startRow = arguments.ContainsKey("startRow") ? Convert.ToInt32(arguments["startRow"]) : 2;
                var endRow = arguments.ContainsKey("endRow") ? Convert.ToInt32(arguments["endRow"]) : 0;

                if (arguments.ContainsKey("imageHeight") && arguments["imageHeight"] != null)
                {
                    _imageHeightCm = Convert.ToSingle(arguments["imageHeight"]);
                    _pictureInsertMode = 1;
                }
                if (arguments.ContainsKey("imageWidth") && arguments["imageWidth"] != null)
                {
                    _imageWidthCm = Convert.ToSingle(arguments["imageWidth"]);
                    _pictureInsertMode = 1;
                }
                if (arguments.ContainsKey("maintainAspectRatio"))
                {
                    _maintainAspectRatio = Convert.ToBoolean(arguments["maintainAspectRatio"]);
                }

                if (!File.Exists(templatePath))
                    return new SkillResult { Success = false, Error = $"模板文件不存在: {templatePath}" };

                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                var excelApp = ThisAddIn.app;
                var workbook = excelApp.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                var usedRange = sheet.UsedRange;
                int columnCount = usedRange.Columns.Count;
                int rowCount = usedRange.Rows.Count;
                int lastRow = endRow > 0 ? Math.Min(endRow, rowCount) : rowCount;

                var headers = new List<string>();
                for (int col = 1; col <= columnCount; col++)
                {
                    var cellValue = (usedRange.Cells[1, col] as Excel.Range).Value2;
                    headers.Add(cellValue?.ToString() ?? $"列{col}");
                }

                string templateExtension = Path.GetExtension(templatePath).ToLower();

                var wordApp = new Word.Application { Visible = false };
                int successCount = 0;
                int failCount = 0;

                try
                {
                    for (int row = startRow; row <= lastRow; row++)
                    {
                        var fileNameCell = usedRange.Cells[row, 1] as Excel.Range;
                        if (fileNameCell.Value2 == null || string.IsNullOrWhiteSpace(fileNameCell.Value2.ToString()))
                            continue;

                        string fileName = SanitizeFilename(fileNameCell.Value2.ToString());
                        if (!fileName.EndsWith(templateExtension, StringComparison.OrdinalIgnoreCase))
                        {
                            fileName = Path.ChangeExtension(fileName, templateExtension);
                        }

                        try
                        {
                            var rowData = new Dictionary<string, string>();
                            var imageData = new Dictionary<string, string>();

                            for (int col = 2; col <= columnCount; col++)
                            {
                                var cell = usedRange.Cells[row, col] as Excel.Range;
                                string header = headers[col - 1];
                                string cellValue = GetCellValueAsString(cell);

                                if (IsImageFile(cellValue))
                                {
                                    imageData[header] = cellValue;
                                }
                                else
                                {
                                    rowData[header] = cellValue;
                                }
                            }

                            string outputPath = Path.Combine(outputFolder, fileName);
                            File.Copy(templatePath, outputPath, true);

                            Word.Document wordDoc = null;
                            try
                            {
                                wordDoc = wordApp.Documents.Open(outputPath);

                                ReplacePlaceholders(wordDoc, rowData, placeholderPattern, imageData.Keys);

                                foreach (var kvp in imageData)
                                {
                                    string placeholder = placeholderPattern.Replace("列名", kvp.Key);
                                    float? height = null;
                                    float? width = null;
                                    bool lockRatio = true;

                                    if (_pictureInsertMode == 1)
                                    {
                                        height = _imageHeightCm;
                                        width = _imageWidthCm;
                                        lockRatio = _maintainAspectRatio;
                                    }

                                    InsertPictureAtPlaceholder(wordDoc, placeholder, kvp.Value, height, width, lockRatio);
                                }

                                wordDoc.Save();
                                successCount++;
                            }
                            finally
                            {
                                if (wordDoc != null)
                                {
                                    wordDoc.Close();
                                    Marshal.ReleaseComObject(wordDoc);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            failCount++;
                            Debug.WriteLine($"生成文档时出错: {ex.Message}");
                        }
                    }
                }
                finally
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }

                return new SkillResult { Success = true, Content = $"文档生成完成\n成功: {successCount} 个\n失败: {failCount} 个\n输出目录: {outputFolder}" };
            });
        }

        private async Task<SkillResult> PreviewDocumentAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var templatePath = arguments["templatePath"].ToString();
                var previewRow = arguments.ContainsKey("previewRow") ? Convert.ToInt32(arguments["previewRow"]) : 2;
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                var placeholderPattern = arguments.ContainsKey("placeholderPattern") ? arguments["placeholderPattern"].ToString() : "{列名}";

                if (!File.Exists(templatePath))
                    return new SkillResult { Success = false, Error = $"模板文件不存在: {templatePath}" };

                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                var usedRange = sheet.UsedRange;
                int lastCol = usedRange.Columns.Count;

                var columnMap = new Dictionary<string, int>();
                for (int c = 1; c <= lastCol; c++)
                {
                    var colName = (usedRange.Cells[1, c] as Excel.Range).Value2?.ToString();
                    if (!string.IsNullOrEmpty(colName))
                        columnMap[colName] = c;
                }

                var previewLines = new List<string>();
                foreach (var kvp in columnMap)
                {
                    var placeholder = placeholderPattern.Replace("列名", kvp.Key);
                    var cell = usedRange.Cells[previewRow, kvp.Value] as Excel.Range;
                    var value = GetCellValueAsString(cell);
                    previewLines.Add($"  {placeholder} => {value}");
                }

                return new SkillResult
                {
                    Success = true,
                    Content = $"文档预览（第{previewRow}行数据）:\n{string.Join("\n", previewLines)}"
                };
            });
        }

        private void ReplacePlaceholders(Word.Document doc, Dictionary<string, string> data, string pattern, ICollection<string> imageHeaders)
        {
            try
            {
                if (doc.StoryRanges != null)
                {
                    foreach (Word.Range range in doc.StoryRanges)
                    {
                        if (range != null)
                        {
                            ReplacePlaceholdersInRange(range, data, pattern, imageHeaders);
                        }
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"处理StoryRanges时出错: {ex.Message}"); }

            try
            {
                if (doc.Sections != null && doc.Sections.Count > 0)
                {
                    foreach (Word.Section section in doc.Sections)
                    {
                        if (section != null)
                        {
                            try
                            {
                                if (section.Headers != null)
                                {
                                    Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                                    if (header != null && header.Range != null)
                                        ReplacePlaceholdersInRange(header.Range, data, pattern, imageHeaders);
                                }
                            }
                            catch (Exception ex) { Debug.WriteLine($"处理页眉时出错: {ex.Message}"); }

                            try
                            {
                                if (section.Footers != null)
                                {
                                    Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                                    if (footer != null && footer.Range != null)
                                        ReplacePlaceholdersInRange(footer.Range, data, pattern, imageHeaders);
                                }
                            }
                            catch (Exception ex) { Debug.WriteLine($"处理页脚时出错: {ex.Message}"); }
                        }
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"处理Sections时出错: {ex.Message}"); }

            try
            {
                if (doc.Shapes != null && doc.Shapes.Count > 0)
                {
                    foreach (Word.Shape shape in doc.Shapes)
                    {
                        if (shape != null && shape.TextFrame != null && shape.TextFrame.HasText != 0)
                        {
                            Word.Range textRange = shape.TextFrame.TextRange;
                            if (textRange != null)
                                ReplacePlaceholdersInRange(textRange, data, pattern, imageHeaders);
                        }
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"处理Shapes时出错: {ex.Message}"); }

            try
            {
                if (doc.Tables != null && doc.Tables.Count > 0)
                {
                    foreach (Word.Table table in doc.Tables)
                    {
                        if (table != null)
                        {
                            for (int row = 1; row <= table.Rows.Count; row++)
                            {
                                for (int col = 1; col <= table.Columns.Count; col++)
                                {
                                    try
                                    {
                                        Word.Cell cell = table.Cell(row, col);
                                        if (cell != null && cell.Range != null)
                                            ReplacePlaceholdersInRange(cell.Range, data, pattern, imageHeaders);
                                    }
                                    catch (Exception ex) { Debug.WriteLine($"处理表格单元格({row},{col})时出错: {ex.Message}"); }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"处理Tables时出错: {ex.Message}"); }
        }

        private void ReplacePlaceholdersInRange(Word.Range range, Dictionary<string, string> data, string pattern, ICollection<string> imageHeaders)
        {
            if (range == null) return;

            string originalText = range.Text;

            foreach (var kvp in data)
            {
                if (imageHeaders.Contains(kvp.Key)) continue;

                string placeholder = pattern.Replace("列名", kvp.Key);
                string replacement = kvp.Value
                    .Replace("\n", "\v")
                    .Replace("^p", "\v")
                    .Replace("^l", "\v");

                if (originalText.Contains(placeholder))
                {
                    try
                    {
                        range.Find.ClearFormatting();
                        range.Find.Replacement.ClearFormatting();

                        object replaceAll = Word.WdReplace.wdReplaceAll;
                        object wrap = Word.WdFindWrap.wdFindContinue;

                        range.Find.Execute(
                            FindText: placeholder,
                            ReplaceWith: replacement,
                            Replace: replaceAll,
                            Wrap: wrap
                        );
                    }
                    catch (Exception ex) { Debug.WriteLine($"替换占位符 {placeholder} 时出错: {ex.Message}"); }
                }
            }
        }

        private void InsertPictureAtPlaceholder(Word.Document doc, string placeholder, string imagePath,
                                             float? heightCm, float? widthCm, bool lockAspectRatio)
        {
            try
            {
                if (doc.StoryRanges != null)
                {
                    foreach (Word.Range storyRange in doc.StoryRanges)
                    {
                        Word.Range current = storyRange;
                        while (current != null)
                        {
                            InsertPictureInRange(current, placeholder, imagePath, heightCm, widthCm, lockAspectRatio);
                            try { current = current.NextStoryRange; }
                            catch { current = null; }
                        }
                    }
                }

                if (doc.Shapes != null && doc.Shapes.Count > 0)
                {
                    foreach (Word.Shape shape in doc.Shapes)
                    {
                        try
                        {
                            if (shape != null && shape.TextFrame != null && shape.TextFrame.HasText != 0)
                                InsertPictureInRange(shape.TextFrame.TextRange, placeholder, imagePath, heightCm, widthCm, lockAspectRatio);
                        }
                        catch (Exception ex) { Debug.WriteLine($"处理Shape时出错: {ex.Message}"); }
                    }
                }

                if (doc.Tables != null && doc.Tables.Count > 0)
                {
                    foreach (Word.Table table in doc.Tables)
                    {
                        if (table == null) continue;
                        for (int r = 1; r <= table.Rows.Count; r++)
                        {
                            for (int c = 1; c <= table.Columns.Count; c++)
                            {
                                try
                                {
                                    Word.Cell cell = table.Cell(r, c);
                                    if (cell != null && cell.Range != null)
                                        InsertPictureInRange(cell.Range, placeholder, imagePath, heightCm, widthCm, lockAspectRatio);
                                }
                                catch (Exception ex) { Debug.WriteLine($"处理表格单元格({r},{c})时出错: {ex.Message}"); }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"插入图片时出错: {ex.Message}"); }
        }

        private void InsertPictureInRange(Word.Range range, string placeholder, string imagePath,
                                          float? heightCm, float? widthCm, bool lockAspectRatio)
        {
            if (range == null) return;

            try
            {
                Word.Range searchRange = range.Duplicate;
                searchRange.Find.ClearFormatting();
                searchRange.Find.Replacement.ClearFormatting();

                searchRange.Find.Text = placeholder;
                searchRange.Find.Forward = true;
                searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;
                searchRange.Find.MatchCase = true;
                searchRange.Find.MatchWholeWord = true;

                while (searchRange.Find.Execute())
                {
                    Word.Range foundRange = range.Document.Range(searchRange.Start, searchRange.End);
                    foundRange.Text = string.Empty;

                    Word.InlineShape insertedShape = foundRange.InlineShapes.AddPicture(imagePath, LinkToFile: false, SaveWithDocument: true);
                    SetImageSize(insertedShape, heightCm, widthCm, lockAspectRatio);

                    int newStart = insertedShape.Range.End;
                    int rangeEnd = range.End;
                    if (newStart >= rangeEnd) break;

                    searchRange.SetRange(newStart, rangeEnd);
                    searchRange.Find.ClearFormatting();
                    searchRange.Find.Replacement.ClearFormatting();
                    searchRange.Find.Text = placeholder;
                    searchRange.Find.Forward = true;
                    searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    searchRange.Find.MatchCase = true;
                    searchRange.Find.MatchWholeWord = true;
                }
            }
            catch (Exception ex) { Debug.WriteLine($"InsertPictureInRange 错误: {ex.Message}"); }
        }

        private void SetImageSize(Word.InlineShape shape, float? heightCm, float? widthCm, bool lockAspectRatio)
        {
            try
            {
                if (heightCm == null && widthCm == null) return;

                float originalWidth = shape.Width;
                float originalHeight = shape.Height;
                float aspectRatio = originalWidth / originalHeight;

                const float cmToPoints = 28.35f;
                float? heightPoints = heightCm * cmToPoints;
                float? widthPoints = widthCm * cmToPoints;

                if (lockAspectRatio)
                {
                    float? finalHeight = heightPoints;
                    float? finalWidth = widthPoints;

                    if (heightPoints.HasValue)
                    {
                        finalHeight = heightPoints.Value;
                        finalWidth = heightPoints.Value * aspectRatio;
                    }
                    else if (widthPoints.HasValue)
                    {
                        finalWidth = widthPoints.Value;
                        finalHeight = widthPoints.Value / aspectRatio;
                    }

                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    if (finalHeight.HasValue) shape.Height = finalHeight.Value;
                    if (finalWidth.HasValue) shape.Width = finalWidth.Value;
                }
                else
                {
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    if (heightPoints.HasValue) shape.Height = heightPoints.Value;
                    if (widthPoints.HasValue) shape.Width = widthPoints.Value;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"设置图片尺寸时出错: {ex.Message}");
            }
        }

        private bool IsImageFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return false;
            try
            {
                if (filePath.IndexOfAny(Path.GetInvalidPathChars()) >= 0) return false;
                string extension = Path.GetExtension(filePath).ToLower();
                string[] imageExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff" };
                if (!imageExtensions.Contains(extension)) return false;
                return File.Exists(filePath);
            }
            catch { return false; }
        }

        private string GetCellValueAsString(Excel.Range cell)
        {
            try
            {
                if (IsDateCell(cell))
                {
                    string formattedText = cell.Text?.ToString()?.Trim();
                    if (string.IsNullOrWhiteSpace(formattedText) || formattedText.Contains("####"))
                    {
                        if (cell.Value2 != null)
                        {
                            if (double.TryParse(cell.Value2.ToString(), out double oaDate))
                            {
                                try
                                {
                                    DateTime date = DateTime.FromOADate(oaDate);
                                    return date.ToString("yyyy/MM/dd");
                                }
                                catch { return cell.Value2.ToString(); }
                            }
                            else { return cell.Value2.ToString(); }
                        }
                        else { return string.Empty; }
                    }
                    return formattedText;
                }
                return cell.Text?.ToString()?.Trim() ?? string.Empty;
            }
            catch { return string.Empty; }
        }

        private bool IsDateCell(Excel.Range cell)
        {
            try
            {
                string format = cell.NumberFormat?.ToString() ?? "";
                bool isDateFormat = format.Contains("y") || format.Contains("m") || format.Contains("d") ||
                                   format.Contains("h") || format.Contains("s");

                if (cell.Value2 != null)
                {
                    try
                    {
                        double oaDate = Convert.ToDouble(cell.Value2);
                        if (oaDate > 0 && oaDate < 2958465.99999) return true;
                    }
                    catch { }
                }
                return isDateFormat;
            }
            catch { return false; }
        }

        private string SanitizeFilename(string filename)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            return new string(filename.Where(c => !invalidChars.Contains(c)).ToArray());
        }
    }
}
