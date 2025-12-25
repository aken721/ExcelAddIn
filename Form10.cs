using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelAddIn
{
    public partial class Form10 : Form
    {
        public Form10()
        {
            InitializeComponent();
        }


        private float? imageHeightCm = null;
        private float? imageWidthCm = null;
        private bool maintainAspectRatio = true;
        private int pictureInsertMode = 0;                  // 0=原尺寸, 1=缩放插入

        private void Form10_Load(object sender, EventArgs e)
        {
            foreach (Excel.Worksheet worksheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
            {
                sheets_name_comboBox.Items.Add(worksheet.Name);
            }
            string currentSheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
            if (sheets_name_comboBox.Items.Count > 0)
            {
                sheets_name_comboBox.Text = currentSheetName; // 默认选择当前工作表名称
            }
            placeholder_comboBox.SelectedIndex = 3;
            flowLayoutPanel2.Visible = false; // 显示尺寸设置面板
            flowLayoutPanel3.Visible = false; // 显示生成文档面板

            pictureOriginal_comboBox.SelectedIndex = 0; // 默认图片原始尺寸
            pictureSize_checkBox.Checked = true;
            width_textBox.ReadOnly = true; // 默认锁定比例时宽度不可编辑
            if (pictureOriginal_comboBox.SelectedIndex == 0)
            {
                height_textBox.Text = string.Empty;  // 默认高度3厘米
                width_textBox.Text = string.Empty;   // 默认宽度3厘米
            }
            else
            {
                height_textBox.Text = "3";   // 默认高度3厘米
            }

            radioButtonAll.Checked = true;     // 默认处理所有行
        }

        private void pictureOriginal_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (pictureOriginal_comboBox.SelectedIndex)
            {
                case 0:
                    pictureInsertMode = 0;
                    height_textBox.Text = string.Empty;  // 清空高度
                    width_textBox.Text = string.Empty;   // 清空宽度
                    flowLayoutPanel2.Visible = false; // 隐藏高设置面板
                    flowLayoutPanel3.Visible = false; // 隐藏宽设置面板
                    pictureSize_checkBox.Visible = false; // 隐藏锁定比例复选框
                    break;
                case 1:
                    pictureInsertMode = 1;
                    flowLayoutPanel2.Visible = true; // 显示高设置面板
                    flowLayoutPanel3.Visible = true; // 显示宽设置面板
                    pictureSize_checkBox.Visible = true; // 隐藏锁定比例复选框
                    break;
            }
        }

        private void model_select_button_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Word Documents (*.docx;*.doc)|*.docx;*.doc";
            openFileDialog1.Title = "选择Word模板文件";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.InitialDirectory = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                docModel_label.Text = openFileDialog1.FileName;
            }
        }

        private void doc_folder_button_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "选择生成的Word文档保存目录";
            folderBrowserDialog1.SelectedPath = Globals.ThisAddIn.Application.ActiveWorkbook.Path;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                docGenerated_label.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private async void docRun_button_Click(object sender, EventArgs e)
        {
            // 根据插入模式决定是否需要验证尺寸
            if (pictureInsertMode == 1 && !ValidateAndSaveDimensions())
            {
                return;
            }

            // 重置状态标签
            result_doc_label.Text = "准备生成文档...";
            Application.DoEvents();

            // 获取窗体控件值
            string sheetName = sheets_name_comboBox.SelectedItem?.ToString();
            string placeholderPattern = placeholder_comboBox.SelectedItem?.ToString();
            string templatePath = docModel_label.Text;
            string outputFolder = docGenerated_label.Text;
            bool processSelectedOnly = radioButtonSelected.Checked;

            // 在后台执行生成任务，避免阻塞UI
            await Task.Run(() => GenerateDocuments(sheetName, placeholderPattern, templatePath, outputFolder, processSelectedOnly));
        }

        private void GenerateDocuments(string sheetName, string placeholderPattern, string templatePath, string outputFolder, bool processSelectedOnly = false)
        {
            try
            {
                // 验证输入
                if (string.IsNullOrEmpty(sheetName))
                {
                    ShowErrorMessage("请选择数据表");
                    return;
                }

                if (string.IsNullOrEmpty(placeholderPattern))
                {
                    ShowErrorMessage("请选择占位符格式");
                    return;
                }

                if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
                {
                    ShowErrorMessage("请选择有效的Word模板");
                    return;
                }

                if (string.IsNullOrEmpty(outputFolder) || !Directory.Exists(outputFolder))
                {
                    ShowErrorMessage("请选择有效的输出目录");
                    return;
                }

                // 获取Excel应用和工作表
                var excelApp = Globals.ThisAddIn.Application;
                Excel.Worksheet sheet = excelApp.ActiveWorkbook.Sheets[sheetName] as Excel.Worksheet;
                Excel.Range usedRange = sheet.UsedRange;

                // 获取列标题
                int columnCount = usedRange.Columns.Count;
                int rowCount = usedRange.Rows.Count;
                int totalDocuments = 0;

                // 获取要处理的行号列表
                List<int> rowsToProcess = new List<int>();

                if (processSelectedOnly)
                {
                    // 只处理选中的行
                    Excel.Range selection = excelApp.Selection as Excel.Range;
                    if (selection != null)
                    {
                        foreach (Excel.Range area in selection.Areas)
                        {
                            foreach (Excel.Range row in area.Rows)
                            {
                                int rowNum = row.Row;
                                // 跳过标题行（第1行）且确保在数据范围内
                                if (rowNum > 1 && rowNum <= rowCount && !rowsToProcess.Contains(rowNum))
                                {
                                    var fileNameCell = usedRange.Cells[rowNum, 1] as Excel.Range;
                                    if (fileNameCell.Value2 != null &&
                                        !string.IsNullOrWhiteSpace(fileNameCell.Value2.ToString()))
                                    {
                                        rowsToProcess.Add(rowNum);
                                    }
                                }
                            }
                        }
                    }

                    if (rowsToProcess.Count == 0)
                    {
                        ShowErrorMessage("请先在Excel中选择要处理的数据行（不包括标题行）");
                        return;
                    }

                    totalDocuments = rowsToProcess.Count;
                }
                else
                {
                    // 处理所有行 - 预计算有效行数（第一列为文件名的非空行）
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var fileNameCell = usedRange.Cells[row, 1] as Excel.Range;
                        if (fileNameCell.Value2 != null &&
                            !string.IsNullOrWhiteSpace(fileNameCell.Value2.ToString()))
                        {
                            rowsToProcess.Add(row);
                            totalDocuments++;
                        }
                    }
                }

                if (totalDocuments < 1)
                {
                    ShowErrorMessage("数据表中没有可处理的数据行");
                    return;
                }

                var headers = new List<string>();
                for (int col = 1; col <= columnCount; col++)
                {
                    var cellValue = (usedRange.Cells[1, col] as Excel.Range).Value2;
                    headers.Add(cellValue?.ToString() ?? $"列{col}");
                }

                // 获取模板文件扩展名
                string templateExtension = Path.GetExtension(templatePath).ToLower();

                // 启动Word应用程序
                var wordApp = new Word.Application { Visible = false };
                int successCount = 0;
                int errorCount = 0;
                int currentDocument = 0;

                // 处理数据行（使用预先计算的行号列表）
                foreach (int row in rowsToProcess)
                {
                    // 获取文件名单元格
                    var fileNameCell = usedRange.Cells[row, 1] as Excel.Range;

                    currentDocument++;
                    try
                    {
                        // 更新状态标签
                        UpdateStatusLabel($"正在生成第 {currentDocument}/{totalDocuments} 个文档...");

                        // 获取文件名（第一列）
                        string fileName = fileNameCell.Value2?.ToString() ?? $"文档_{currentDocument}";

                        // 处理文件名中的非法字符
                        fileName = SanitizeFilename(fileName);

                        // 使用模板文件相同的扩展名
                        if (!fileName.EndsWith(templateExtension, StringComparison.OrdinalIgnoreCase))
                        {
                            fileName = Path.ChangeExtension(fileName, templateExtension);
                        }

                        var rowData = new Dictionary<string, string>();
                        var imageData = new Dictionary<string, string>();         // 存储图片路径

                        // 收集数据（遍历所有列）
                        for (int col = 2; col <= columnCount; col++)
                        {
                            var cell = usedRange.Cells[row, col] as Excel.Range;
                            string header = headers[col - 1];
                            string cellValue = GetCellValueAsString(cell);


                            bool isImageFile = IsImageFile(cellValue);
                            // 检查是否是图片文件
                            if (isImageFile)
                            {
                                imageData[header] = cellValue;
                            }
                            else
                            {
                                rowData[header] = cellValue;
                            }
                        }

                        // 生成输出路径
                        string outputPath = Path.Combine(outputFolder, fileName);

                        // 复制模板并替换内容
                        File.Copy(templatePath, outputPath, true);

                        Word.Document wordDoc = null;
                        try
                        {
                            wordDoc = wordApp.Documents.Open(outputPath);

                            // 1. 先替换文本占位符（跳过图片占位符）
                            ReplacePlaceholders(wordDoc, rowData, placeholderPattern, imageData.Keys);

                            // 2. 单独处理图片占位符
                            foreach (var kvp in imageData)
                            {
                                string placeholder = placeholderPattern.Replace("列名", kvp.Key);

                                // 根据插入模式决定尺寸参数
                                float? height = null;
                                float? width = null;
                                bool lockRatio = true;

                                if (pictureInsertMode == 1) // 缩放插入
                                {
                                    height = imageHeightCm;
                                    width = imageWidthCm;
                                    lockRatio = maintainAspectRatio;
                                }

                                InsertPictureAtPlaceholder(
                                    wordDoc,
                                    placeholder,
                                    kvp.Value,
                                    height,
                                    width,
                                    lockRatio
                                );
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
                        errorCount++;
                        // 记录错误但继续处理
                        Debug.WriteLine($"生成文档时出错: {ex.Message}");
                    }
                }

                // 清理Word资源
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);

                // 更新最终状态
                UpdateStatusLabel($"已生成 {successCount} 个文档");

                // 打开目标文件夹
                try
                {
                    Process.Start("explorer.exe", outputFolder);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"无法打开文件夹: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                UpdateStatusLabel("生成过程中出错");
                ShowErrorMessage($"生成文档时出错: {ex.Message}");
            }
        }

        // 检查是否是图片文件
        private bool IsImageFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return false;

            try
            {
                // 检查路径是否包含非法字符
                if (filePath.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
                    return false;

                // 检查文件扩展名
                string extension = Path.GetExtension(filePath).ToLower();
                string[] imageExtensions = { ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff" };

                // 检查扩展名是否有效
                if (!imageExtensions.Contains(extension))
                    return false;

                // 检查文件是否存在
                return File.Exists(filePath);
            }
            catch (Exception)
            {
                // 捕获任何异常（如ArgumentException等），返回false
                return false;
            }
        }

        private string GetCellValueAsString(Excel.Range cell)
        {
            try
            {
                // 检查是否是日期类型单元格
                if (IsDateCell(cell))
                {
                    // 尝试获取格式化文本
                    string formattedText = cell.Text?.ToString()?.Trim();

                    // 如果显示为 "########" 或空，则使用原始值
                    if (string.IsNullOrWhiteSpace(formattedText) ||
                        formattedText.Contains("####"))
                    {
                        // 使用原始值并尝试转换为日期
                        if (cell.Value2 != null)
                        {
                            // 尝试解析为double（OADate）
                            if (double.TryParse(cell.Value2.ToString(), out double oaDate))
                            {
                                try
                                {
                                    DateTime date = DateTime.FromOADate(oaDate);
                                    return date.ToString("yyyy/MM/dd");
                                }
                                catch
                                {
                                    // 转换失败，返回原始值的字符串表示
                                    return cell.Value2.ToString();
                                }
                            }
                            else
                            {
                                // 不是数字，返回原始值的字符串表示
                                return cell.Value2.ToString();
                            }
                        }
                        else
                        {
                            return string.Empty;
                        }
                    }

                    // 返回Excel显示的格式化文本
                    return formattedText;
                }

                // 非日期单元格直接返回文本值
                return cell.Text?.ToString()?.Trim() ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        // 判断单元格是否为日期类型
        private bool IsDateCell(Excel.Range cell)
        {
            try
            {
                // 检查NumberFormat是否包含日期格式字符
                string format = cell.NumberFormat?.ToString() ?? "";

                // 检查常见日期格式标识符
                bool isDateFormat = format.Contains("y") ||
                                   format.Contains("m") ||
                                   format.Contains("d") ||
                                   format.Contains("h") ||
                                   format.Contains("s");

                // 检查值类型
                if (cell.Value2 != null)
                {
                    try
                    {
                        // 尝试解析为OADate（Excel日期值）
                        double oaDate = Convert.ToDouble(cell.Value2);

                        // Excel日期范围：1900年1月1日到9999年12月31日
                        if (oaDate > 0 && oaDate < 2958465.99999)
                        {
                            return true;
                        }
                    }
                    catch
                    {
                        // 不是数字，可能不是日期
                    }
                }

                return isDateFormat;
            }
            catch
            {
                return false;
            }
        }

        // 更新状态标签（线程安全）
        private void UpdateStatusLabel(string message)
        {
            if (result_doc_label.InvokeRequired)
            {
                result_doc_label.Invoke(new Action(() => UpdateStatusLabel(message)));
            }
            else
            {
                result_doc_label.Text = message;
            }
        }

        // 显示错误信息
        private void ShowErrorMessage(string message)
        {
            if (result_doc_label.InvokeRequired)
            {
                result_doc_label.Invoke(new Action(() => ShowErrorMessage(message)));
            }
            else
            {
                result_doc_label.Text = message;
                MessageBox.Show(message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 替换文档中的所有占位符
        private void ReplacePlaceholders(Word.Document doc, Dictionary<string, string> data, string pattern, ICollection<string> imageHeaders)
        {
            try
            {
                // 安全处理文档正文
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
            catch (Exception ex)
            {
                Debug.WriteLine($"处理StoryRanges时出错: {ex.Message}");
            }

            try
            {
                // 安全处理页眉页脚
                if (doc.Sections != null && doc.Sections.Count > 0)
                {
                    foreach (Word.Section section in doc.Sections)
                    {
                        if (section != null)
                        {
                            // 处理页眉
                            try
                            {
                                if (section.Headers != null)
                                {
                                    Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                                    if (header != null && header.Range != null)
                                    {
                                        ReplacePlaceholdersInRange(header.Range, data, pattern, imageHeaders);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"处理页眉时出错: {ex.Message}");
                            }

                            // 处理页脚
                            try
                            {
                                if (section.Footers != null)
                                {
                                    Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                                    if (footer != null && footer.Range != null)
                                    {
                                        ReplacePlaceholdersInRange(footer.Range, data, pattern, imageHeaders);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"处理页脚时出错: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理Sections时出错: {ex.Message}");
            }

            try
            {
                // 安全处理文本框
                if (doc.Shapes != null && doc.Shapes.Count > 0)
                {
                    foreach (Word.Shape shape in doc.Shapes)
                    {
                        if (shape != null && shape.TextFrame != null && shape.TextFrame.HasText != 0)
                        {
                            Word.Range textRange = shape.TextFrame.TextRange;
                            if (textRange != null)
                            {
                                ReplacePlaceholdersInRange(textRange, data, pattern, imageHeaders);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理Shapes时出错: {ex.Message}");
            }

            try
            {
                // 安全处理表格
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
                                        if (cell != null)
                                        {
                                            Word.Range cellRange = cell.Range;
                                            if (cellRange != null)
                                            {
                                                ReplacePlaceholdersInRange(cellRange, data, pattern, imageHeaders);
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"处理表格单元格({row},{col})时出错: {ex.Message}");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"处理Tables时出错: {ex.Message}");
            }
        }

        // 替换指定区域中的所有占位符
        private void ReplacePlaceholdersInRange(Word.Range range, Dictionary<string, string> data, string pattern, ICollection<string> imageHeaders)
        {
            if (range == null) return;

            string originalText = range.Text;

            foreach (var kvp in data)
            {
                if (imageHeaders.Contains(kvp.Key)) continue;

                string placeholder = pattern.Replace("列名", kvp.Key);
                string replacement = kvp.Value
                    .Replace("\n", "\v")  // Excel换行符→Word换行符
                    .Replace("^p", "\v")  // 兼容旧格式
                    .Replace("^l", "\v"); // 兼容旧格式

                if (originalText.Contains(placeholder))
                {
                    try
                    {
                        range.Find.ClearFormatting();
                        range.Find.Replacement.ClearFormatting();

                        object replaceAll = Word.WdReplace.wdReplaceAll;
                        object wrap = Word.WdFindWrap.wdFindContinue;

                        bool found = range.Find.Execute(
                            FindText: placeholder,
                            ReplaceWith: replacement,
                            Replace: replaceAll,
                            Wrap: wrap
                        );

                        if (!found)
                        {
                            Debug.WriteLine($"未找到占位符: {placeholder}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"替换占位符 {placeholder} 时出错: {ex.Message}");
                    }
                }
            }
        }


        // 在图片占位符位置插入图片
        private void InsertPictureAtPlaceholder(Word.Document doc, string placeholder, string imagePath,
                                             float? heightCm, float? widthCm, bool lockAspectRatio)
        {
            try
            {
                if (doc == null) return;

                // 1. 遍历所有 story ranges（包括正文、页眉、页脚、脚注等）
                if (doc.StoryRanges != null)
                {
                    foreach (Word.Range storyRange in doc.StoryRanges)
                    {
                        Word.Range current = storyRange;
                        while (current != null)
                        {
                            InsertPictureInRange(current, placeholder, imagePath, heightCm, widthCm, lockAspectRatio);
                            try
                            {
                                current = current.NextStoryRange;
                            }
                            catch
                            {
                                current = null;
                            }
                        }
                    }
                }

                // 2. 遍历 Shapes（文本框内的占位符）
                if (doc.Shapes != null && doc.Shapes.Count > 0)
                {
                    foreach (Word.Shape shape in doc.Shapes)
                    {
                        try
                        {
                            if (shape != null && shape.TextFrame != null && shape.TextFrame.HasText != 0)
                            {
                                InsertPictureInRange(shape.TextFrame.TextRange, placeholder, imagePath, heightCm, widthCm, lockAspectRatio);
                            }
                        }
                        catch (Exception exShape)
                        {
                            Debug.WriteLine($"处理 Shape 时出错: {exShape.Message}");
                        }
                    }
                }

                // 3. 遍历表格单元格（表格内文本）
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
                                    {
                                        InsertPictureInRange(cell.Range, placeholder, imagePath, heightCm, widthCm, lockAspectRatio);
                                    }
                                }
                                catch (Exception exCell)
                                {
                                    Debug.WriteLine($"处理表格单元格({r},{c})时出错: {exCell.Message}");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"插入图片时出错: {ex.Message}");
            }
        }

        // 在指定 Range 中查找所有占位符并按顺序插入图片（辅助方法）
        private void InsertPictureInRange(Word.Range range, string placeholder, string imagePath,
                                          float? heightCm, float? widthCm, bool lockAspectRatio)
        {
            if (range == null) return;

            try
            {
                // 复制范围以免修改传入范围的边界
                Word.Range searchRange = range.Duplicate;
                searchRange.Find.ClearFormatting();
                searchRange.Find.Replacement.ClearFormatting();

                // 配置查找
                searchRange.Find.Text = placeholder;
                searchRange.Find.Forward = true;
                searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;
                searchRange.Find.MatchCase = true;
                searchRange.Find.MatchWholeWord = true;

                while (searchRange.Find.Execute())
                {
                    // 在找到的位置创建一个独立的 Range 表示找到的文本
                    Word.Range foundRange = range.Document.Range(searchRange.Start, searchRange.End);

                    // 删除占位符文本
                    foundRange.Text = string.Empty;

                    // 在该位置插入图片（嵌入文档）
                    Word.InlineShape insertedShape = foundRange.InlineShapes.AddPicture(imagePath, LinkToFile: false, SaveWithDocument: true);

                    // 设置尺寸
                    SetImageSize(insertedShape, heightCm, widthCm, lockAspectRatio);

                    // 将搜索起点推进到新插入图片之后（避免再次匹配到同一位置）
                    int newStart = insertedShape.Range.End;
                    int rangeEnd = range.End;
                    if (newStart >= rangeEnd)
                    {
                        // 已到段落或区域末尾，退出循环
                        break;
                    }

                    // 重新设置搜索范围并保持查找设置
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
            catch (Exception ex)
            {
                Debug.WriteLine($"InsertPictureInRange 错误: {ex.Message}");
            }
        }

        // 设置图片尺寸
        private void SetImageSize(Word.InlineShape shape, float? heightCm, float? widthCm, bool lockAspectRatio)
        {
            try
            {
                // 如果是原尺寸插入，直接返回
                if (heightCm == null && widthCm == null)
                {
                    return; // 保持原始尺寸
                }

                // 保存原始尺寸（以磅为单位）
                float originalWidth = shape.Width;
                float originalHeight = shape.Height;
                float aspectRatio = originalWidth / originalHeight;

                // 将厘米转换为磅（1厘米 = 28.35磅）
                const float cmToPoints = 28.35f;
                float? heightPoints = heightCm * cmToPoints;
                float? widthPoints = widthCm * cmToPoints;

                // 处理锁定比例的情况
                if (lockAspectRatio)
                {
                    // 手动计算保持比例的尺寸（兼容.doc格式）
                    float? finalHeight = heightPoints;
                    float? finalWidth = widthPoints;

                    // 优先使用高度设置
                    if (heightPoints.HasValue)
                    {
                        finalHeight = heightPoints.Value;
                        finalWidth = heightPoints.Value * aspectRatio;
                    }
                    // 如果高度未设置但宽度设置了
                    else if (widthPoints.HasValue)
                    {
                        finalWidth = widthPoints.Value;
                        finalHeight = widthPoints.Value / aspectRatio;
                    }

                    // 手动设置两个维度（解决.doc格式问题）
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    if (finalHeight.HasValue) shape.Height = finalHeight.Value;
                    if (finalWidth.HasValue) shape.Width = finalWidth.Value;
                }
                // 不锁定比例的情况
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
                // 出错时恢复默认设置
                try
                {
                    shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                }
                catch { }
            }
        }

        // 清理文件名中的非法字符
        private string SanitizeFilename(string filename)
        {
            var invalidChars = Path.GetInvalidFileNameChars();
            return new string(filename
                .Where(c => !invalidChars.Contains(c))
                .ToArray());
        }

        private void docReset_button_Click(object sender, EventArgs e)
        {
            sheets_name_comboBox.Text = Globals.ThisAddIn.Application.ActiveSheet.Name;
            docModel_label.Text = string.Empty;
            docGenerated_label.Text = string.Empty;
            result_doc_label.Text = string.Empty;
            pictureSize_checkBox.Checked = true; // 默认锁定比例
            pictureOriginal_comboBox.SelectedIndex = 0; // 默认图片原始尺寸
            height_textBox.Text = string.Empty;  // 清空高度
            width_textBox.Text = string.Empty;   // 清空宽度
        }

        private void pictureSize_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            maintainAspectRatio = pictureSize_checkBox.Checked;
            width_textBox.ReadOnly = maintainAspectRatio;

            // 锁定比例时清空宽度输入
            if (maintainAspectRatio)
            {
                width_textBox.Text = string.Empty;
            }
        }

        // 验证并保存尺寸设置
        private bool ValidateAndSaveDimensions()
        {
            // 重置尺寸设置
            imageHeightCm = null;
            imageWidthCm = null;

            // 解析高度
            if (!string.IsNullOrWhiteSpace(height_textBox.Text))
            {
                if (float.TryParse(height_textBox.Text, out float height) && height > 0)
                {
                    imageHeightCm = height;
                }
                else
                {
                    ShowErrorMessage("高度必须是大于0的数字");
                    return false;
                }
            }

            // 解析宽度
            if (!string.IsNullOrWhiteSpace(width_textBox.Text))
            {
                if (float.TryParse(width_textBox.Text, out float width) && width > 0)
                {
                    imageWidthCm = width;
                }
                else
                {
                    ShowErrorMessage("宽度必须是大于0的数字");
                    return false;
                }
            }

            // 检查锁定比例时的设置
            if (maintainAspectRatio && !imageHeightCm.HasValue && !imageWidthCm.HasValue)
            {
                ShowErrorMessage("锁定比例时至少需要设置高度或宽度");
                return false;
            }

            return true;
        }

        private void docQuit_button_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
}