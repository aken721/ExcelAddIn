using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using Word= Microsoft.Office.Interop.Word;

namespace ExcelAddIn
{
    public partial class Form10 : Form
    {
        public Form10()
        {
            InitializeComponent();
        }



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
            placeholder_comboBox.SelectedIndex=3;
        }

        private void model_select_button_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Word Documents (*.docx)|*.docx;*.doc";
            openFileDialog1.Title = "选择Word模板文件";
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

            // 重置状态标签
            result_doc_label.Text = "准备生成文档...";
            Application.DoEvents();

            // 获取窗体控件值
            string sheetName = sheets_name_comboBox.SelectedItem?.ToString();
            string placeholderPattern = placeholder_comboBox.SelectedItem?.ToString();
            string templatePath = docModel_label.Text;
            string outputFolder = docGenerated_label.Text;

            // 在后台执行生成任务，避免阻塞UI
            await Task.Run(() => GenerateDocuments(sheetName, placeholderPattern, templatePath, outputFolder));
        }

        private void GenerateDocuments(string sheetName,string placeholderPattern,string templatePath,string outputFolder)
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
                int totalDocuments = rowCount - 1; // 总文档数（排除标题行）

                if (totalDocuments < 1)
                {
                    ShowErrorMessage("数据表中没有可处理的数据行");
                    return;
                }

                var headers = new List<string>();
                for (int col = 2; col <= columnCount; col++)
                {
                    var cellValue = (usedRange.Cells[1, col] as Excel.Range).Value2;
                    headers.Add(cellValue?.ToString() ?? $"列{col}");
                }

                // 启动Word应用程序
                var wordApp = new Word.Application { Visible = false };
                int successCount = 0;
                int errorCount = 0;

                // 处理数据行（从第二行开始）
                for (int row = 2; row <= rowCount; row++)
                {
                    try
                    {
                        // 更新状态标签
                        int currentDocument = row - 1;
                        UpdateStatusLabel($"正在生成第 {currentDocument}/{totalDocuments} 个文档...");

                        // 获取文件名（第一列）
                        var fileNameCell = usedRange.Cells[row, 1] as Excel.Range;
                        string fileName = fileNameCell.Value2?.ToString() ?? $"文档_{currentDocument}";

                        // 处理文件名中的非法字符
                        fileName = SanitizeFilename(fileName);
                        if (!fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
                            fileName += ".docx";

                        var rowData = new Dictionary<string, string>();

                        // 收集数据（从第二列开始）
                        for (int col = 2; col <= columnCount; col++)
                        {
                            var cell = usedRange.Cells[row, col] as Excel.Range;

                            // 获取单元格的原始值和格式
                            object value = cell.Value2;
                            string numberFormat = cell.NumberFormat?.ToString() ?? "";

                            // 如果是日期类型
                            if (value is double && IsDateCell(cell))
                            {
                                DateTime dateValue = DateTime.FromOADate((double)value);

                                // 使用Excel的格式字符串格式化日期
                                if (!string.IsNullOrEmpty(numberFormat))
                                {
                                    try
                                    {
                                        // 解析Excel格式字符串并转换为.NET格式
                                        string netFormat = ConvertExcelFormatToNetFormat(numberFormat);

                                        // 格式化日期
                                        string formattedDate = dateValue.ToString(netFormat);

                                        rowData[headers[col - 2]] = formattedDate;
                                    }
                                    catch
                                    {
                                        // 如果格式化失败，使用默认格式
                                        rowData[headers[col - 2]] = dateValue.ToString("yyyy/MM/dd");
                                    }
                                }
                                else
                                {
                                    // 没有格式字符串，使用Excel显示的文本值
                                    rowData[headers[col - 2]] = cell.Text.ToString();
                                }
                            }
                            else
                            {
                                // 非日期类型，正常处理
                                rowData[headers[col - 2]] = value?.ToString() ?? string.Empty;
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
                            ReplacePlaceholders(wordDoc, rowData, placeholderPattern);
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

        // 将Excel格式字符串转换为.NET格式字符串
        private string ConvertExcelFormatToNetFormat(string excelFormat)
        {
            // 创建映射关系：Excel格式代码 -> .NET格式代码
            var formatMap = new Dictionary<string, string>
    {
        {"yyyy", "yyyy"}, // 四位年份
        {"yy", "yy"},     // 两位年份
        {"mmmm", "MMMM"}, // 月份全名
        {"mmm", "MMM"},   // 月份缩写
        {"mm", "MM"},     // 两位月份
        {"m", "M"},       // 月份（无前导零）
        {"dddd", "dddd"}, // 星期全名
        {"ddd", "ddd"},   // 星期缩写
        {"dd", "dd"},     // 两位日期
        {"d", "d"},       // 日期（无前导零）
        {"hh", "hh"},     // 12小时制小时
        {"HH", "HH"},     // 24小时制小时
        {"mm", "mm"},     // 分钟（注意：分钟和月份都使用mm，但在上下文中区分）
        {"ss", "ss"},     // 秒
        {"tt", "tt"}      // AM/PM
    };

            // 特殊处理：区分月份和分钟
            // 在日期格式中，mm通常表示月份，在时间格式中表示分钟
            // 这里我们假设如果格式字符串包含"h"或"t"，则mm表示分钟
            bool isTimeFormat = excelFormat.Contains("h") || excelFormat.Contains("t");

            // 替换所有格式代码
            string netFormat = excelFormat;
            foreach (var mapping in formatMap)
            {
                // 特殊处理：如果是时间格式且mapping.Key是"mm"，则替换为分钟格式
                if (isTimeFormat && mapping.Key == "mm")
                {
                    netFormat = netFormat.Replace("mm", "mm");
                }
                else
                {
                    netFormat = netFormat.Replace(mapping.Key, mapping.Value);
                }
            }

            return netFormat;
        }

        // 判断单元格是否为日期类型
        private bool IsDateCell(Excel.Range cell)
        {
            try
            {
                // 检查NumberFormat是否包含日期格式字符
                string format = cell.NumberFormat?.ToString() ?? "";
                return format.Contains("y") || format.Contains("m") || format.Contains("d");
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
        private void ReplacePlaceholders(Word.Document doc, Dictionary<string, string> data, string pattern)
        {
            // 处理文档正文
            foreach (Word.Range range in doc.StoryRanges)
            {
                ReplacePlaceholdersInRange(range, data, pattern);
            }

            // 处理页眉
            foreach (Word.Section section in doc.Sections)
            {
                ReplacePlaceholdersInRange(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range, data, pattern);
            }

            // 处理页脚
            foreach (Word.Section section in doc.Sections)
            {
                ReplacePlaceholdersInRange(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range, data, pattern);
            }

            // 处理文本框
            foreach (Word.Shape shape in doc.Shapes)
            {
                if (shape.TextFrame.HasText != 0)
                {
                    ReplacePlaceholdersInRange(shape.TextFrame.TextRange, data, pattern);
                }
            }

            // 处理表格
            foreach (Word.Table table in doc.Tables)
            {
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        Word.Cell cell = table.Cell(row, col);
                        ReplacePlaceholdersInRange(cell.Range, data, pattern);
                    }
                }
            }
        }

        // 替换指定区域中的所有占位符
        private void ReplacePlaceholdersInRange(Word.Range range, Dictionary<string, string> data, string pattern)
        {
            // 保存原始文本
            string originalText = range.Text;

            foreach (var kvp in data)
            {
                // 根据模式创建占位符
                string placeholder = pattern.Replace("列名", kvp.Key);

                if (originalText.Contains(placeholder))
                {
                    range.Find.ClearFormatting();
                    range.Find.Text = placeholder;
                    range.Find.Replacement.ClearFormatting();
                    range.Find.Replacement.Text = kvp.Value;

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object wrap = Word.WdFindWrap.wdFindContinue;

                    range.Find.Execute(
                        FindText: placeholder,
                        ReplaceWith: kvp.Value,
                        Replace: replaceAll,
                        Wrap: wrap
                    );
                }
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
        }
    }
}
