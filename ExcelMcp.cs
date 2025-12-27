using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    /// <summary>
    /// Excel MCP Agent Core - 基于Microsoft.Office.Interop.Excel提供Excel文件操作的完整功能集
    /// (已为Web API改造，使用文件名作为ID管理状态)
    /// </summary>
    public class ExcelMcp : IDisposable
    {
        private readonly string _excelFilesPath;
        private readonly Application _excelApp;
        private bool _disposed = false;
        private readonly ConcurrentDictionary<string, Workbook> _openWorkbooks = new ConcurrentDictionary<string, Workbook>();

        public ExcelMcp(string excelFilesPath = "./excel_files")
        {
            _excelFilesPath = Path.GetFullPath(excelFilesPath);

            if (!Directory.Exists(_excelFilesPath))
            {
                Directory.CreateDirectory(_excelFilesPath);
            }

            _excelApp = new Application();
            _excelApp.Visible = false;
            _excelApp.DisplayAlerts = false;
        }

        #region Internal Helpers
        private Workbook GetWorkbookById(string fileName)
        {
            if (!_openWorkbooks.TryGetValue(fileName, out var workbook))
            {
                throw new ArgumentException($"工作簿 '{fileName}' 未打开或不存在。请先调用OpenWorkbook。");
            }
            return workbook;
        }

        private Worksheet GetWorksheetById(string fileName, string sheetName)
        {
            var workbook = GetWorkbookById(fileName);
            try
            {
                return (Worksheet)workbook.Worksheets[sheetName];
            }
            catch
            {
                Marshal.ReleaseComObject(workbook);
                throw new ArgumentException($"在工作簿'{fileName}'中，工作表 '{sheetName}' 不存在。");
            }
        }
        #endregion

        #region 工作簿操作 (Workbook Operations)

        public string CreateWorkbook(string fileName, string sheetName = "Sheet1")
        {
            if (_openWorkbooks.ContainsKey(fileName))
            {
                throw new ArgumentException($"名为 '{fileName}' 的工作簿已处于打开状态。");
            }
            var filePath = Path.Combine(_excelFilesPath, fileName);
            if (File.Exists(filePath))
            {
                throw new ArgumentException($"名为 '{fileName}' 的文件已存在于磁盘上。");
            }

            Workbook workbook = null;
            Worksheet worksheet = null;
            try
            {
                workbook = _excelApp.Workbooks.Add();
                worksheet = (Worksheet)workbook.Worksheets[1];
                worksheet.Name = sheetName;
                workbook.SaveAs(filePath);
                return filePath;
            }
            finally
            {
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                {
                    workbook.Close(false);
                    Marshal.ReleaseComObject(workbook);
                }
            }
        }

        public string OpenWorkbook(string fileName)
        {
            if (_openWorkbooks.ContainsKey(fileName))
            {
                return fileName; // Already open
            }

            var filePath = Path.Combine(_excelFilesPath, fileName);
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"文件不存在: {filePath}");
            }

            var workbook = _excelApp.Workbooks.Open(filePath);
            _openWorkbooks.TryAdd(fileName, workbook);
            return fileName;
        }

        public void CloseWorkbook(string fileName)
        {
            if (_openWorkbooks.TryRemove(fileName, out var workbook))
            {
                workbook.Close(true); // Save changes on close
                Marshal.ReleaseComObject(workbook);
            }
        }

        public void SaveWorkbook(string fileName)
        {
            var workbook = GetWorkbookById(fileName);
            workbook.Save();
        }

        public void SaveWorkbookAs(string fileName, string newFileName)
        {
            var workbook = GetWorkbookById(fileName);
            var newFilePath = Path.Combine(_excelFilesPath, newFileName);
            workbook.SaveAs(newFilePath);

            if (_openWorkbooks.TryRemove(fileName, out var oldWorkbook))
            {
                _openWorkbooks.TryAdd(newFileName, oldWorkbook);
            }
        }

        #endregion

        #region 工作表操作 (Worksheet Operations)

        public string CreateWorksheet(string fileName, string sheetName)
        {
            var workbook = GetWorkbookById(fileName);
            Worksheet worksheet = (Worksheet)workbook.Worksheets.Add();
            worksheet.Name = sheetName;
            Marshal.ReleaseComObject(worksheet);
            return sheetName;
        }

        public void RenameWorksheet(string fileName, string oldSheetName, string newSheetName)
        {
            Worksheet worksheet = GetWorksheetById(fileName, oldSheetName);
            worksheet.Name = newSheetName;
            Marshal.ReleaseComObject(worksheet);
        }

        public void DeleteWorksheet(string fileName, string sheetName)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            worksheet.Delete();
            Marshal.ReleaseComObject(worksheet);
        }

        public List<string> GetWorksheetNames(string fileName)
        {
            var workbook = GetWorkbookById(fileName);
            var names = new List<string>();
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                names.Add(worksheet.Name);
                Marshal.ReleaseComObject(worksheet);
            }
            return names;
        }

        /// <summary>
        /// 激活指定的工作表（使其成为当前活跃工作表）
        /// </summary>
        /// <param name="fileName">工作簿文件名</param>
        /// <param name="sheetName">要激活的工作表名称</param>
        /// <returns>被激活的工作表名称</returns>
        public string ActivateWorksheet(string fileName, string sheetName)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            worksheet.Activate();
            Marshal.ReleaseComObject(worksheet);
            return sheetName;
        }

        /// <summary>
        /// 获取当前活跃的工作表名称
        /// </summary>
        /// <param name="fileName">工作簿文件名</param>
        /// <returns>当前活跃工作表的名称</returns>
        public string GetActiveWorksheetName(string fileName)
        {
            var workbook = GetWorkbookById(fileName);
            Worksheet activeSheet = (Worksheet)workbook.ActiveSheet;
            string name = activeSheet.Name;
            Marshal.ReleaseComObject(activeSheet);
            return name;
        }

        #endregion

        #region 数据操作 (Data Operations)

        public void SetCellValue(string fileName, string sheetName, int row, int column, object value)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = (Range)worksheet.Cells[row, column];
            cell.Value = value;
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
        }

        public object GetCellValue(string fileName, string sheetName, int row, int column)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = (Range)worksheet.Cells[row, column];
            object value = cell.Value;
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
            return value;
        }

        public void SetRangeValues(string fileName, string sheetName, string rangeAddress, object[,] data)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            range.Value = data;
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public object[,] GetRangeValues(string fileName, string sheetName, string rangeAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            object[,] result = (object[,])range.Value;
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
            return result;
        }

        public void SetFormula(string fileName, string sheetName, string cellAddress, string formula)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = worksheet.get_Range(cellAddress);
            cell.Formula = formula;
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
        }

        public string GetFormula(string fileName, string sheetName, string cellAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = worksheet.get_Range(cellAddress);
            string formula = cell.Formula as string;
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
            return formula ?? "";
        }

        #endregion

        #region 格式化操作 (Formatting Operations)

        public void SetCellFormat(string fileName, string sheetName, string rangeAddress,
            string fontColor = null, string backgroundColor = null, int? fontSize = null,
            bool? bold = null, bool? italic = null,
            string horizontalAlignment = null, string verticalAlignment = null)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);

            // 字体颜色
            if (!string.IsNullOrEmpty(fontColor))
            {
                range.Font.Color = ParseColor(fontColor);
            }

            // 背景色
            if (!string.IsNullOrEmpty(backgroundColor))
            {
                range.Interior.Color = ParseColor(backgroundColor);
            }

            // 字号
            if (fontSize.HasValue)
            {
                range.Font.Size = fontSize.Value;
            }

            // 加粗
            if (bold.HasValue)
            {
                range.Font.Bold = bold.Value;
            }

            // 斜体
            if (italic.HasValue)
            {
                range.Font.Italic = italic.Value;
            }

            // 水平对齐
            if (!string.IsNullOrEmpty(horizontalAlignment))
            {
                range.HorizontalAlignment = horizontalAlignment.ToLower() switch
                {
                    "left" => XlHAlign.xlHAlignLeft,
                    "center" => XlHAlign.xlHAlignCenter,
                    "right" => XlHAlign.xlHAlignRight,
                    _ => XlHAlign.xlHAlignGeneral
                };
            }

            // 垂直对齐
            if (!string.IsNullOrEmpty(verticalAlignment))
            {
                range.VerticalAlignment = verticalAlignment.ToLower() switch
                {
                    "top" => XlVAlign.xlVAlignTop,
                    "center" => XlVAlign.xlVAlignCenter,
                    "bottom" => XlVAlign.xlVAlignBottom,
                    _ => XlVAlign.xlVAlignCenter
                };
            }

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public void SetBorder(string fileName, string sheetName, string rangeAddress,
            string borderType, string lineStyle = "continuous")
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);

            var xlLineStyle = lineStyle.ToLower() switch
            {
                "dash" => XlLineStyle.xlDash,
                "dot" => XlLineStyle.xlDot,
                _ => XlLineStyle.xlContinuous
            };

            switch (borderType.ToLower())
            {
                case "all":
                    range.Borders.LineStyle = xlLineStyle;
                    break;
                case "outline":
                    range.BorderAround(xlLineStyle);
                    break;
                case "horizontal":
                    range.Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = xlLineStyle;
                    break;
                case "vertical":
                    range.Borders[XlBordersIndex.xlInsideVertical].LineStyle = xlLineStyle;
                    break;
            }

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public void MergeCells(string fileName, string sheetName, string rangeAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            range.Merge();
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public void UnmergeCells(string fileName, string sheetName, string rangeAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            range.UnMerge();
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public void SetRowHeight(string fileName, string sheetName, int rowNumber, double height)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range row = (Range)worksheet.Rows[rowNumber];
            row.RowHeight = height;
            Marshal.ReleaseComObject(row);
            Marshal.ReleaseComObject(worksheet);
        }

        public void SetColumnWidth(string fileName, string sheetName, int columnNumber, double width)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range column = (Range)worksheet.Columns[columnNumber];
            column.ColumnWidth = width;
            Marshal.ReleaseComObject(column);
            Marshal.ReleaseComObject(worksheet);
        }

        // 辅助方法：解析颜色
        private int ParseColor(string colorStr)
        {
            // 支持颜色名称和十六进制颜色
            if (colorStr.StartsWith("#"))
            {
                // 十六进制颜色 #RRGGBB
                var hex = colorStr.Substring(1);
                var r = Convert.ToInt32(hex.Substring(0, 2), 16);
                var g = Convert.ToInt32(hex.Substring(2, 2), 16);
                var b = Convert.ToInt32(hex.Substring(4, 2), 16);
                return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
            }
            else
            {
                // 颜色名称
                var color = colorStr.ToLower() switch
                {
                    "红色" or "red" => System.Drawing.Color.Red,
                    "绿色" or "green" => System.Drawing.Color.Green,
                    "蓝色" or "blue" => System.Drawing.Color.Blue,
                    "黄色" or "yellow" => System.Drawing.Color.Yellow,
                    "橙色" or "orange" => System.Drawing.Color.Orange,
                    "紫色" or "purple" => System.Drawing.Color.Purple,
                    "黑色" or "black" => System.Drawing.Color.Black,
                    "白色" or "white" => System.Drawing.Color.White,
                    "灰色" or "gray" => System.Drawing.Color.Gray,
                    _ => System.Drawing.Color.Black
                };
                return System.Drawing.ColorTranslator.ToOle(color);
            }
        }

        #endregion

        #region 行列操作 (Row/Column Operations)

        public void InsertRows(string fileName, string sheetName, int rowIndex, int count = 1)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range row = (Range)worksheet.Rows[rowIndex];
            for (int i = 0; i < count; i++)
            {
                row.Insert(XlInsertShiftDirection.xlShiftDown, Type.Missing);
            }
            Marshal.ReleaseComObject(row);
            Marshal.ReleaseComObject(worksheet);
        }

        public void InsertColumns(string fileName, string sheetName, int columnIndex, int count = 1)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range column = (Range)worksheet.Columns[columnIndex];
            for (int i = 0; i < count; i++)
            {
                column.Insert(XlInsertShiftDirection.xlShiftToRight, Type.Missing);
            }
            Marshal.ReleaseComObject(column);
            Marshal.ReleaseComObject(worksheet);
        }

        public void DeleteRows(string fileName, string sheetName, int rowIndex, int count = 1)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range rows = (Range)worksheet.Range[worksheet.Rows[rowIndex], worksheet.Rows[rowIndex + count - 1]];
            rows.Delete(XlDeleteShiftDirection.xlShiftUp);
            Marshal.ReleaseComObject(rows);
            Marshal.ReleaseComObject(worksheet);
        }

        public void DeleteColumns(string fileName, string sheetName, int columnIndex, int count = 1)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range columns = (Range)worksheet.Range[worksheet.Columns[columnIndex], worksheet.Columns[columnIndex + count - 1]];
            columns.Delete(XlDeleteShiftDirection.xlShiftToLeft);
            Marshal.ReleaseComObject(columns);
            Marshal.ReleaseComObject(worksheet);
        }

        #endregion

        #region 复制操作 (Copy Operations)

        public void CopyWorksheet(string fileName, string sourceSheetName, string targetSheetName)
        {
            var workbook = GetWorkbookById(fileName);
            Worksheet sourceSheet = null;
            Worksheet newSheet = null;

            try
            {
                sourceSheet = (Worksheet)workbook.Worksheets[sourceSheetName];
                sourceSheet.Copy(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]);
                newSheet = (Worksheet)workbook.Worksheets[workbook.Worksheets.Count];
                newSheet.Name = targetSheetName;
            }
            finally
            {
                if (sourceSheet != null) Marshal.ReleaseComObject(sourceSheet);
                if (newSheet != null) Marshal.ReleaseComObject(newSheet);
                Marshal.ReleaseComObject(workbook);
            }
        }

        public void CopyRange(string fileName, string sheetName, string sourceRange, string targetRange)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range source = worksheet.get_Range(sourceRange);
            Range target = worksheet.get_Range(targetRange);
            source.Copy(target);
            Marshal.ReleaseComObject(source);
            Marshal.ReleaseComObject(target);
            Marshal.ReleaseComObject(worksheet);
        }

        public void ClearRange(string fileName, string sheetName, string rangeAddress, string clearType = "all")
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);

            switch (clearType.ToLower())
            {
                case "contents":
                    range.ClearContents();
                    break;
                case "formats":
                    range.ClearFormats();
                    break;
                case "all":
                default:
                    range.Clear();
                    break;
            }

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        #endregion

        #region 工作簿元数据 (Workbook Metadata)

        public string GetWorkbookMetadata(string fileName, bool includeRanges = false)
        {
            var workbook = GetWorkbookById(fileName);
            var metadata = new System.Text.StringBuilder();

            metadata.AppendLine($"工作簿名称: {workbook.Name}");
            metadata.AppendLine($"工作表数量: {workbook.Worksheets.Count}");
            metadata.AppendLine($"完整路径: {workbook.FullName}");
            metadata.AppendLine("工作表列表:");

            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                metadata.AppendLine($"  - {worksheet.Name}");

                if (includeRanges)
                {
                    Range usedRange = worksheet.UsedRange;
                    metadata.AppendLine($"    已使用范围: {usedRange.Address}");
                    metadata.AppendLine($"    行数: {usedRange.Rows.Count}, 列数: {usedRange.Columns.Count}");
                    Marshal.ReleaseComObject(usedRange);
                }

                Marshal.ReleaseComObject(worksheet);
            }

            Marshal.ReleaseComObject(workbook);
            return metadata.ToString();
        }

        /// <summary>
        /// 获取当前Excel应用程序信息（活跃工作簿、工作表等）
        /// 注意：此方法依赖ThisAddIn.app，仅在Excel插件环境中可用
        /// </summary>
        /// <returns>当前Excel信息的字符串描述</returns>
        public string GetCurrentExcelInfo()
        {
            var sb = new System.Text.StringBuilder();
            
            if (ThisAddIn.app == null)
            {
                return "Excel应用程序未初始化";
            }

            try
            {
                sb.AppendLine($"Excel版本: {ThisAddIn.app.Version}");
                
                if (ThisAddIn.app.Workbooks.Count > 0)
                {
                    sb.AppendLine($"打开的工作簿数量: {ThisAddIn.app.Workbooks.Count}");
                    
                    if (ThisAddIn.app.ActiveWorkbook != null)
                    {
                        var activeWb = ThisAddIn.app.ActiveWorkbook;
                        sb.AppendLine($"当前活跃工作簿: {activeWb.Name}");
                        sb.AppendLine($"工作表数量: {activeWb.Worksheets.Count}");
                        
                        sb.AppendLine("工作表列表:");
                        foreach (Worksheet ws in activeWb.Worksheets)
                        {
                            sb.AppendLine($"  - {ws.Name}");
                        }
                        
                        if (ThisAddIn.app.ActiveSheet != null)
                        {
                            Worksheet activeSheet = ThisAddIn.app.ActiveSheet;
                            sb.AppendLine($"当前活跃工作表: {activeSheet.Name}");
                        }
                    }
                }
                else
                {
                    sb.AppendLine("没有打开的工作簿");
                }
            }
            catch (Exception ex)
            {
                sb.AppendLine($"获取Excel信息时出错: {ex.Message}");
            }

            return sb.ToString();
        }

        /// <summary>
        /// 获取当前选中的单元格地址
        /// 注意：此方法依赖ThisAddIn.app，仅在Excel插件环境中可用
        /// </summary>
        /// <returns>当前选中单元格的地址</returns>
        public string GetCurrentSelection()
        {
            if (ThisAddIn.app == null)
            {
                return "Excel应用程序未初始化";
            }

            try
            {
                Range selection = ThisAddIn.app.Selection as Range;
                if (selection != null)
                {
                    string address = selection.Address;
                    string sheetName = ((Worksheet)selection.Worksheet).Name;
                    return $"当前选中: {sheetName}!{address}";
                }
                return "没有选中的单元格";
            }
            catch (Exception ex)
            {
                return $"获取选中单元格时出错: {ex.Message}";
            }
        }

        #endregion

        #region 数据验证 (Data Validation)

        public void SetDataValidation(string fileName, string sheetName, string rangeAddress,
            string validationType, string operatorType = "between",
            string formula1 = null, string formula2 = null,
            string inputMessage = null, string errorMessage = null,
            bool showInput = true, bool showError = true)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);

            // 删除现有验证
            range.Validation.Delete();

            // 设置验证类型
            XlDVType xlType = validationType.ToLower() switch
            {
                "whole" => XlDVType.xlValidateWholeNumber,
                "decimal" => XlDVType.xlValidateDecimal,
                "list" => XlDVType.xlValidateList,
                "date" => XlDVType.xlValidateDate,
                "time" => XlDVType.xlValidateTime,
                "textlength" => XlDVType.xlValidateTextLength,
                "custom" => XlDVType.xlValidateCustom,
                _ => XlDVType.xlValidateInputOnly
            };

            // 设置操作符类型
            XlDVAlertStyle alertStyle = XlDVAlertStyle.xlValidAlertStop;
            XlFormatConditionOperator xlOperator = operatorType.ToLower() switch
            {
                "between" => XlFormatConditionOperator.xlBetween,
                "notbetween" => XlFormatConditionOperator.xlNotBetween,
                "equal" => XlFormatConditionOperator.xlEqual,
                "notequal" => XlFormatConditionOperator.xlNotEqual,
                "greater" => XlFormatConditionOperator.xlGreater,
                "less" => XlFormatConditionOperator.xlLess,
                "greaterorequal" => XlFormatConditionOperator.xlGreaterEqual,
                "lessorequal" => XlFormatConditionOperator.xlLessEqual,
                _ => XlFormatConditionOperator.xlBetween
            };

            // 添加验证
            range.Validation.Add(xlType, alertStyle, xlOperator, formula1, formula2);

            // 设置输入提示
            if (showInput && !string.IsNullOrEmpty(inputMessage))
            {
                range.Validation.IgnoreBlank = true;
                range.Validation.InCellDropdown = true;
                range.Validation.ShowInput = true;
                range.Validation.InputTitle = "输入提示";
                range.Validation.InputMessage = inputMessage;
            }

            // 设置错误提示
            if (showError && !string.IsNullOrEmpty(errorMessage))
            {
                range.Validation.ShowError = true;
                range.Validation.ErrorTitle = "输入错误";
                range.Validation.ErrorMessage = errorMessage;
            }

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public string GetValidationRules(string fileName, string sheetName, string rangeAddress = null)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = string.IsNullOrEmpty(rangeAddress) ? worksheet.UsedRange : worksheet.get_Range(rangeAddress);

            var result = new System.Text.StringBuilder();
            result.AppendLine($"范围 {range.Address} 的数据验证规则:");

            try
            {
                if (range.Validation != null)
                {
                    result.AppendLine($"  类型: {range.Validation.Type}");
                    result.AppendLine($"  公式1: {range.Validation.Formula1}");
                    if (range.Validation.Type == (int)XlDVType.xlValidateList ||
                        range.Validation.Operator == (int)XlFormatConditionOperator.xlBetween)
                    {
                        result.AppendLine($"  公式2: {range.Validation.Formula2}");
                    }
                    result.AppendLine($"  输入提示: {range.Validation.InputMessage}");
                    result.AppendLine($"  错误提示: {range.Validation.ErrorMessage}");
                }
                else
                {
                    result.AppendLine("  无验证规则");
                }
            }
            catch
            {
                result.AppendLine("  无验证规则");
            }

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
            return result.ToString();
        }

        #endregion

        #region 条件格式与数字格式 (Conditional Formatting & Number Format)

        public void SetNumberFormat(string fileName, string sheetName, string rangeAddress, string formatCode)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            range.NumberFormat = formatCode;
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public void ApplyConditionalFormatting(string fileName, string sheetName, string rangeAddress,
            string ruleType, string formula1 = null, string formula2 = null,
            string color1 = null, string color2 = null, string color3 = null)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);

            // 清除现有条件格式
            range.FormatConditions.Delete();

            switch (ruleType.ToLower())
            {
                case "cellvalue":
                    // 单元格值条件
                    FormatCondition condition = (FormatCondition)range.FormatConditions.Add(
                        XlFormatConditionType.xlCellValue,
                        XlFormatConditionOperator.xlGreater,
                        formula1);
                    condition.Interior.Color = ParseColor(color1 ?? "yellow");
                    Marshal.ReleaseComObject(condition);
                    break;

                case "colorscale":
                    // 色阶
                    ColorScale colorScale = (ColorScale)range.FormatConditions.AddColorScale(3);
                    colorScale.ColorScaleCriteria[1].Type = XlConditionValueTypes.xlConditionValueLowestValue;
                    colorScale.ColorScaleCriteria[1].FormatColor.Color = ParseColor(color1 ?? "red");
                    colorScale.ColorScaleCriteria[2].Type = XlConditionValueTypes.xlConditionValuePercentile;
                    colorScale.ColorScaleCriteria[2].Value = 50;
                    colorScale.ColorScaleCriteria[2].FormatColor.Color = ParseColor(color2 ?? "yellow");
                    colorScale.ColorScaleCriteria[3].Type = XlConditionValueTypes.xlConditionValueHighestValue;
                    colorScale.ColorScaleCriteria[3].FormatColor.Color = ParseColor(color3 ?? "green");
                    Marshal.ReleaseComObject(colorScale);
                    break;

                case "databar":
                    // 数据条
                    Databar databar = (Databar)range.FormatConditions.AddDatabar();
                    databar.BarColor.Color = ParseColor(color1 ?? "blue");
                    Marshal.ReleaseComObject(databar);
                    break;

                case "iconset":
                    // 图标集 - AddIconSetCondition会自动应用默认图标集(3个交通灯)
                    IconSetCondition iconSet = (IconSetCondition)range.FormatConditions.AddIconSetCondition();
                    // 默认已经是3个交通灯图标集，无需额外设置
                    Marshal.ReleaseComObject(iconSet);
                    break;

                case "expression":
                    // 使用公式
                    FormatCondition exprCondition = (FormatCondition)range.FormatConditions.Add(
                        XlFormatConditionType.xlExpression,
                        Type.Missing,
                        formula1);
                    exprCondition.Interior.Color = ParseColor(color1 ?? "yellow");
                    Marshal.ReleaseComObject(exprCondition);
                    break;
            }

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        #endregion

        #region 图表操作 (Chart Operations)

        public void CreateChart(string fileName, string sheetName, string chartType, string dataRange,
            string chartPosition, string title = null, int width = 400, int height = 300)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range dataRangeObj = worksheet.get_Range(dataRange);
            Range chartPositionObj = worksheet.get_Range(chartPosition);

            // 创建图表
            ChartObjects chartObjects = (ChartObjects)worksheet.ChartObjects(Type.Missing);
            ChartObject chartObject = chartObjects.Add(
                (double)chartPositionObj.Left,
                (double)chartPositionObj.Top,
                width,
                height);

            Chart chart = chartObject.Chart;

            // 设置图表类型
            XlChartType xlChartType = chartType.ToLower() switch
            {
                "line" => XlChartType.xlLine,
                "bar" => XlChartType.xlBarClustered,
                "column" => XlChartType.xlColumnClustered,
                "pie" => XlChartType.xlPie,
                "scatter" => XlChartType.xlXYScatter,
                "area" => XlChartType.xlArea,
                "radar" => XlChartType.xlRadar,
                "doughnut" => XlChartType.xlDoughnut,
                _ => XlChartType.xlColumnClustered
            };

            chart.ChartType = xlChartType;

            // 设置数据源
            chart.SetSourceData(dataRangeObj);

            // 设置标题
            if (!string.IsNullOrEmpty(title))
            {
                chart.HasTitle = true;
                chart.ChartTitle.Text = title;
            }

            Marshal.ReleaseComObject(dataRangeObj);
            Marshal.ReleaseComObject(chartPositionObj);
            Marshal.ReleaseComObject(chart);
            Marshal.ReleaseComObject(chartObject);
            Marshal.ReleaseComObject(chartObjects);
            Marshal.ReleaseComObject(worksheet);
        }

        #endregion

        #region Excel表格操作 (Excel Table Operations)

        public void CreateTable(string fileName, string sheetName, string rangeAddress, string tableName,
            bool hasHeaders = true, string tableStyle = "TableStyleMedium2")
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);

            // 创建表格
            ListObject table = worksheet.ListObjects.Add(
                XlListObjectSourceType.xlSrcRange,
                range,
                Type.Missing,
                hasHeaders ? XlYesNoGuess.xlYes : XlYesNoGuess.xlNo,
                Type.Missing);

            table.Name = tableName;

            // 设置表格样式
            try
            {
                table.TableStyle = tableStyle;
            }
            catch
            {
                // 如果样式不存在，使用默认样式
                table.TableStyle = "TableStyleMedium2";
            }

            Marshal.ReleaseComObject(table);
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public List<string> GetTableNames(string fileName, string sheetName)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            var tableNames = new List<string>();

            foreach (ListObject table in worksheet.ListObjects)
            {
                tableNames.Add(table.Name);
                Marshal.ReleaseComObject(table);
            }

            Marshal.ReleaseComObject(worksheet);
            return tableNames;
        }

        #endregion

        #region 数据透视表操作 (Pivot Table Operations)

        public void CreatePivotTable(string fileName, string sourceSheetName, string sourceRange,
            string pivotSheetName, string pivotPosition, string pivotTableName,
            List<string> rowFields = null, List<string> columnFields = null,
            Dictionary<string, string> valueFields = null)
        {
            var workbook = GetWorkbookById(fileName);
            Worksheet sourceSheet = (Worksheet)workbook.Worksheets[sourceSheetName];
            Worksheet pivotSheet = (Worksheet)workbook.Worksheets[pivotSheetName];

            Range sourceRangeObj = sourceSheet.get_Range(sourceRange);
            Range pivotPositionObj = pivotSheet.get_Range(pivotPosition);

            // 创建数据透视缓存
            PivotCache pivotCache = workbook.PivotCaches().Create(
                XlPivotTableSourceType.xlDatabase,
                sourceRangeObj);

            // 创建数据透视表
            PivotTable pivotTable = pivotCache.CreatePivotTable(
                pivotPositionObj,
                pivotTableName);

            // 添加行字段
            if (rowFields != null)
            {
                foreach (var fieldName in rowFields)
                {
                    PivotField rowField = (PivotField)pivotTable.PivotFields(fieldName);
                    rowField.Orientation = XlPivotFieldOrientation.xlRowField;
                    Marshal.ReleaseComObject(rowField);
                }
            }

            // 添加列字段
            if (columnFields != null)
            {
                foreach (var fieldName in columnFields)
                {
                    PivotField colField = (PivotField)pivotTable.PivotFields(fieldName);
                    colField.Orientation = XlPivotFieldOrientation.xlColumnField;
                    Marshal.ReleaseComObject(colField);
                }
            }

            // 添加值字段
            if (valueFields != null)
            {
                foreach (var kvp in valueFields)
                {
                    PivotField valueField = (PivotField)pivotTable.PivotFields(kvp.Key);
                    valueField.Orientation = XlPivotFieldOrientation.xlDataField;

                    // 设置聚合函数
                    valueField.Function = kvp.Value.ToLower() switch
                    {
                        "sum" => XlConsolidationFunction.xlSum,
                        "count" => XlConsolidationFunction.xlCount,
                        "average" => XlConsolidationFunction.xlAverage,
                        "max" => XlConsolidationFunction.xlMax,
                        "min" => XlConsolidationFunction.xlMin,
                        _ => XlConsolidationFunction.xlSum
                    };

                    Marshal.ReleaseComObject(valueField);
                }
            }

            Marshal.ReleaseComObject(pivotTable);
            Marshal.ReleaseComObject(pivotCache);
            Marshal.ReleaseComObject(sourceRangeObj);
            Marshal.ReleaseComObject(pivotPositionObj);
            Marshal.ReleaseComObject(sourceSheet);
            Marshal.ReleaseComObject(pivotSheet);
            Marshal.ReleaseComObject(workbook);
        }

        #endregion

        #region 公式验证 (Formula Validation)

        public string ValidateFormula(string formula)
        {
            try
            {
                // 创建临时工作簿进行公式验证
                Workbook tempWorkbook = _excelApp.Workbooks.Add();
                Worksheet tempSheet = (Worksheet)tempWorkbook.Worksheets[1];
                Range tempCell = (Range)tempSheet.Cells[1, 1];

                try
                {
                    tempCell.Formula = formula;
                    string result = "公式语法正确";
                    Marshal.ReleaseComObject(tempCell);
                    tempWorkbook.Close(false);
                    Marshal.ReleaseComObject(tempSheet);
                    Marshal.ReleaseComObject(tempWorkbook);
                    return result;
                }
                catch (Exception ex)
                {
                    Marshal.ReleaseComObject(tempCell);
                    tempWorkbook.Close(false);
                    Marshal.ReleaseComObject(tempSheet);
                    Marshal.ReleaseComObject(tempWorkbook);
                    return $"公式语法错误: {ex.Message}";
                }
            }
            catch (Exception ex)
            {
                return $"验证失败: {ex.Message}";
            }
        }

        #endregion

        #region 文件管理 (File Management)

        public List<string> GetExcelFiles()
        {
            return Directory.GetFiles(_excelFilesPath, "*.xlsx")
                            .Select(Path.GetFileName)
                            .ToList();
        }

        public void DeleteExcelFile(string fileName)
        {
            if (_openWorkbooks.ContainsKey(fileName))
            {
                throw new InvalidOperationException($"文件 '{fileName}' 正在使用中，请先关闭它。");
            }
            var filePath = Path.Combine(_excelFilesPath, fileName);
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }

        #endregion

        #region 查找和搜索 (Find & Search Operations)

        public List<string> FindValue(string fileName, string sheetName, string searchValue, bool matchCase = false)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            var results = new List<string>();

            Range usedRange = worksheet.UsedRange;
            Range foundCell = usedRange.Find(
                What: searchValue,
                LookIn: XlFindLookIn.xlValues,
                LookAt: XlLookAt.xlPart,
                SearchOrder: XlSearchOrder.xlByRows,
                MatchCase: matchCase);

            if (foundCell != null)
            {
                string firstAddress = foundCell.Address;
                do
                {
                    results.Add(foundCell.Address);
                    foundCell = usedRange.FindNext(foundCell);
                }
                while (foundCell != null && foundCell.Address != firstAddress);
            }

            Marshal.ReleaseComObject(usedRange);
            Marshal.ReleaseComObject(worksheet);

            return results;
        }

        public int FindAndReplace(string fileName, string sheetName, string findWhat, string replaceWith, bool matchCase = false)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range usedRange = worksheet.UsedRange;

            int count = 0;
            Range foundCell = usedRange.Find(
                What: findWhat,
                LookIn: XlFindLookIn.xlValues,
                LookAt: XlLookAt.xlPart,
                MatchCase: matchCase);

            if (foundCell != null)
            {
                string firstAddress = foundCell.Address;
                do
                {
                    foundCell.Value = replaceWith;
                    count++;
                    foundCell = usedRange.FindNext(foundCell);
                }
                while (foundCell != null && foundCell.Address != firstAddress);
            }

            Marshal.ReleaseComObject(usedRange);
            Marshal.ReleaseComObject(worksheet);

            return count;
        }

        #endregion

        #region 视图和布局 (View & Layout Operations)

        public void FreezePanes(string fileName, string sheetName, int row, int column)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            
            // 选择要冻结的位置（冻结点的右下角单元格）
            Range cell = (Range)worksheet.Cells[row, column];
            cell.Select();
            
            // 冻结窗格
            ThisAddIn.app.ActiveWindow.FreezePanes = true;

            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
        }

        public void UnfreezePanes(string fileName, string sheetName)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            ThisAddIn.app.ActiveWindow.FreezePanes = false;
            Marshal.ReleaseComObject(worksheet);
        }

        public void AutoFitColumns(string fileName, string sheetName, string rangeAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            range.Columns.AutoFit();
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public void AutoFitRows(string fileName, string sheetName, string rangeAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            range.Rows.AutoFit();
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public void SetColumnVisible(string fileName, string sheetName, int columnIndex, bool visible)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range column = (Range)worksheet.Columns[columnIndex];
            column.Hidden = !visible;
            Marshal.ReleaseComObject(column);
            Marshal.ReleaseComObject(worksheet);
        }

        public void SetRowVisible(string fileName, string sheetName, int rowIndex, bool visible)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range row = (Range)worksheet.Rows[rowIndex];
            row.Hidden = !visible;
            Marshal.ReleaseComObject(row);
            Marshal.ReleaseComObject(worksheet);
        }

        #endregion

        #region 批注操作 (Comment Operations)

        public void AddComment(string fileName, string sheetName, string cellAddress, string commentText)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = worksheet.get_Range(cellAddress);
            
            // 如果已有批注，先删除
            if (cell.Comment != null)
            {
                cell.Comment.Delete();
            }
            
            cell.AddComment(commentText);
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
        }

        public void DeleteComment(string fileName, string sheetName, string cellAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = worksheet.get_Range(cellAddress);
            
            if (cell.Comment != null)
            {
                cell.Comment.Delete();
            }
            
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
        }

        public string GetComment(string fileName, string sheetName, string cellAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = worksheet.get_Range(cellAddress);
            
            string commentText = cell.Comment?.Text() ?? "";
            
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
            
            return commentText;
        }

        #endregion

        #region 超链接操作 (Hyperlink Operations)

        public void AddHyperlink(string fileName, string sheetName, string cellAddress, string url, string displayText = null)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = worksheet.get_Range(cellAddress);
            
            // 判断是否为文档内跳转
            if (url.StartsWith("#"))
            {
                // 文档内跳转，使用SubAddress参数
                worksheet.Hyperlinks.Add(
                    Anchor: cell,
                    Address: "",
                    SubAddress: url.TrimStart('#'),
                    TextToDisplay: displayText ?? url);
            }
            else
            {
                // 外部链接（网址、文件等），使用Address参数
                worksheet.Hyperlinks.Add(
                    Anchor: cell,
                    Address: url,
                    TextToDisplay: displayText ?? url);
            }
            
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
        }

        /// <summary>
        /// 使用HYPERLINK公式设置内部跳转链接
        /// </summary>
        /// <param name="fileName">工作簿文件名</param>
        /// <param name="sheetName">工作表名</param>
        /// <param name="cellAddress">单元格地址</param>
        /// <param name="targetLocation">目标位置（如"Sheet2!A1"）</param>
        /// <param name="displayText">显示文本</param>
        public void SetHyperlinkFormula(string fileName, string sheetName, string cellAddress, string targetLocation, string displayText)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = worksheet.get_Range(cellAddress);
            
            // 使用HYPERLINK公式创建内部跳转
            // 格式: =HYPERLINK("#Sheet2!A1", "显示文本")
            string formula = $"=HYPERLINK(\"#{targetLocation}\", \"{displayText}\")";
            cell.Formula = formula;
            
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
        }

        public void DeleteHyperlink(string fileName, string sheetName, string cellAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range cell = worksheet.get_Range(cellAddress);
            
            if (cell.Hyperlinks.Count > 0)
            {
                cell.Hyperlinks.Delete();
            }
            
            Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(worksheet);
        }

        #endregion

        #region 数据分析 (Data Analysis Operations)

        public string GetUsedRange(string fileName, string sheetName)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range usedRange = worksheet.UsedRange;
            string address = usedRange.Address;
            Marshal.ReleaseComObject(usedRange);
            Marshal.ReleaseComObject(worksheet);
            return address;
        }

        public Dictionary<string, object> GetRangeStatistics(string fileName, string sheetName, string rangeAddress)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            
            var stats = new Dictionary<string, object>();
            
            try
            {
                stats["Sum"] = ThisAddIn.app.WorksheetFunction.Sum(range);
            }
            catch { stats["Sum"] = "N/A"; }
            
            try
            {
                stats["Average"] = ThisAddIn.app.WorksheetFunction.Average(range);
            }
            catch { stats["Average"] = "N/A"; }
            
            try
            {
                stats["Count"] = ThisAddIn.app.WorksheetFunction.Count(range);
            }
            catch { stats["Count"] = "N/A"; }
            
            try
            {
                stats["Max"] = ThisAddIn.app.WorksheetFunction.Max(range);
            }
            catch { stats["Max"] = "N/A"; }
            
            try
            {
                stats["Min"] = ThisAddIn.app.WorksheetFunction.Min(range);
            }
            catch { stats["Min"] = "N/A"; }
            
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
            
            return stats;
        }

        public int GetLastRow(string fileName, string sheetName, int columnIndex = 1)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range column = (Range)worksheet.Columns[columnIndex];
            Range lastCell = column.Find("*", SearchOrder: XlSearchOrder.xlByRows, SearchDirection: XlSearchDirection.xlPrevious);
            
            int lastRow = lastCell?.Row ?? 0;
            
            if (lastCell != null) Marshal.ReleaseComObject(lastCell);
            Marshal.ReleaseComObject(column);
            Marshal.ReleaseComObject(worksheet);
            
            return lastRow;
        }

        public int GetLastColumn(string fileName, string sheetName, int rowIndex = 1)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range row = (Range)worksheet.Rows[rowIndex];
            Range lastCell = row.Find("*", SearchOrder: XlSearchOrder.xlByColumns, SearchDirection: XlSearchDirection.xlPrevious);
            
            int lastColumn = lastCell?.Column ?? 0;
            
            if (lastCell != null) Marshal.ReleaseComObject(lastCell);
            Marshal.ReleaseComObject(row);
            Marshal.ReleaseComObject(worksheet);
            
            return lastColumn;
        }

        #endregion

        #region 排序和筛选 (Sort & Filter Operations)

        public void SortRange(string fileName, string sheetName, string rangeAddress, int sortColumnIndex, bool ascending = true)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            Range sortKey = (Range)range.Columns[sortColumnIndex];
            
            range.Sort(
                Key1: sortKey,
                Order1: ascending ? XlSortOrder.xlAscending : XlSortOrder.xlDescending,
                Header: XlYesNoGuess.xlYes);
            
            Marshal.ReleaseComObject(sortKey);
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public void SetAutoFilter(string fileName, string sheetName, string rangeAddress, int columnIndex = 0, string criteria = null)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            
            // 如果已有筛选，先清除
            if (worksheet.AutoFilterMode)
            {
                worksheet.AutoFilterMode = false;
            }
            
            if (columnIndex > 0 && !string.IsNullOrEmpty(criteria))
            {
                range.AutoFilter(Field: columnIndex, Criteria1: criteria);
            }
            else
            {
                range.AutoFilter();
            }
            
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        public void RemoveDuplicates(string fileName, string sheetName, string rangeAddress, int[] columnIndices)
        {
            Worksheet worksheet = GetWorksheetById(fileName, sheetName);
            Range range = worksheet.get_Range(rangeAddress);
            
            // 转换为object数组（Excel COM需要）
            object[] columns = columnIndices.Cast<object>().ToArray();
            
            range.RemoveDuplicates(Columns: columns, Header: XlYesNoGuess.xlYes);
            
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        #endregion

        #region 工作表高级操作 (Advanced Worksheet Operations)

        public void MoveWorksheet(string fileName, string sheetName, int position)
        {
            var workbook = GetWorkbookById(fileName);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[sheetName];
            
            if (position == 1)
            {
                worksheet.Move(Before: workbook.Worksheets[1]);
            }
            else if (position > workbook.Worksheets.Count)
            {
                worksheet.Move(After: workbook.Worksheets[workbook.Worksheets.Count]);
            }
            else
            {
                worksheet.Move(Before: workbook.Worksheets[position]);
            }
            
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
        }

        public void SetWorksheetVisible(string fileName, string sheetName, bool visible)
        {
            var workbook = GetWorkbookById(fileName);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[sheetName];
            worksheet.Visible = visible ? XlSheetVisibility.xlSheetVisible : XlSheetVisibility.xlSheetHidden;
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
        }

        public int GetWorksheetIndex(string fileName, string sheetName)
        {
            var workbook = GetWorkbookById(fileName);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[sheetName];
            int index = worksheet.Index;
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            return index;
        }

        #endregion

        #region 命名区域操作 (Named Ranges Operations)

        /// <summary>
        /// 创建命名区域
        /// </summary>
        public void CreateNamedRange(string fileName, string sheetName, string rangeName, string rangeAddress)
        {
            var worksheet = GetWorksheetById(fileName, sheetName);
            var range = worksheet.get_Range(rangeAddress);

            var workbook = GetWorkbookById(fileName);
            workbook.Names.Add(rangeName, range);

            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        /// <summary>
        /// 删除命名区域
        /// </summary>
        public void DeleteNamedRange(string fileName, string rangeName)
        {
            var workbook = GetWorkbookById(fileName);

            foreach (Name name in workbook.Names)
            {
                if (name.Name == rangeName)
                {
                    name.Delete();
                    Marshal.ReleaseComObject(name);
                    break;
                }
                Marshal.ReleaseComObject(name);
            }
        }

        /// <summary>
        /// 获取所有命名区域
        /// </summary>
        public List<string> GetNamedRanges(string fileName)
        {
            var workbook = GetWorkbookById(fileName);
            var names = new List<string>();

            foreach (Name name in workbook.Names)
            {
                names.Add($"{name.Name} = {name.RefersTo}");
                Marshal.ReleaseComObject(name);
            }

            return names;
        }

        /// <summary>
        /// 获取命名区域的引用地址
        /// </summary>
        public string GetNamedRangeAddress(string fileName, string rangeName)
        {
            var workbook = GetWorkbookById(fileName);

            foreach (Name name in workbook.Names)
            {
                if (name.Name == rangeName)
                {
                    string refersTo = name.RefersTo.ToString();
                    Marshal.ReleaseComObject(name);
                    return refersTo;
                }
                Marshal.ReleaseComObject(name);
            }

            throw new ArgumentException($"命名区域 '{rangeName}' 不存在。");
        }

        #endregion

        #region 单元格格式增强 (Cell Format Enhancement)

        /// <summary>
        /// 设置单元格文本自动换行
        /// </summary>
        public void SetCellTextWrap(string fileName, string sheetName, string rangeAddress, bool wrap)
        {
            var worksheet = GetWorksheetById(fileName, sheetName);
            var range = worksheet.get_Range(rangeAddress);
            range.WrapText = wrap;
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        /// <summary>
        /// 设置单元格缩进级别
        /// </summary>
        public void SetCellIndent(string fileName, string sheetName, string rangeAddress, int indentLevel)
        {
            var worksheet = GetWorksheetById(fileName, sheetName);
            var range = worksheet.get_Range(rangeAddress);
            range.IndentLevel = indentLevel;
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        /// <summary>
        /// 设置单元格文本旋转角度
        /// </summary>
        public void SetCellOrientation(string fileName, string sheetName, string rangeAddress, int degrees)
        {
            var worksheet = GetWorksheetById(fileName, sheetName);
            var range = worksheet.get_Range(rangeAddress);
            range.Orientation = degrees; // -90 to 90
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        /// <summary>
        /// 设置单元格缩小字体填充
        /// </summary>
        public void SetCellShrinkToFit(string fileName, string sheetName, string rangeAddress, bool shrink)
        {
            var worksheet = GetWorksheetById(fileName, sheetName);
            var range = worksheet.get_Range(rangeAddress);
            range.ShrinkToFit = shrink;
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(worksheet);
        }

        #endregion

        #region IDisposable Implementation

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    foreach (var fileName in _openWorkbooks.Keys.ToList())
                    {
                        CloseWorkbook(fileName);
                    }
                }

                if (_excelApp != null)
                {
                    _excelApp.Quit();
                    Marshal.ReleaseComObject(_excelApp);
                }

                _disposed = true;
            }
        }

        ~ExcelMcp()
        {
            Dispose(false);
        }

        #endregion
    }
}
