using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

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
