using ClosedXML.Excel;
using System.Drawing;
using System.Text;

namespace TableMagic.Cli.Excel;

public class ClosedXmlExcelProvider : IExcelProvider
{
    private readonly string _basePath;
    private readonly Dictionary<string, XLWorkbook> _openWorkbooks = new();
    private readonly HashSet<string> _dirtyWorkbooks = new();
    private readonly Dictionary<string, HashSet<string>> _worksheetNameCache = new();

    public ClosedXmlExcelProvider(string basePath = "./excel_files")
    {
        _basePath = Path.GetFullPath(basePath);
        if (!Directory.Exists(_basePath))
            Directory.CreateDirectory(_basePath);
    }

    private XLWorkbook GetWorkbook(string fileName)
    {
        if (string.IsNullOrEmpty(fileName))
        {
            if (_openWorkbooks.Count > 0)
                return _openWorkbooks.Values.First();
            throw new ArgumentException("没有打开的工作簿");
        }

        if (_openWorkbooks.TryGetValue(fileName, out var wb))
            return wb;

        var filePath = Path.Combine(_basePath, fileName);
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"文件不存在: {filePath}");

        var workbook = new XLWorkbook(filePath);
        _openWorkbooks[fileName] = workbook;
        return workbook;
    }

    private IXLWorksheet GetWorksheet(XLWorkbook workbook, string sheetName, string? fileName = null)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            if (workbook.Worksheets.Count > 0)
                return workbook.Worksheets.First();
            throw new ArgumentException("没有可用的工作表");
        }
        if (fileName != null && _worksheetNameCache.TryGetValue(fileName, out var cache) && cache.Contains(sheetName))
            return workbook.Worksheets.Worksheet(sheetName);
        return workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName)
            ?? throw new ArgumentException($"工作表 '{sheetName}' 不存在");
    }

    private void SaveWorkbookIfNeeded(XLWorkbook workbook, string fileName)
    {
        _dirtyWorkbooks.Add(fileName);
    }

    private void FlushDirtyWorkbooks()
    {
        foreach (var fileName in _dirtyWorkbooks.ToList())
        {
            if (_openWorkbooks.TryGetValue(fileName, out var wb))
            {
                var filePath = Path.Combine(_basePath, fileName);
                wb.SaveAs(filePath);
            }
        }
        _dirtyWorkbooks.Clear();
    }

    public string CreateWorkbook(string fileName, string sheetName = "Sheet1")
    {
        var filePath = Path.Combine(_basePath, fileName);
        if (File.Exists(filePath))
            throw new ArgumentException($"文件已存在: {filePath}");

        var workbook = new XLWorkbook();
        var ws = workbook.Worksheets.Add(sheetName);
        workbook.SaveAs(filePath);
        _openWorkbooks[fileName] = workbook;
        _worksheetNameCache[fileName] = new HashSet<string> { sheetName };
        return filePath;
    }

    public string OpenWorkbook(string fileName)
    {
        if (_openWorkbooks.ContainsKey(fileName))
            return fileName;

        var filePath = Path.Combine(_basePath, fileName);
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"文件不存在: {filePath}");

        var workbook = new XLWorkbook(filePath);
        _openWorkbooks[fileName] = workbook;
        _worksheetNameCache[fileName] = new HashSet<string>(workbook.Worksheets.Select(ws => ws.Name));
        return fileName;
    }

    public void CloseWorkbook(string fileName)
    {
        if (_openWorkbooks.TryGetValue(fileName, out var workbook))
        {
            var filePath = Path.Combine(_basePath, fileName);
            workbook.SaveAs(filePath);
            workbook.Dispose();
            _openWorkbooks.Remove(fileName);
            _dirtyWorkbooks.Remove(fileName);
            _worksheetNameCache.Remove(fileName);
        }
    }

    public void SaveWorkbook(string fileName)
    {
        var workbook = GetWorkbook(fileName);
        var filePath = Path.Combine(_basePath, fileName);
        workbook.SaveAs(filePath);
        _dirtyWorkbooks.Remove(fileName);
    }

    public void SaveWorkbookAs(string fileName, string newFileName)
    {
        var workbook = GetWorkbook(fileName);
        var newFilePath = Path.Combine(_basePath, newFileName);
        workbook.SaveAs(newFilePath);

        if (_openWorkbooks.Remove(fileName))
        {
            _openWorkbooks[newFileName] = workbook;
            if (_worksheetNameCache.Remove(fileName, out var cache))
                _worksheetNameCache[newFileName] = cache;
        }
    }

    public void DeleteWorkbook(string fileName)
    {
        if (_openWorkbooks.TryGetValue(fileName, out var workbook))
        {
            workbook.Dispose();
            _openWorkbooks.Remove(fileName);
            _worksheetNameCache.Remove(fileName);
        }
        var filePath = Path.Combine(_basePath, fileName);
        if (File.Exists(filePath))
            File.Delete(filePath);
    }

    public List<string> GetOpenWorkbooks() => _openWorkbooks.Keys.ToList();

    public string GetWorkbookMetadata(string fileName)
    {
        var workbook = GetWorkbook(fileName);
        var sb = new StringBuilder();
        sb.AppendLine($"工作簿名称: {fileName}");
        sb.AppendLine($"工作表数量: {workbook.Worksheets.Count}");
        sb.AppendLine("工作表列表:");
        foreach (var ws in workbook.Worksheets)
        {
            sb.AppendLine($"  - {ws.Name} (已使用范围: {ws.RangeUsed()?.RangeAddress?.ToString() ?? "空"})");
        }
        return sb.ToString();
    }

    public string CreateWorksheet(string fileName, string sheetName)
    {
        var workbook = GetWorkbook(fileName);
        if (workbook.Worksheets.Any(ws => ws.Name == sheetName))
            throw new ArgumentException($"工作表 '{sheetName}' 已存在");
        var ws = workbook.Worksheets.Add(sheetName);
        if (_worksheetNameCache.TryGetValue(fileName, out var cache))
            cache.Add(sheetName);
        SaveWorkbookIfNeeded(workbook, fileName);
        return sheetName;
    }

    public void RenameWorksheet(string fileName, string oldSheetName, string newSheetName)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, oldSheetName, fileName);
        ws.Name = newSheetName;
        if (_worksheetNameCache.TryGetValue(fileName, out var cache))
        {
            cache.Remove(oldSheetName);
            cache.Add(newSheetName);
        }
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void DeleteWorksheet(string fileName, string sheetName)
    {
        var workbook = GetWorkbook(fileName);
        workbook.Worksheets.Delete(sheetName);
        if (_worksheetNameCache.TryGetValue(fileName, out var cache))
            cache.Remove(sheetName);
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public List<string> GetWorksheetNames(string fileName)
    {
        if (_worksheetNameCache.TryGetValue(fileName, out var cache))
            return cache.ToList();
        var workbook = GetWorkbook(fileName);
        var names = workbook.Worksheets.Select(ws => ws.Name).ToList();
        _worksheetNameCache[fileName] = new HashSet<string>(names);
        return names;
    }

    public string ActivateWorksheet(string fileName, string sheetName)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.SetTabActive();
        SaveWorkbookIfNeeded(workbook, fileName);
        return sheetName;
    }

    public string GetActiveWorksheetName(string fileName)
    {
        var workbook = GetWorkbook(fileName);
        var active = workbook.Worksheets.FirstOrDefault(ws => ws.TabActive);
        return active?.Name ?? workbook.Worksheets.First().Name;
    }

    public void CopyWorksheet(string fileName, string sourceSheetName, string targetSheetName)
    {
        var workbook = GetWorkbook(fileName);
        var source = GetWorksheet(workbook, sourceSheetName, fileName);
        var newWs = source.CopyTo(targetSheetName);
        if (_worksheetNameCache.TryGetValue(fileName, out var cache))
            cache.Add(targetSheetName);
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void MoveWorksheet(string fileName, string sheetName, int position)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Position = position;
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void SetWorksheetVisible(string fileName, string sheetName, bool visible)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Visibility = visible ? XLWorksheetVisibility.Visible : XLWorksheetVisibility.Hidden;
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public int GetWorksheetIndex(string fileName, string sheetName)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        return ws.Position;
    }

    public void FreezePanes(string fileName, string sheetName, int row, int column)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.SheetView.FreezeRows(row);
        if (column > 1) ws.SheetView.FreezeColumns(column);
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void UnfreezePanes(string fileName, string sheetName)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.SheetView.FreezeRows(0);
        ws.SheetView.FreezeColumns(0);
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void SetCellValue(string fileName, string sheetName, int row, int column, object value)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Cell(row, column).Value = ConvertToXLCellValue(value);
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public object GetCellValue(string fileName, string sheetName, int row, int column)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        var cell = ws.Cell(row, column);
        return cell.Value.ToString();
    }

    public void SetRangeValues(string fileName, string sheetName, string rangeAddress, object[,] data)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        var range = ws.Range(rangeAddress);

        int rows = data.GetLength(0);
        int cols = data.GetLength(1);
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                var val = data[r, c];
                if (val != null)
                    range.Cell(r + 1, c + 1).Value = ConvertToXLCellValue(val);
            }
        }
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public object[,] GetRangeValues(string fileName, string sheetName, string rangeAddress)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        var range = ws.Range(rangeAddress);

        int rows = range.RowCount();
        int cols = range.ColumnCount();
        var result = new object[rows, cols];

        for (int r = 1; r <= rows; r++)
        {
            for (int c = 1; c <= cols; c++)
            {
                result[r - 1, c - 1] = range.Cell(r, c).Value.ToString();
            }
        }
        return result;
    }

    public void SetFormula(string fileName, string sheetName, string cellAddress, string formula)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Cell(cellAddress).FormulaA1 = formula;
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public string GetFormula(string fileName, string sheetName, string cellAddress)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        return ws.Cell(cellAddress).FormulaA1;
    }

    public void SetCellFormat(string fileName, string sheetName, string rangeAddress,
        string? fontColor = null, string? backgroundColor = null, int? fontSize = null,
        bool? bold = null, bool? italic = null,
        string? horizontalAlignment = null, string? verticalAlignment = null)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        var range = ws.Range(rangeAddress);

        if (fontColor != null) range.Style.Font.FontColor = ParseXLColor(fontColor);
        if (backgroundColor != null) range.Style.Fill.BackgroundColor = ParseXLColor(backgroundColor);
        if (fontSize.HasValue) range.Style.Font.FontSize = fontSize.Value;
        if (bold.HasValue) range.Style.Font.Bold = bold.Value;
        if (italic.HasValue) range.Style.Font.Italic = italic.Value;

        if (horizontalAlignment != null)
        {
            range.Style.Alignment.Horizontal = horizontalAlignment.ToLower() switch
            {
                "left" => XLAlignmentHorizontalValues.Left,
                "center" => XLAlignmentHorizontalValues.Center,
                "right" => XLAlignmentHorizontalValues.Right,
                _ => XLAlignmentHorizontalValues.General
            };
        }

        if (verticalAlignment != null)
        {
            range.Style.Alignment.Vertical = verticalAlignment.ToLower() switch
            {
                "top" => XLAlignmentVerticalValues.Top,
                "center" => XLAlignmentVerticalValues.Center,
                "bottom" => XLAlignmentVerticalValues.Bottom,
                _ => XLAlignmentVerticalValues.Center
            };
        }

        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void SetBorder(string fileName, string sheetName, string rangeAddress, string borderType, string lineStyle = "continuous")
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        var range = ws.Range(rangeAddress);

        var xlLineStyle = lineStyle.ToLower() switch
        {
            "dash" => XLBorderStyleValues.Dashed,
            "dot" => XLBorderStyleValues.Dotted,
            _ => XLBorderStyleValues.Thin
        };

        switch (borderType.ToLower())
        {
            case "all":
                range.Style.Border.SetOutsideBorder(xlLineStyle);
                range.Style.Border.SetInsideBorder(xlLineStyle);
                break;
            case "outline":
                range.Style.Border.SetOutsideBorder(xlLineStyle);
                break;
            case "horizontal":
                range.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
                break;
            case "vertical":
                range.Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
                break;
        }

        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void MergeCells(string fileName, string sheetName, string rangeAddress)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Range(rangeAddress).Merge();
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void UnmergeCells(string fileName, string sheetName, string rangeAddress)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Range(rangeAddress).Unmerge();
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void SetCellTextWrap(string fileName, string sheetName, string rangeAddress, bool wrap)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Range(rangeAddress).Style.Alignment.WrapText = wrap;
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void SetRowHeight(string fileName, string sheetName, int rowNumber, double height)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Row(rowNumber).Height = height;
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void SetColumnWidth(string fileName, string sheetName, int columnNumber, double width)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Column(columnNumber).Width = width;
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void InsertRows(string fileName, string sheetName, int rowIndex, int count = 1)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Row(rowIndex).InsertRowsAbove(count);
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void InsertColumns(string fileName, string sheetName, int columnIndex, int count = 1)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Column(columnIndex).InsertColumnsBefore(count);
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void DeleteRows(string fileName, string sheetName, int rowIndex, int count = 1)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        for (int i = 0; i < count; i++)
            ws.Row(rowIndex).Delete();
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void DeleteColumns(string fileName, string sheetName, int columnIndex, int count = 1)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        for (int i = 0; i < count; i++)
            ws.Column(columnIndex).Delete();
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void CopyRange(string fileName, string sheetName, string sourceRange, string targetRange)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        var source = ws.Range(sourceRange);
        var target = ws.Range(targetRange);
        source.CopyTo(target);
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void ClearRange(string fileName, string sheetName, string rangeAddress, string clearType = "all")
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        var range = ws.Range(rangeAddress);

        switch (clearType.ToLower())
        {
            case "contents":
                range.Clear();
                break;
            case "formats":
                range.Clear(XLClearOptions.NormalFormats);
                break;
            case "all":
            default:
                range.Clear(XLClearOptions.All);
                break;
        }

        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public string CreateChart(string fileName, string sheetName, string dataRange, string chartType = "column", string title = "")
    {
        throw new NotSupportedException("ClosedXML不支持图表创建，请使用Excel COM模式");
    }

    public string CreatePivotTable(string fileName, string sheetName, string sourceRange, string pivotSheetName,
        string? rowFields = null, string? columnFields = null, string? valueFields = null)
    {
        return $"数据透视表创建功能在ClosedXML模式下有限支持。请使用Excel COM模式获取完整功能。";
    }

    public string GetRangeStatistics(string fileName, string sheetName, string range)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        var rng = ws.Range(range);

        var sb = new StringBuilder();
        sb.AppendLine($"范围: {range}");

        var numericCells = rng.CellsUsed()
            .Where(c => c.DataType == XLDataType.Number)
            .Select(c => c.GetDouble())
            .ToList();

        if (numericCells.Count > 0)
        {
            sb.AppendLine($"数值单元格数: {numericCells.Count}");
            sb.AppendLine($"最小值: {numericCells.Min():F2}");
            sb.AppendLine($"最大值: {numericCells.Max():F2}");
            sb.AppendLine($"平均值: {numericCells.Average():F2}");
            sb.AppendLine($"总和: {numericCells.Sum():F2}");
        }

        var textCells = rng.CellsUsed()
            .Where(c => c.DataType == XLDataType.Text)
            .ToList();

        sb.AppendLine($"文本单元格数: {textCells.Count}");
        sb.AppendLine($"已使用单元格总数: {rng.CellsUsed().Count()}");

        return sb.ToString();
    }

    public string AnalyzeData(string fileName, string sheetName, string range)
    {
        return GetRangeStatistics(fileName, sheetName, range);
    }

    public void SortRange(string fileName, string sheetName, string rangeAddress, string sortColumn, bool ascending = true)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        var range = ws.Range(rangeAddress);

        var sortCol = int.TryParse(sortColumn, out var colIdx) ? colIdx : 1;
        range.SortColumns.Add(sortCol, ascending ? XLSortOrder.Ascending : XLSortOrder.Descending);
        range.Sort();

        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void SetAutoFilter(string fileName, string sheetName, string rangeAddress)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        ws.Range(rangeAddress).SetAutoFilter();
        SaveWorkbookIfNeeded(workbook, fileName);
    }

    public void RemoveDuplicates(string fileName, string sheetName, string rangeAddress, int[] columns)
    {
        throw new NotSupportedException("ClosedXML不支持RemoveDuplicates，请使用Excel COM模式");
    }


    public void ExportToPdf(string fileName, string sheetName, string pdfPath)
    {
        var wb = GetWorkbook(fileName);
        var ws = GetWorksheet(wb, sheetName);
        var data = ReadSheetDataFromWorksheet(ws);
        PdfExporter.ExportSheet(data, sheetName, pdfPath);
    }

    public void ExportWorkbookToPdf(string fileName, string pdfPath)
    {
        var wb = GetWorkbook(fileName);
        var sheets = new List<(string Name, List<string[]> Data)>();
        foreach (var ws in wb.Worksheets)
        {
            sheets.Add((ws.Name, ReadSheetDataFromWorksheet(ws)));
        }
        PdfExporter.ExportWorkbook(sheets, pdfPath);
    }

    private static List<string[]> ReadSheetDataFromWorksheet(IXLWorksheet ws)
    {
        var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
        var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
        var data = new List<string[]>();
        for (int r = 1; r <= lastRow; r++)
        {
            var row = new string[lastCol];
            for (int c = 1; c <= lastCol; c++)
            {
                row[c - 1] = ws.Cell(r, c).GetString();
            }
            data.Add(row);
        }
        return data;
    }

    public string GetUsedRange(string fileName, string sheetName)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        return ws.RangeUsed()?.RangeAddress?.ToString() ?? "A1";
    }

    public int GetLastRow(string fileName, string sheetName)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        return ws.LastRowUsed()?.RowNumber() ?? 0;
    }

    public int GetLastColumn(string fileName, string sheetName)
    {
        var workbook = GetWorkbook(fileName);
        var ws = GetWorksheet(workbook, sheetName);
        return ws.LastColumnUsed()?.ColumnNumber() ?? 0;
    }

    private static XLCellValue ConvertToXLCellValue(object value)
    {
        return value switch
        {
            int i => i,
            long l => l,
            double d => d,
            float f => f,
            decimal dec => dec,
            bool b => b,
            DateTime dt => dt,
            string s when double.TryParse(s, out var num) => num,
            string s => s,
            _ => value?.ToString() ?? ""
        };
    }

    private static XLColor ParseXLColor(string colorStr)
    {
        if (colorStr.StartsWith("#"))
        {
            var hex = colorStr.Substring(1);
            var r = Convert.ToInt32(hex.Substring(0, 2), 16);
            var g = Convert.ToInt32(hex.Substring(2, 2), 16);
            var b = Convert.ToInt32(hex.Substring(4, 2), 16);
            return XLColor.FromColor(Color.FromArgb(r, g, b));
        }

        return colorStr.ToLower() switch
        {
            "红色" or "red" => XLColor.Red,
            "绿色" or "green" => XLColor.Green,
            "蓝色" or "blue" => XLColor.Blue,
            "黄色" or "yellow" => XLColor.Yellow,
            "橙色" or "orange" => XLColor.Orange,
            "紫色" or "purple" => XLColor.Purple,
            "黑色" or "black" => XLColor.Black,
            "白色" or "white" => XLColor.White,
            "灰色" or "gray" => XLColor.Gray,
            _ => XLColor.Black
        };
    }

    public void Dispose()
    {
        FlushDirtyWorkbooks();
        foreach (var wb in _openWorkbooks.Values)
        {
            try { wb.Dispose(); } catch { }
        }
        _openWorkbooks.Clear();
        _dirtyWorkbooks.Clear();
        _worksheetNameCache.Clear();
    }
}