namespace TableMagic.Cli.Excel;

public interface IExcelProvider : IDisposable
{
    string CreateWorkbook(string fileName, string sheetName = "Sheet1");
    string OpenWorkbook(string fileName);
    void CloseWorkbook(string fileName);
    void SaveWorkbook(string fileName);
    void SaveWorkbookAs(string fileName, string newFileName);
    void DeleteWorkbook(string fileName);
    List<string> GetOpenWorkbooks();
    string GetWorkbookMetadata(string fileName);

    string CreateWorksheet(string fileName, string sheetName);
    void RenameWorksheet(string fileName, string oldSheetName, string newSheetName);
    void DeleteWorksheet(string fileName, string sheetName);
    List<string> GetWorksheetNames(string fileName);
    string ActivateWorksheet(string fileName, string sheetName);
    string GetActiveWorksheetName(string fileName);
    void CopyWorksheet(string fileName, string sourceSheetName, string targetSheetName);
    void MoveWorksheet(string fileName, string sheetName, int position);
    void SetWorksheetVisible(string fileName, string sheetName, bool visible);
    int GetWorksheetIndex(string fileName, string sheetName);
    void FreezePanes(string fileName, string sheetName, int row, int column);
    void UnfreezePanes(string fileName, string sheetName);

    void SetCellValue(string fileName, string sheetName, int row, int column, object value);
    object GetCellValue(string fileName, string sheetName, int row, int column);
    void SetRangeValues(string fileName, string sheetName, string rangeAddress, object[,] data);
    object[,] GetRangeValues(string fileName, string sheetName, string rangeAddress);
    void SetFormula(string fileName, string sheetName, string cellAddress, string formula);
    string GetFormula(string fileName, string sheetName, string cellAddress);

    void SetCellFormat(string fileName, string sheetName, string rangeAddress,
        string? fontColor = null, string? backgroundColor = null, int? fontSize = null,
        bool? bold = null, bool? italic = null,
        string? horizontalAlignment = null, string? verticalAlignment = null);
    void SetBorder(string fileName, string sheetName, string rangeAddress, string borderType, string lineStyle = "continuous");
    void MergeCells(string fileName, string sheetName, string rangeAddress);
    void UnmergeCells(string fileName, string sheetName, string rangeAddress);
    void SetCellTextWrap(string fileName, string sheetName, string rangeAddress, bool wrap);
    void SetRowHeight(string fileName, string sheetName, int rowNumber, double height);
    void SetColumnWidth(string fileName, string sheetName, int columnNumber, double width);
    void InsertRows(string fileName, string sheetName, int rowIndex, int count = 1);
    void InsertColumns(string fileName, string sheetName, int columnIndex, int count = 1);
    void DeleteRows(string fileName, string sheetName, int rowIndex, int count = 1);
    void DeleteColumns(string fileName, string sheetName, int columnIndex, int count = 1);

    void CopyRange(string fileName, string sheetName, string sourceRange, string targetRange);
    void ClearRange(string fileName, string sheetName, string rangeAddress, string clearType = "all");

    string CreateChart(string fileName, string sheetName, string dataRange, string chartType = "column", string title = "");
    string CreatePivotTable(string fileName, string sheetName, string sourceRange, string pivotSheetName,
        string? rowFields = null, string? columnFields = null, string? valueFields = null);

    string GetRangeStatistics(string fileName, string sheetName, string range);
    string AnalyzeData(string fileName, string sheetName, string range);

    void SortRange(string fileName, string sheetName, string rangeAddress, string sortColumn, bool ascending = true);
    void SetAutoFilter(string fileName, string sheetName, string rangeAddress);
    void RemoveDuplicates(string fileName, string sheetName, string rangeAddress, int[] columns);

    void ExportToPdf(string fileName, string sheetName, string pdfPath);
    void ExportWorkbookToPdf(string fileName, string pdfPath);

    string GetUsedRange(string fileName, string sheetName);
    int GetLastRow(string fileName, string sheetName);
    int GetLastColumn(string fileName, string sheetName);
}