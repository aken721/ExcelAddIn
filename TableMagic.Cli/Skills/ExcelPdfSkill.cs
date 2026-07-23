using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelPdfSkill : ISkill
{
    private readonly IExcelProvider _provider;
    public ExcelPdfSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelPdf";
    public string Description => "PDF导出：单表/全簿/批量/区域导出";

    static ExcelPdfSkill()
    {
        QuestPDF.Settings.License = LicenseType.Community;
    }

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "export_sheet_to_pdf", Description = "将指定工作表导出为PDF",
                Parameters = P(new[]{"pdfPath"}, new[]{"fileName","sheetName"}), RequiredParameters = new List<string>{"pdfPath"} },
            new() { Name = "export_workbook_to_pdf", Description = "将整个工作簿导出为PDF",
                Parameters = P(new[]{"pdfPath"}, new[]{"fileName"}), RequiredParameters = new List<string>{"pdfPath"} },
            new() { Name = "batch_export_to_pdf", Description = "批量导出多个工作表为PDF",
                Parameters = P(new[]{"outputFolder"}, new[]{"fileName","sheetNames"}), RequiredParameters = new List<string>{"outputFolder"} },
            new() { Name = "export_range_to_pdf", Description = "将指定区域导出为PDF",
                Parameters = P(new[]{"pdfPath","rangeAddress"}, new[]{"fileName","sheetName"}), RequiredParameters = new List<string>{"pdfPath","rangeAddress"} }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "export_sheet_to_pdf" => ExportSheetToPdf(arguments),
                "export_workbook_to_pdf" => ExportWorkbookToPdf(arguments),
                "batch_export_to_pdf" => BatchExportToPdf(arguments),
                "export_range_to_pdf" => ExportRangeToPdf(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult ExportSheetToPdf(Dictionary<string, object> args)
    {
        var pdfPath = GetStr(args, "pdfPath");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");

        if (string.IsNullOrEmpty(pdfPath))
            return SkillResult.MissingParamsResult("export_sheet_to_pdf", new List<MissingParam>
            {
                new() { Name = "pdfPath", Description = "PDF输出文件路径", PromptHint = "请提供PDF输出文件路径，如 C:\\output\\sheet.pdf" }
            });

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;

        var data = ReadSheetData(wb, sn);
        GeneratePdf(data, pdfPath, sn);

        return SkillResult.Ok($"工作表 '{sn}' 已导出为PDF:\n{pdfPath}");
    }

    private SkillResult ExportWorkbookToPdf(Dictionary<string, object> args)
    {
        var pdfPath = GetStr(args, "pdfPath");
        var fileName = GetStr(args, "fileName");

        if (string.IsNullOrEmpty(pdfPath))
            return SkillResult.MissingParamsResult("export_workbook_to_pdf", new List<MissingParam>
            {
                new() { Name = "pdfPath", Description = "PDF输出文件路径", PromptHint = "请提供PDF输出文件路径" }
            });

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");

        var sheets = _provider.GetWorksheetNames(wb);
        var allData = new List<(string SheetName, List<string[]> Data)>();
        foreach (var sn in sheets)
        {
            allData.Add((sn, ReadSheetData(wb, sn)));
        }

        GenerateMultiSheetPdf(allData, pdfPath);
        return SkillResult.Ok($"工作簿已导出为PDF:\n{pdfPath}");
    }

    private SkillResult BatchExportToPdf(Dictionary<string, object> args)
    {
        var outputFolder = GetStr(args, "outputFolder");
        var fileName = GetStr(args, "fileName");
        var sheetNamesStr = GetStr(args, "sheetNames");

        if (string.IsNullOrEmpty(outputFolder))
            return SkillResult.MissingParamsResult("batch_export_to_pdf", new List<MissingParam>
            {
                new() { Name = "outputFolder", Description = "PDF输出文件夹", PromptHint = "请提供PDF输出文件夹路径" }
            });

        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");

        List<string> sheetNames;
        if (!string.IsNullOrEmpty(sheetNamesStr))
        {
            try { sheetNames = System.Text.Json.JsonSerializer.Deserialize<List<string>>(sheetNamesStr)!; }
            catch { sheetNames = sheetNamesStr.Split(',').Select(s => s.Trim()).ToList(); }
        }
        else
        {
            sheetNames = _provider.GetWorksheetNames(wb);
        }

        int count = 0;
        foreach (var sn in sheetNames)
        {
            try
            {
                var data = ReadSheetData(wb, sn);
                var safeName = string.Join("_", sn.Split(Path.GetInvalidFileNameChars()));
                var pdfPath = Path.Combine(outputFolder, $"{safeName}.pdf");
                GeneratePdf(data, pdfPath, sn);
                count++;
            }
            catch { }
        }

        return SkillResult.Ok($"批量导出完成，共导出 {count} 个工作表为PDF，保存到: {outputFolder}");
    }

    private SkillResult ExportRangeToPdf(Dictionary<string, object> args)
    {
        var pdfPath = GetStr(args, "pdfPath");
        var rangeAddress = GetStr(args, "rangeAddress");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");

        if (string.IsNullOrEmpty(pdfPath))
            return SkillResult.MissingParamsResult("export_range_to_pdf", new List<MissingParam>
            {
                new() { Name = "pdfPath", Description = "PDF输出文件路径", PromptHint = "请提供PDF输出文件路径" }
            });
        if (string.IsNullOrEmpty(rangeAddress))
            return SkillResult.MissingParamsResult("export_range_to_pdf", new List<MissingParam>
            {
                new() { Name = "rangeAddress", Description = "要导出的区域地址（如A1:D20）", PromptHint = "请提供要导出的区域地址，如 A1:D20" }
            });

        var wb = ResolveWorkbook(fileName);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;

        var data = ReadRangeData(wb, sn, rangeAddress);
        GeneratePdf(data, pdfPath, $"区域 {rangeAddress}");

        return SkillResult.Ok($"区域 '{rangeAddress}' 已导出为PDF:\n{pdfPath}");
    }

    private string? ResolveWorkbook(string fileName)
    {
        if (!string.IsNullOrEmpty(fileName)) return fileName;
        var openWbs = _provider.GetOpenWorkbooks();
        return openWbs.FirstOrDefault();
    }

    private List<string[]> ReadSheetData(string wb, string sn)
    {
        var lastRow = _provider.GetLastRow(wb, sn);
        var lastCol = _provider.GetLastColumn(wb, sn);
        var data = new List<string[]>();
        for (int r = 1; r <= lastRow; r++)
        {
            var row = new string[lastCol];
            for (int c = 1; c <= lastCol; c++)
            {
                row[c - 1] = _provider.GetCellValue(wb, sn, r, c)?.ToString() ?? "";
            }
            data.Add(row);
        }
        return data;
    }

    private List<string[]> ReadRangeData(string wb, string sn, string rangeAddress)
    {
        var rangeData = _provider.GetRangeValues(wb, sn, rangeAddress);
        var rows = rangeData.GetLength(0);
        var cols = rangeData.GetLength(1);
        var data = new List<string[]>();
        for (int r = 0; r < rows; r++)
        {
            var row = new string[cols];
            for (int c = 0; c < cols; c++)
            {
                row[c] = rangeData[r, c]?.ToString() ?? "";
            }
            data.Add(row);
        }
        return data;
    }

    private void GeneratePdf(List<string[]> data, string pdfPath, string title)
    {
        if (data.Count == 0)
        {
            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4);
                    page.Margin(1, Unit.Centimetre);
                    page.Content().Text("（空工作表）");
                });
            }).GeneratePdf(pdfPath);
            return;
        }

        var colCount = data[0].Length;
        var headerRow = data[0];
        var bodyRows = data.Skip(1).ToList();

        Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4.Landscape());
                page.Margin(1, Unit.Centimetre);
                page.Header().Text(title).FontSize(14).Bold();
                page.Content().Table(table =>
                {
                    table.ColumnsDefinition(columns =>
                    {
                        for (int i = 0; i < colCount; i++)
                            columns.RelativeColumn();
                    });

                    table.Header(header =>
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            header.Cell().Element(CellStyle).Text(headerRow[c]);
                        }
                    });

                    foreach (var row in bodyRows)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            var cellText = c < row.Length ? row[c] : "";
                            table.Cell().Element(CellStyle).Text(cellText);
                        }
                    }
                });
            });
        }).GeneratePdf(pdfPath);
    }

    private void GenerateMultiSheetPdf(List<(string SheetName, List<string[]> Data)> sheets, string pdfPath)
    {
        Document.Create(container =>
        {
            foreach (var (sheetName, data) in sheets)
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(1, Unit.Centimetre);
                    page.Header().Text(sheetName).FontSize(14).Bold();

                    if (data.Count == 0)
                    {
                        page.Content().Text("（空工作表）");
                        return;
                    }

                    var colCount = data[0].Length;
                    var headerRow = data[0];
                    var bodyRows = data.Skip(1).ToList();

                    page.Content().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            for (int i = 0; i < colCount; i++)
                                columns.RelativeColumn();
                        });

                        table.Header(header =>
                        {
                            for (int c = 0; c < colCount; c++)
                            {
                                header.Cell().Element(CellStyle).Text(headerRow[c]);
                            }
                        });

                        foreach (var row in bodyRows)
                        {
                            for (int c = 0; c < colCount; c++)
                            {
                                var cellText = c < row.Length ? row[c] : "";
                                table.Cell().Element(CellStyle).Text(cellText);
                            }
                        }
                    });
                });
            }
        }).GeneratePdf(pdfPath);
    }

    private static IContainer CellStyle(IContainer container)
    {
        return container.Border(1).BorderColor(Colors.Grey.Lighten2).PaddingVertical(2).PaddingHorizontal(4);
    }

    private static Dictionary<string, object> P(string[] req, string[] opt)
    {
        var p = new Dictionary<string, object>();
        foreach (var r in req) p[r] = new { type = "string", description = $"{r}（必需）" };
        foreach (var o in opt) p[o] = new { type = "string", description = $"{o}（可选）" };
        return new Dictionary<string, object> { { "type", "object" }, { "properties", p } };
    }
    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
}
