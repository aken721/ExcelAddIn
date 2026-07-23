using System.Collections.Generic;
using System.IO;
using System.Linq;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace TableMagic.Cli.Excel;

internal static class PdfExporter
{
    static PdfExporter()
    {
        QuestPDF.Settings.License = LicenseType.Community;
    }

    public static void ExportSheet(List<string[]> data, string sheetName, string pdfPath)
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
                page.Header().Text(sheetName).FontSize(14).Bold();
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

    public static void ExportWorkbook(List<(string Name, List<string[]> Data)> sheets, string pdfPath)
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
}