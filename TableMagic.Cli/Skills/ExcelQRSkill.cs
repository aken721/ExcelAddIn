using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using SkiaSharp;
using TableMagic.Cli.Excel;
using ZXing;
using ZXing.QrCode.Internal;
using ZXing.SkiaSharp;

namespace TableMagic.Cli.Skills;

public class ExcelQRSkill : ISkill
{
    private readonly IExcelProvider _provider;
    public ExcelQRSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelQR";
    public string Description => "二维码/条形码：生成、扫描、识别";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "generate_qr_code", Description = "为指定列的数据生成二维码图片文件",
                Parameters = P(new[]{"columnNames","outputFolder"}, new[]{"fileName","sheetName","size"}), RequiredParameters = new List<string>{"columnNames","outputFolder"} },
            new() { Name = "generate_barcode", Description = "为指定列的数据生成条形码图片文件",
                Parameters = P(new[]{"columnName","outputFolder"}, new[]{"fileName","sheetName","width","height"}), RequiredParameters = new List<string>{"columnName","outputFolder"} },
            new() { Name = "scan_qr_code", Description = "扫描图片文件中的二维码并返回内容",
                Parameters = P(new[]{"imagePaths"}, new[]{"outputFileName","outputSheetName"}), RequiredParameters = new List<string>{"imagePaths"} },
            new() { Name = "scan_qr_code_folder", Description = "批量扫描文件夹中所有图片的二维码",
                Parameters = P(new[]{"folderPath"}, new[]{"includeSubfolders","outputFileName","outputSheetName"}), RequiredParameters = new List<string>{"folderPath"} }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "generate_qr_code" => GenerateQrCode(arguments),
                "generate_barcode" => GenerateBarcode(arguments),
                "scan_qr_code" => ScanQrCode(arguments),
                "scan_qr_code_folder" => ScanQrCodeFolder(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult GenerateQrCode(Dictionary<string, object> args)
    {
        var columnNamesStr = GetStr(args, "columnNames");
        var outputFolder = GetStr(args, "outputFolder");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");
        var size = GetInt(args, "size", 200);

        if (string.IsNullOrEmpty(columnNamesStr))
            return SkillResult.MissingParamsResult("generate_qr_code", new List<MissingParam>
            {
                new() { Name = "columnNames", Description = "要生成二维码的列名列表（JSON数组格式）", PromptHint = "请提供要生成二维码的列名，如 [\"网址\",\"编号\"]" }
            });
        if (string.IsNullOrEmpty(outputFolder))
            return SkillResult.MissingParamsResult("generate_qr_code", new List<MissingParam>
            {
                new() { Name = "outputFolder", Description = "二维码图片输出文件夹", PromptHint = "请提供二维码图片的输出文件夹路径" }
            });

        List<string> columnNames;
        try { columnNames = JsonSerializer.Deserialize<List<string>>(columnNamesStr)!; }
        catch { columnNames = new List<string> { columnNamesStr }; }

        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        var wb = string.IsNullOrEmpty(fileName) ? _provider.GetOpenWorkbooks().FirstOrDefault() : fileName;
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;
        var lastRow = _provider.GetLastRow(wb, sn);
        var lastCol = _provider.GetLastColumn(wb, sn);

        var headerRow = new Dictionary<string, int>();
        for (int c = 1; c <= lastCol; c++)
        {
            var val = _provider.GetCellValue(wb, sn, 1, c)?.ToString();
            if (!string.IsNullOrEmpty(val)) headerRow[val] = c;
        }

        var colIndices = new List<int>();
        foreach (var colName in columnNames)
        {
            if (headerRow.TryGetValue(colName, out var idx))
                colIndices.Add(idx);
        }

        if (colIndices.Count == 0)
            return SkillResult.FromError($"未找到指定的列: {string.Join(", ", columnNames)}");

        var writer = new BarcodeWriter
        {
            Format = BarcodeFormat.QR_CODE,
            Options = new ZXing.QrCode.QrCodeEncodingOptions
            {
                Height = size,
                Width = size,
                CharacterSet = "UTF-8",
                ErrorCorrection = ErrorCorrectionLevel.H,
                Margin = 1
            }
        };

        int generatedCount = 0;
        for (int r = 2; r <= lastRow; r++)
        {
            string data;
            if (colIndices.Count > 1)
            {
                var parts = new List<string>();
                foreach (var idx in colIndices)
                {
                    var key = _provider.GetCellValue(wb, sn, 1, idx)?.ToString() ?? "";
                    var value = _provider.GetCellValue(wb, sn, r, idx)?.ToString() ?? "";
                    parts.Add($"{key}:{value}");
                }
                data = string.Join(";", parts);
            }
            else
            {
                data = _provider.GetCellValue(wb, sn, r, colIndices[0])?.ToString() ?? "";
            }

            if (string.IsNullOrEmpty(data)) continue;

            try
            {
                using var bitmap = writer.Write(data);
                using var skImage = SKImage.FromBitmap(bitmap);
                using var imgData = skImage.Encode(SKEncodedImageFormat.Png, 100);
                var outPath = Path.Combine(outputFolder, $"qr_row{r}_{Guid.NewGuid():N}.png");
                File.WriteAllBytes(outPath, imgData.ToArray());
                generatedCount++;
            }
            catch { }
        }

        return SkillResult.Ok($"二维码生成完成，共生成 {generatedCount} 个二维码图片，保存到: {outputFolder}");
    }

    private SkillResult GenerateBarcode(Dictionary<string, object> args)
    {
        var columnName = GetStr(args, "columnName");
        var outputFolder = GetStr(args, "outputFolder");
        var fileName = GetStr(args, "fileName");
        var sheetName = GetStr(args, "sheetName");
        var width = GetInt(args, "width", 300);
        var height = GetInt(args, "height", 100);

        if (string.IsNullOrEmpty(columnName))
            return SkillResult.MissingParamsResult("generate_barcode", new List<MissingParam>
            {
                new() { Name = "columnName", Description = "要生成条形码的列名", PromptHint = "请提供要生成条形码的列名" }
            });
        if (string.IsNullOrEmpty(outputFolder))
            return SkillResult.MissingParamsResult("generate_barcode", new List<MissingParam>
            {
                new() { Name = "outputFolder", Description = "条形码图片输出文件夹", PromptHint = "请提供条形码图片的输出文件夹路径" }
            });

        if (!Directory.Exists(outputFolder))
            Directory.CreateDirectory(outputFolder);

        var wb = string.IsNullOrEmpty(fileName) ? _provider.GetOpenWorkbooks().FirstOrDefault() : fileName;
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");
        var sn = string.IsNullOrEmpty(sheetName) ? _provider.GetActiveWorksheetName(wb) : sheetName;
        var lastRow = _provider.GetLastRow(wb, sn);
        var lastCol = _provider.GetLastColumn(wb, sn);

        var headerRow = new Dictionary<string, int>();
        for (int c = 1; c <= lastCol; c++)
        {
            var val = _provider.GetCellValue(wb, sn, 1, c)?.ToString();
            if (!string.IsNullOrEmpty(val)) headerRow[val] = c;
        }

        if (!headerRow.TryGetValue(columnName, out var colIdx))
            return SkillResult.FromError($"未找到列: {columnName}");

        var writer = new BarcodeWriter
        {
            Format = BarcodeFormat.CODE_128,
            Options = new ZXing.Common.EncodingOptions
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
            var value = _provider.GetCellValue(wb, sn, r, colIdx)?.ToString();
            if (string.IsNullOrEmpty(value)) continue;
            if (!asciiPattern.IsMatch(value)) continue;

            try
            {
                using var bitmap = writer.Write(value);
                using var skImage = SKImage.FromBitmap(bitmap);
                using var imgData = skImage.Encode(SKEncodedImageFormat.Png, 100);
                var outPath = Path.Combine(outputFolder, $"barcode_row{r}_{Guid.NewGuid():N}.png");
                File.WriteAllBytes(outPath, imgData.ToArray());
                generatedCount++;
            }
            catch { }
        }

        return SkillResult.Ok($"条形码生成完成，共生成 {generatedCount} 个条形码图片，保存到: {outputFolder}");
    }

    private SkillResult ScanQrCode(Dictionary<string, object> args)
    {
        var imagePathsStr = GetStr(args, "imagePaths");
        var outputFileName = GetStr(args, "outputFileName");
        var outputSheetName = GetStr(args, "outputSheetName") ?? "二维码识别结果";

        if (string.IsNullOrEmpty(imagePathsStr))
            return SkillResult.MissingParamsResult("scan_qr_code", new List<MissingParam>
            {
                new() { Name = "imagePaths", Description = "图片文件路径列表（JSON数组格式）", PromptHint = "请提供要扫描的图片路径列表" }
            });

        List<string> imagePaths;
        try { imagePaths = JsonSerializer.Deserialize<List<string>>(imagePathsStr)!; }
        catch { imagePaths = new List<string> { imagePathsStr }; }

        var reader = new BarcodeReader();
        var results = new List<(string Path, string Content)>();

        foreach (var path in imagePaths)
        {
            if (!File.Exists(path)) continue;
            try
            {
                using var stream = File.OpenRead(path);
                using var skBitmap = SKBitmap.Decode(stream);
                if (skBitmap == null) { results.Add((path, "无法解码图片")); continue; }
                var result = reader.Decode(skBitmap);
                results.Add((path, result?.Text ?? "无法识别"));
            }
            catch (Exception ex) { results.Add((path, $"识别失败: {ex.Message}")); }
        }

        if (!string.IsNullOrEmpty(outputFileName))
        {
            try
            {
                var wb = outputFileName;
                if (!_provider.GetOpenWorkbooks().Contains(wb))
                    wb = _provider.CreateWorkbook(outputFileName, outputSheetName);
                var sn = outputSheetName;
                _provider.SetCellValue(wb, sn, 1, 1, "文件路径");
                _provider.SetCellValue(wb, sn, 1, 2, "识别内容");
                for (int i = 0; i < results.Count; i++)
                {
                    _provider.SetCellValue(wb, sn, i + 2, 1, results[i].Path);
                    _provider.SetCellValue(wb, sn, i + 2, 2, results[i].Content);
                }
                _provider.SaveWorkbook(wb);
            }
            catch { }
        }

        var content = string.Join("\n", results.Select((r, i) => $"{i + 1}. {r.Path}: {r.Content}"));
        return SkillResult.Ok($"二维码扫描完成，共扫描 {results.Count} 张图片:\n{content}");
    }

    private SkillResult ScanQrCodeFolder(Dictionary<string, object> args)
    {
        var folderPath = GetStr(args, "folderPath");
        var includeSubfolders = GetBool(args, "includeSubfolders", true);
        var outputFileName = GetStr(args, "outputFileName");
        var outputSheetName = GetStr(args, "outputSheetName") ?? "二维码识别结果";

        if (string.IsNullOrEmpty(folderPath))
            return SkillResult.MissingParamsResult("scan_qr_code_folder", new List<MissingParam>
            {
                new() { Name = "folderPath", Description = "图片文件夹路径", PromptHint = "请提供包含图片的文件夹路径" }
            });

        if (!Directory.Exists(folderPath))
            return SkillResult.FromError($"文件夹不存在: {folderPath}");

        var extensions = new[] { "*.png", "*.jpg", "*.jpeg", "*.bmp", "*.gif" };
        var files = new List<string>();
        var option = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        foreach (var ext in extensions)
            files.AddRange(Directory.GetFiles(folderPath, ext, option));

        if (files.Count == 0)
            return SkillResult.Ok($"文件夹中没有找到图片文件: {folderPath}");

        var scanArgs = new Dictionary<string, object>
        {
            { "imagePaths", JsonSerializer.Serialize(files) },
            { "outputFileName", outputFileName ?? "" },
            { "outputSheetName", outputSheetName }
        };
        return ScanQrCode(scanArgs);
    }

    private static Dictionary<string, object> P(string[] req, string[] opt)
    {
        var p = new Dictionary<string, object>();
        foreach (var r in req) p[r] = new { type = "string", description = $"{r}（必需）" };
        foreach (var o in opt) p[o] = new { type = "string", description = $"{o}（可选）" };
        return new Dictionary<string, object> { { "type", "object" }, { "properties", p } };
    }
    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
    private static int GetInt(Dictionary<string, object> a, string k, int def = 0) => a.ContainsKey(k) && int.TryParse(a[k]?.ToString(), out var v) ? v : def;
    private static bool GetBool(Dictionary<string, object> a, string k, bool def = false) => a.ContainsKey(k) && bool.TryParse(a[k]?.ToString(), out var v) ? v : def;
}
