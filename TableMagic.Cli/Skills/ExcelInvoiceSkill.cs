using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml.Linq;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelInvoiceSkill : ISkill
{
    private readonly IExcelProvider _provider;
    public ExcelInvoiceSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelInvoice";
    public string Description => "发票识别：XML发票导入、批量导入、OCR识别";

    private static readonly string[] Headers = {
        "发票号码", "开票日期", "销售方纳税识别号", "销售方名称",
        "销售方地址", "销售方电话号码", "销售方开户银行", "销售方银行账号",
        "购买方纳税识别号", "购买方名称", "购买方地址", "购买方电话号码",
        "购买方开户银行", "购买方银行账号", "不含税价格", "税额",
        "含税价格", "项目名称", "发票类型", "发票监制税务机关", "电子发票文件路径"
    };

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "import_xml_invoice", Description = "导入XML格式发票数据到Excel",
                Parameters = P(new[]{"xmlPath"}, new[]{"outputFileName","outputSheetName"}), RequiredParameters = new List<string>{"xmlPath"} },
            new() { Name = "batch_import_invoices", Description = "批量导入文件夹中所有XML发票",
                Parameters = P(new[]{"folderPath"}, new[]{"outputFileName","outputSheetName","includeSubfolders"}), RequiredParameters = new List<string>{"folderPath"} },
            new() { Name = "get_invoice_fields", Description = "获取发票可提取的字段列表",
                Parameters = P(Array.Empty<string>(), Array.Empty<string>()), RequiredParameters = new List<string>() },
            new() { Name = "export_invoice_summary", Description = "导出发票汇总表",
                Parameters = P(new[]{"outputFileName"}, new[]{"outputSheetName"}), RequiredParameters = new List<string>{"outputFileName"} },
            new() { Name = "ocr_invoice", Description = "OCR识别发票图片（需要PaddleOCR运行时）",
                Parameters = P(new[]{"imagePath"}, new[]{"outputFileName","outputSheetName"}), RequiredParameters = new List<string>{"imagePath"} }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "import_xml_invoice" => ImportXmlInvoice(arguments),
                "batch_import_invoices" => BatchImport(arguments),
                "get_invoice_fields" => GetInvoiceFields(),
                "export_invoice_summary" => ExportSummary(arguments),
                "ocr_invoice" => OcrInvoice(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult ImportXmlInvoice(Dictionary<string, object> args)
    {
        var xmlPath = GetStr(args, "xmlPath");
        var outputFileName = GetStr(args, "outputFileName");
        var outputSheetName = GetStr(args, "outputSheetName") ?? "_FaPiao";

        if (string.IsNullOrEmpty(xmlPath))
            return SkillResult.MissingParamsResult("import_xml_invoice", new List<MissingParam>
            {
                new() { Name = "xmlPath", Description = "XML发票文件路径", PromptHint = "请提供XML发票文件路径" }
            });

        if (!File.Exists(xmlPath))
            return SkillResult.FromError($"文件不存在: {xmlPath}");

        var invoiceData = ParseXmlInvoice(xmlPath);
        var wb = ResolveOrCreateWorkbook(outputFileName, outputSheetName);
        WriteInvoiceToSheet(wb, outputSheetName, invoiceData);
        _provider.SaveWorkbook(wb);

        return SkillResult.Ok($"发票导入成功\n发票号码: {invoiceData.InvoiceNumber}\n开票日期: {invoiceData.IssueDate}\n含税金额: {invoiceData.TotalAmount}");
    }

    private SkillResult BatchImport(Dictionary<string, object> args)
    {
        var folder = GetStr(args, "folderPath");
        var outputFileName = GetStr(args, "outputFileName");
        var outputSheetName = GetStr(args, "outputSheetName") ?? "_FaPiao";
        var includeSubfolders = GetBool(args, "includeSubfolders", true);

        if (string.IsNullOrEmpty(folder))
            return SkillResult.MissingParamsResult("batch_import_invoices", new List<MissingParam>
            {
                new() { Name = "folderPath", Description = "发票文件夹路径", PromptHint = "请提供包含XML发票的文件夹路径" }
            });

        if (!Directory.Exists(folder))
            return SkillResult.FromError($"文件夹不存在: {folder}");

        var option = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var xmlFiles = Directory.GetFiles(folder, "*.xml", option);

        if (xmlFiles.Length == 0)
            return SkillResult.Ok("文件夹中没有XML发票文件");

        var wb = ResolveOrCreateWorkbook(outputFileName, outputSheetName);
        int successCount = 0, failCount = 0;

        foreach (var file in xmlFiles)
        {
            try
            {
                var invoiceData = ParseXmlInvoice(file);
                WriteInvoiceToSheet(wb, outputSheetName, invoiceData);
                successCount++;
            }
            catch { failCount++; }
        }

        _provider.SaveWorkbook(wb);
        return SkillResult.Ok($"批量导入完成\n成功: {successCount} 个\n失败: {failCount} 个");
    }

    private SkillResult GetInvoiceFields() =>
        SkillResult.Ok("发票可提取字段：\n" + string.Join("\n", Headers.Select((f, i) => $"  {i + 1}. {f}")));

    private SkillResult ExportSummary(Dictionary<string, object> args)
    {
        var outputFileName = GetStr(args, "outputFileName");
        var outputSheetName = GetStr(args, "outputSheetName") ?? "_FaPiao汇总";

        var wb = ResolveWorkbook(outputFileName!);
        if (wb == null) return SkillResult.FromError("没有打开的工作簿");

        var sheets = _provider.GetWorksheetNames(wb);
        if (!sheets.Contains("_FaPiao"))
            return SkillResult.FromError("未找到发票数据表 '_FaPiao'，请先导入发票");

        var lastRow = _provider.GetLastRow(wb, "_FaPiao");
        var lastCol = _provider.GetLastColumn(wb, "_FaPiao");

        if (!sheets.Contains(outputSheetName))
            _provider.CreateWorksheet(wb, outputSheetName);

        var data = _provider.GetRangeValues(wb, "_FaPiao", $"A1:{GetColumnLetter(lastCol)}{lastRow}");
        _provider.SetRangeValues(wb, outputSheetName, $"A1:{GetColumnLetter(lastCol)}{lastRow}", data);
        _provider.SaveWorkbook(wb);

        return SkillResult.Ok($"发票汇总表已导出到 '{outputSheetName}'，共 {lastRow - 1} 条记录");
    }

    private SkillResult OcrInvoice(Dictionary<string, object> args)
    {
        var imagePath = GetStr(args, "imagePath");
        var outputFileName = GetStr(args, "outputFileName");
        var outputSheetName = GetStr(args, "outputSheetName") ?? "_FaPiao";

        if (string.IsNullOrEmpty(imagePath))
            return SkillResult.MissingParamsResult("ocr_invoice", new List<MissingParam>
            {
                new() { Name = "imagePath", Description = "发票图片路径", PromptHint = "请提供发票图片路径（支持 png/jpg/bmp 格式）" }
            });

        if (!File.Exists(imagePath))
            return SkillResult.FromError($"文件不存在: {imagePath}");

        var ext = Path.GetExtension(imagePath).ToLowerInvariant();
        if (ext != ".png" && ext != ".jpg" && ext != ".jpeg" && ext != ".bmp" && ext != ".pdf")
            return SkillResult.FromError($"不支持的文件格式: {ext}，仅支持 png/jpg/bmp/pdf");

        var paddleResult = DetectPaddleOCR();
        if (!paddleResult.Found)
            return SkillResult.FromError($"未检测到 PaddleOCR 运行时。\n\n{paddleResult.InstallGuide}");

        try
        {
            var ocrOutput = RunPaddleOCR(paddleResult, imagePath);
            if (ocrOutput == null || ocrOutput.Count == 0)
                return SkillResult.FromError("OCR识别未返回任何文字内容，请检查图片是否清晰");

            var invoiceData = ExtractInvoiceFromOcr(ocrOutput, imagePath);
            var wb = ResolveOrCreateWorkbook(outputFileName, outputSheetName);
            WriteInvoiceToSheet(wb, outputSheetName, invoiceData);
            _provider.SaveWorkbook(wb);

            return SkillResult.Ok($"OCR发票识别完成\n发票号码: {invoiceData.InvoiceNumber}\n开票日期: {invoiceData.IssueDate}\n含税金额: {invoiceData.TotalAmount}\n识别文字行数: {ocrOutput.Count}");
        }
        catch (Exception ex)
        {
            return SkillResult.FromError($"OCR识别失败: {ex.Message}");
        }
    }

    private static PaddleDetectionResult DetectPaddleOCR()
    {
        var candidates = new[] { "paddleocr", "python", "python3", "py" };

        foreach (var cmd in candidates)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = cmd,
                    Arguments = cmd == "paddleocr" ? "--version" : "-c \"import paddleocr; print(paddleocr.__version__)\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var process = Process.Start(psi);
                if (process == null) continue;
                process.WaitForExit(10000);

                if (process.ExitCode == 0)
                    return new PaddleDetectionResult { Found = true, Command = cmd };
            }
            catch { }
        }

        return new PaddleDetectionResult { Found = false };
    }

    private static List<string> RunPaddleOCR(PaddleDetectionResult paddle, string imagePath)
    {
        var isPdf = imagePath.ToLowerInvariant().EndsWith(".pdf");
        var script = isPdf
            ? $"import paddleocr; ocr=paddleocr.PaddleOCR(use_angle_cls=True,lang='ch'); import fitz; doc=fitz.open(r'{imagePath}'); page=doc[0]; pix=page.get_pixmap(); pix.save('_tmp_ocr.png'); result=ocr.ocr('_tmp_ocr.png',cls=True); import os; os.remove('_tmp_ocr.png'); [print(line[1][0]) for line in result[0]]"
            : $"import paddleocr; ocr=paddleocr.PaddleOCR(use_angle_cls=True,lang='ch'); result=ocr.ocr(r'{imagePath}',cls=True); [print(line[1][0]) for line in result[0]]";

        var psi = new ProcessStartInfo
        {
            FileName = paddle.Command == "paddleocr" ? "paddleocr" : paddle.Command,
            Arguments = paddle.Command == "paddleocr"
                ? $"ocr --image_dir \"{imagePath}\" --use_angle_cls true --lang ch"
                : $"-c \"{script}\"",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi);
        if (process == null) throw new Exception("无法启动 PaddleOCR 进程");

        var output = process.StandardOutput.ReadToEnd();
        process.WaitForExit(60000);

        if (process.ExitCode != 0)
        {
            var err = process.StandardError.ReadToEnd();
            throw new Exception($"PaddleOCR 执行失败 (exit code {process.ExitCode}): {err}");
        }

        var lines = new List<string>();
        foreach (var line in output.Split('\n'))
        {
            var trimmed = line.Trim();
            if (!string.IsNullOrEmpty(trimmed))
                lines.Add(trimmed);
        }

        return lines;
    }

    private static InvoiceData ExtractInvoiceFromOcr(List<string> ocrLines, string filePath)
    {
        var data = new InvoiceData { FilePath = filePath };
        var fullText = string.Join("\n", ocrLines);

        data.InvoiceNumber = ExtractByPatterns(ocrLines, new[]
        {
            @"发票号码[：:]\s*(\d{8,20})",
            @"No[.：:]\s*(\d{8,20})",
            @"(\d{8,20})"
        }) ?? "";

        data.IssueDate = ExtractByPatterns(ocrLines, new[]
        {
            @"开票日期[：:]\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)",
            @"开票日期[：:]\s*(\d{4}[-/]\d{1,2}[-/]\d{1,2})",
            @"(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)"
        }) ?? "";

        data.SellerName = ExtractByPatterns(ocrLines, new[]
        {
            @"销售方[：:]*\s*名称[：:]\s*(.+?)(?:\s|$)",
            @"销\s*售\s*方.*?名\s*称[：:]\s*(.+?)(?:\s|$)"
        }) ?? "";

        data.SellerIdNum = ExtractByPatterns(ocrLines, new[]
        {
            @"销售方.*?纳税识别号[：:]\s*([A-Z0-9]{15,20})",
            @"销\s*售\s*方.*?税\s*号[：:]\s*([A-Z0-9]{15,20})"
        }) ?? "";

        data.BuyerName = ExtractByPatterns(ocrLines, new[]
        {
            @"购买方[：:]*\s*名称[：:]\s*(.+?)(?:\s|$)",
            @"购\s*买\s*方.*?名\s*称[：:]\s*(.+?)(?:\s|$)"
        }) ?? "";

        data.BuyerIdNum = ExtractByPatterns(ocrLines, new[]
        {
            @"购买方.*?纳税识别号[：:]\s*([A-Z0-9]{15,20})",
            @"购\s*买\s*方.*?税\s*号[：:]\s*([A-Z0-9]{15,20})"
        }) ?? "";

        var totalMatch = ExtractByPatterns(ocrLines, new[]
        {
            @"价税合计[（(]大写[)）][：:].*?[（(]小写[)）][：:]\s*[\￥¥]\s*([\d,.]+)",
            @"价税合计.*?[\￥¥]\s*([\d,.]+)",
            @"合\s*计.*?[\￥¥]\s*([\d,.]+)"
        });
        data.TotalAmount = totalMatch ?? "";

        var amountMatch = ExtractByPatterns(ocrLines, new[]
        {
            @"金\s*额.*?[\￥¥]\s*([\d,.]+)",
            @"不含税金额[：:]\s*[\￥¥]?\s*([\d,.]+)"
        });
        data.TotalAmWithoutTax = amountMatch ?? "";

        var taxMatch = ExtractByPatterns(ocrLines, new[]
        {
            @"税\s*额.*?[\￥¥]\s*([\d,.]+)"
        });
        data.TotalTaxAm = taxMatch ?? "";

        if (fullText.Contains("增值税专用发票"))
            data.InvoiceType = "增值税专用发票";
        else if (fullText.Contains("增值税普通发票") || fullText.Contains("增值税电子普通发票"))
            data.InvoiceType = "增值税普通发票";
        else if (fullText.Contains("电子发票"))
            data.InvoiceType = "电子发票";

        return data;
    }

    private static string? ExtractByPatterns(List<string> lines, string[] patterns)
    {
        foreach (var pattern in patterns)
        {
            try
            {
                var regex = new System.Text.RegularExpressions.Regex(pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                foreach (var line in lines)
                {
                    var match = regex.Match(line);
                    if (match.Success && match.Groups.Count > 1)
                        return match.Groups[1].Value.Trim();
                }
            }
            catch { }
        }
        return null;
    }

    private string? ResolveWorkbook(string fileName)
    {
        if (!string.IsNullOrEmpty(fileName)) return fileName;
        return _provider.GetOpenWorkbooks().FirstOrDefault();
    }

    private string ResolveOrCreateWorkbook(string? fileName, string sheetName)
    {
        if (!string.IsNullOrEmpty(fileName))
        {
            var openWbs = _provider.GetOpenWorkbooks();
            if (openWbs.Contains(fileName)) return fileName;
            return _provider.CreateWorkbook(fileName, sheetName);
        }
        var wb = _provider.GetOpenWorkbooks().FirstOrDefault();
        if (wb != null) return wb;
        return _provider.CreateWorkbook("invoices.xlsx", sheetName);
    }

    private void WriteInvoiceToSheet(string wb, string sheetName, InvoiceData data)
    {
        var sheets = _provider.GetWorksheetNames(wb);
        if (!sheets.Contains(sheetName))
            _provider.CreateWorksheet(wb, sheetName);

        var lastCol = _provider.GetLastColumn(wb, sheetName);
        if (lastCol == 0 || lastCol < Headers.Length)
        {
            for (int c = 0; c < Headers.Length; c++)
                _provider.SetCellValue(wb, sheetName, 1, c + 1, Headers[c]);
        }

        var lastRow = _provider.GetLastRow(wb, sheetName);
        var newRow = lastRow < 1 ? 2 : lastRow + 1;

        var values = new object[] {
            data.InvoiceNumber ?? "", data.IssueDate ?? "", data.SellerIdNum ?? "", data.SellerName ?? "",
            data.SellerAddr ?? "", data.SellerTelNum ?? "", data.SellerBankName ?? "", data.SellerBankAccNum ?? "",
            data.BuyerIdNum ?? "", data.BuyerName ?? "", data.BuyerAddr ?? "", data.BuyerTelNum ?? "",
            data.BuyerBankName ?? "", data.BuyerBankAccNum ?? "", data.TotalAmWithoutTax ?? "", data.TotalTaxAm ?? "",
            data.TotalAmount ?? "", data.ItemName ?? "", data.InvoiceType ?? "", data.TaxBureauName ?? "", data.FilePath ?? ""
        };

        for (int j = 0; j < values.Length; j++)
            _provider.SetCellValue(wb, sheetName, newRow, j + 1, values[j]?.ToString() ?? "");
    }

    private InvoiceData ParseXmlInvoice(string filePath)
    {
        var doc = XElement.Load(filePath);
        var data = new InvoiceData();

        var taxSupervisionInfo = doc.Element("TaxSupervisionInfo");
        var eInvoiceData = doc.Element("EInvoiceData");
        var header = doc.Element("Header");

        if (taxSupervisionInfo != null)
        {
            data.InvoiceNumber = taxSupervisionInfo.Element("InvoiceNumber")?.Value ?? "";
            data.IssueDate = taxSupervisionInfo.Element("IssueTime")?.Value ?? "";
            data.TaxBureauName = taxSupervisionInfo.Element("TaxBureauName")?.Value ?? "";
        }

        if (eInvoiceData != null)
        {
            var sellerInfo = eInvoiceData.Element("SellerInformation");
            if (sellerInfo != null)
            {
                data.SellerIdNum = sellerInfo.Element("SellerIdNum")?.Value ?? "";
                data.SellerName = sellerInfo.Element("SellerName")?.Value ?? "";
                data.SellerAddr = sellerInfo.Element("SellerAddr")?.Value ?? "";
                data.SellerTelNum = sellerInfo.Element("SellerTelNum")?.Value ?? "";
                data.SellerBankName = sellerInfo.Element("SellerBankName")?.Value ?? "";
                data.SellerBankAccNum = sellerInfo.Element("SellerBankAccNum")?.Value ?? "";
            }

            var buyerInfo = eInvoiceData.Element("BuyerInformation");
            if (buyerInfo != null)
            {
                data.BuyerIdNum = buyerInfo.Element("BuyerIdNum")?.Value ?? "";
                data.BuyerName = buyerInfo.Element("BuyerName")?.Value ?? "";
                data.BuyerAddr = buyerInfo.Element("BuyerAddr")?.Value ?? "";
                data.BuyerTelNum = buyerInfo.Element("BuyerTelNum")?.Value ?? "";
                data.BuyerBankName = buyerInfo.Element("BuyerBankName")?.Value ?? "";
                data.BuyerBankAccNum = buyerInfo.Element("BuyerBankAccNum")?.Value ?? "";
            }

            var basicInfo = eInvoiceData.Element("BasicInformation");
            if (basicInfo != null)
            {
                data.TotalAmWithoutTax = basicInfo.Element("TotalAmWithoutTax")?.Value ?? "";
                data.TotalTaxAm = basicInfo.Element("TotalTaxAm")?.Value ?? "";
                data.TotalAmount = basicInfo.Element("TotalTax-includedAmount")?.Value ?? "";
            }

            var itemInfo = eInvoiceData.Element("IssuItemInformation");
            if (itemInfo != null)
            {
                data.ItemName = itemInfo.Element("ItemName")?.Value ?? "";
            }
        }

        if (header != null)
        {
            var vatLabel = header.Element("InherentLabel")?.Element("GeneralOrSpecialVAT");
            data.InvoiceType = vatLabel?.Element("LabelName")?.Value ?? "";
        }

        data.FilePath = filePath;
        return data;
    }

    private static string GetColumnLetter(int columnNumber)
    {
        string letter = "";
        while (columnNumber > 0)
        {
            int mod = (columnNumber - 1) % 26;
            letter = Convert.ToChar(65 + mod) + letter;
            columnNumber = (columnNumber - mod) / 26;
        }
        return letter;
    }

    private static Dictionary<string, object> P(string[] req, string[] opt)
    {
        var p = new Dictionary<string, object>();
        foreach (var r in req) p[r] = new { type = "string", description = $"{r}（必需）" };
        foreach (var o in opt) p[o] = new { type = "string", description = $"{o}（可选）" };
        return new Dictionary<string, object> { { "type", "object" }, { "properties", p } };
    }
    private static string? GetStr(Dictionary<string, object> a, string k) => a.ContainsKey(k) ? a[k]?.ToString() : null;
    private static bool GetBool(Dictionary<string, object> a, string k, bool def = false) => a.ContainsKey(k) && bool.TryParse(a[k]?.ToString(), out var v) ? v : def;

    private class InvoiceData
    {
        public string InvoiceNumber { get; set; } = "";
        public string IssueDate { get; set; } = "";
        public string SellerIdNum { get; set; } = "";
        public string SellerName { get; set; } = "";
        public string SellerAddr { get; set; } = "";
        public string SellerTelNum { get; set; } = "";
        public string SellerBankName { get; set; } = "";
        public string SellerBankAccNum { get; set; } = "";
        public string BuyerIdNum { get; set; } = "";
        public string BuyerName { get; set; } = "";
        public string BuyerAddr { get; set; } = "";
        public string BuyerTelNum { get; set; } = "";
        public string BuyerBankName { get; set; } = "";
        public string BuyerBankAccNum { get; set; } = "";
        public string TotalAmWithoutTax { get; set; } = "";
        public string TotalTaxAm { get; set; } = "";
        public string TotalAmount { get; set; } = "";
        public string ItemName { get; set; } = "";
        public string InvoiceType { get; set; } = "";
        public string TaxBureauName { get; set; } = "";
        public string FilePath { get; set; } = "";
    }

    private class PaddleDetectionResult
    {
        public bool Found { get; set; }
        public string Command { get; set; } = "";

        public string InstallGuide =>
            "请安装 PaddleOCR 运行时后再使用 OCR 发票识别功能。\n\n" +
            "安装方法（任选其一）：\n\n" +
            "方法一：pip 安装（推荐）\n" +
            "  pip install paddlepaddle paddleocr\n\n" +
            "方法二：conda 安装\n" +
            "  conda install paddlepaddle\n" +
            "  pip install paddleocr\n\n" +
            "方法三：使用国内镜像加速\n" +
            "  pip install paddlepaddle -i https://mirror.baidu.com/pypi/simple\n" +
            "  pip install paddleocr\n\n" +
            "验证安装：\n" +
            "  python -c \"import paddleocr; print(paddleocr.__version__)\"\n\n" +
            "注意：\n" +
            "- 需要 Python 3.7+ 环境\n" +
            "- PDF 发票识别还需安装 PyMuPDF: pip install pymupdf\n" +
            "- 首次运行会自动下载模型文件（约 100MB）\n" +
            "- 如已安装但仍检测不到，请确保 python/pip 在系统 PATH 中";
    }
}
