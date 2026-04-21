using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelInvoiceSkill : ISkill
    {
        public string Name => "ExcelInvoice";
        public string Description => "发票识别技能，支持XML电子发票导入和发票信息提取";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "import_xml_invoice",
                    Description = "导入单个XML电子发票文件。当用户要求导入电子发票、读取发票信息时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "filePath", new { type = "string", description = "XML发票文件路径" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "filePath" }
                },
                new SkillTool
                {
                    Name = "batch_import_invoices",
                    Description = "批量导入文件夹中的XML电子发票。当用户要求批量导入发票时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "folderPath", new { type = "string", description = "发票文件夹路径" } },
                                { "includeSubfolders", new { type = "boolean", description = "是否包含子目录（默认true）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "folderPath" }
                },
                new SkillTool
                {
                    Name = "get_invoice_fields",
                    Description = "获取发票可提取的字段列表。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "export_invoice_summary",
                    Description = "导出发票汇总表到新工作表。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "sheetName", new { type = "string", description = "输出工作表名称（默认'_FaPiao'）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "clear_invoice_data",
                    Description = "清空发票数据表。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "import_xml_invoice":
                        return await ImportXmlInvoiceAsync(arguments);
                    case "batch_import_invoices":
                        return await BatchImportInvoicesAsync(arguments);
                    case "get_invoice_fields":
                        return GetInvoiceFields();
                    case "export_invoice_summary":
                        return await ExportInvoiceSummaryAsync(arguments);
                    case "clear_invoice_data":
                        return await ClearInvoiceDataAsync();
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelInvoiceSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private async Task<SkillResult> ImportXmlInvoiceAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var filePath = arguments["filePath"].ToString();

                if (!File.Exists(filePath))
                    return new SkillResult { Success = false, Error = $"文件不存在: {filePath}" };

                var invoiceData = ParseXmlInvoice(filePath);
                WriteInvoiceToExcel(invoiceData, filePath);

                return new SkillResult 
                { 
                    Success = true, 
                    Content = $"发票导入成功\n发票号码: {invoiceData.InvoiceNumber}\n开票日期: {invoiceData.IssueDate}\n含税金额: {invoiceData.TotalAmount}" 
                };
            });
        }

        private async Task<SkillResult> BatchImportInvoicesAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var folderPath = arguments["folderPath"].ToString();
                var includeSubfolders = !arguments.ContainsKey("includeSubfolders") || Convert.ToBoolean(arguments["includeSubfolders"]);

                if (!Directory.Exists(folderPath))
                    return new SkillResult { Success = false, Error = $"文件夹不存在: {folderPath}" };

                var searchOption = includeSubfolders ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                var xmlFiles = Directory.GetFiles(folderPath, "*.xml", searchOption).ToList();

                if (xmlFiles.Count == 0)
                    return new SkillResult { Success = false, Error = "文件夹中没有XML发票文件" };

                int successCount = 0;
                int failCount = 0;

                foreach (var file in xmlFiles)
                {
                    try
                    {
                        var invoiceData = ParseXmlInvoice(file);
                        WriteInvoiceToExcel(invoiceData, file);
                        successCount++;
                    }
                    catch
                    {
                        failCount++;
                    }
                }

                return new SkillResult 
                { 
                    Success = true, 
                    Content = $"批量导入完成\n成功: {successCount} 个\n失败: {failCount} 个" 
                };
            });
        }

        private SkillResult GetInvoiceFields()
        {
            var fields = new List<string>
            {
                "发票号码",
                "开票日期",
                "销售方纳税识别号",
                "销售方名称",
                "销售方地址",
                "销售方电话号码",
                "销售方开户银行",
                "销售方银行账号",
                "购买方纳税识别号",
                "购买方名称",
                "购买方地址",
                "购买方电话号码",
                "购买方开户银行",
                "购买方银行账号",
                "不含税价格",
                "税额",
                "含税价格",
                "项目名称",
                "发票类型",
                "发票监制税务机关",
                "电子发票文件路径"
            };

            return new SkillResult 
            { 
                Success = true, 
                Content = "可提取的发票字段：\n" + string.Join("\n", fields.Select((f, i) => $"  {i + 1}. {f}")) 
            };
        }

        private async Task<SkillResult> ExportInvoiceSummaryAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : "_FaPiao汇总";

                var workbook = ThisAddIn.app.ActiveWorkbook;
                Excel.Worksheet sourceSheet = null;

                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == "_FaPiao")
                    {
                        sourceSheet = sheet;
                        break;
                    }
                }

                if (sourceSheet == null)
                    return new SkillResult { Success = false, Error = "未找到发票数据表 '_FaPiao'" };

                var usedRange = sourceSheet.UsedRange;
                int lastRow = usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;

                ThisAddIn.app.ScreenUpdating = false;
                ThisAddIn.app.DisplayAlerts = false;

                var summarySheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                summarySheet.Name = sheetName;

                sourceSheet.UsedRange.Copy(summarySheet.Range["A1"]);

                summarySheet.Activate();

                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;

                return new SkillResult { Success = true, Content = $"发票汇总表已导出到 '{sheetName}'，共 {lastRow - 1} 条记录" };
            });
        }

        private async Task<SkillResult> ClearInvoiceDataAsync()
        {
            return await Task.Run(() =>
            {
                var workbook = ThisAddIn.app.ActiveWorkbook;

                ThisAddIn.app.DisplayAlerts = false;

                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == "_FaPiao")
                    {
                        sheet.Delete();
                        break;
                    }
                }

                ThisAddIn.app.DisplayAlerts = true;

                return new SkillResult { Success = true, Content = "发票数据表已清空" };
            });
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
                data.InvoiceNumber = taxSupervisionInfo.Element("InvoiceNumber")?.Value;
                data.IssueDate = taxSupervisionInfo.Element("IssueTime")?.Value;
                data.TaxBureauName = taxSupervisionInfo.Element("TaxBureauName")?.Value;
            }

            if (eInvoiceData != null)
            {
                var sellerInfo = eInvoiceData.Element("SellerInformation");
                if (sellerInfo != null)
                {
                    data.SellerIdNum = sellerInfo.Element("SellerIdNum")?.Value;
                    data.SellerName = sellerInfo.Element("SellerName")?.Value;
                    data.SellerAddr = sellerInfo.Element("SellerAddr")?.Value;
                    data.SellerTelNum = sellerInfo.Element("SellerTelNum")?.Value;
                    data.SellerBankName = sellerInfo.Element("SellerBankName")?.Value;
                    data.SellerBankAccNum = sellerInfo.Element("SellerBankAccNum")?.Value;
                }

                var buyerInfo = eInvoiceData.Element("BuyerInformation");
                if (buyerInfo != null)
                {
                    data.BuyerIdNum = buyerInfo.Element("BuyerIdNum")?.Value;
                    data.BuyerName = buyerInfo.Element("BuyerName")?.Value;
                    data.BuyerAddr = buyerInfo.Element("BuyerAddr")?.Value;
                    data.BuyerTelNum = buyerInfo.Element("BuyerTelNum")?.Value;
                    data.BuyerBankName = buyerInfo.Element("BuyerBankName")?.Value;
                    data.BuyerBankAccNum = buyerInfo.Element("BuyerBankAccNum")?.Value;
                }

                var basicInfo = eInvoiceData.Element("BasicInformation");
                if (basicInfo != null)
                {
                    data.TotalAmWithoutTax = basicInfo.Element("TotalAmWithoutTax")?.Value;
                    data.TotalTaxAm = basicInfo.Element("TotalTaxAm")?.Value;
                    data.TotalAmount = basicInfo.Element("TotalTax-includedAmount")?.Value;
                }

                var itemInfo = eInvoiceData.Element("IssuItemInformation");
                if (itemInfo != null)
                {
                    data.ItemName = itemInfo.Element("ItemName")?.Value;
                }
            }

            if (header != null)
            {
                var vatLabel = header.Element("InherentLabel")?.Element("GeneralOrSpecialVAT");
                data.InvoiceType = vatLabel?.Element("LabelName")?.Value;
            }

            data.FilePath = filePath;

            return data;
        }

        private bool IsFieldExist(Excel.Worksheet sheet, string fieldName)
        {
            int colCount = sheet.UsedRange.Columns.Count;
            for (int c = 1; c <= colCount; c++)
            {
                if (sheet.Cells[1, c].Value?.ToString() == fieldName)
                    return true;
            }
            return false;
        }

        private void WriteInvoiceToExcel(InvoiceData data, string filePath)
        {
            var workbook = ThisAddIn.app.ActiveWorkbook;

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            Excel.Worksheet invoiceSheet = null;
            bool isFaPiaoSheetExist = false;

            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (string.Equals(sheet.Name, "_FaPiao", StringComparison.OrdinalIgnoreCase) && IsFieldExist(sheet, "电子发票文件路径") && IsFieldExist(sheet, "发票类型"))
                {
                    invoiceSheet = sheet;
                    isFaPiaoSheetExist = true;
                }
                else if (string.Equals(sheet.Name, "_FaPiao", StringComparison.OrdinalIgnoreCase))
                {
                    string newName = "_FaPiao_Original";
                    int suffix = 2;
                    var existingNames = new List<string>();
                    foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
                    while (existingNames.Any(n => string.Equals(n, newName, StringComparison.OrdinalIgnoreCase)))
                    {
                        newName = $"_FaPiao_Original_{suffix}";
                        suffix++;
                    }
                    sheet.Name = newName;
                }
                else
                {
                    continue;
                }
            }

            if (!isFaPiaoSheetExist)
            {
                invoiceSheet = workbook.Worksheets.Add(Before: workbook.ActiveSheet);
                invoiceSheet.Name = "_FaPiao";
                invoiceSheet.Activate();

                var headers = new[] {
                    "发票号码", "开票日期", "销售方纳税识别号", "销售方名称",
                    "销售方地址", "销售方电话号码", "销售方开户银行", "销售方银行账号",
                    "购买方纳税识别号", "购买方名称", "购买方地址", "购买方电话号码",
                    "购买方开户银行", "购买方银行账号", "不含税价格", "税额",
                    "含税价格", "项目名称", "发票类型", "发票监制税务机关", "电子发票文件路径"
                };

                for (int c = 0; c < headers.Length; c++)
                {
                    invoiceSheet.Cells[1, c + 1].Value = headers[c];
                }
            }

            invoiceSheet.Activate();
            long usedRow = invoiceSheet.UsedRange.Rows.Count;
            int newRow = isFaPiaoSheetExist ? (int)usedRow + 1 : 2;

            var values = new object[] {
                data.InvoiceNumber, data.IssueDate, data.SellerIdNum, data.SellerName,
                data.SellerAddr, data.SellerTelNum, data.SellerBankName, data.SellerBankAccNum,
                data.BuyerIdNum, data.BuyerName, data.BuyerAddr, data.BuyerTelNum,
                data.BuyerBankName, data.BuyerBankAccNum, data.TotalAmWithoutTax, data.TotalTaxAm,
                data.TotalAmount, data.ItemName, data.InvoiceType, data.TaxBureauName, filePath
            };

            for (int j = 0; j < values.Length; j++)
            {
                invoiceSheet.Cells[newRow, j + 1].NumberFormat = "@";
                invoiceSheet.Cells[newRow, j + 1].Value = values[j];

                if (invoiceSheet.Cells[1, j + 1].Value?.ToString() == "电子发票文件路径")
                {
                    string str = Path.GetDirectoryName(invoiceSheet.Cells[newRow, j + 1].Value?.ToString());
                    if (!string.IsNullOrEmpty(str))
                    {
                        invoiceSheet.Hyperlinks.Add(invoiceSheet.Cells[newRow, j + 1], str, Type.Missing, Type.Missing, str);
                    }
                }
            }

            invoiceSheet.UsedRange.Columns.AutoFit();
            invoiceSheet.Activate();

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ActiveWorkbook.RefreshAll();
        }

        private class InvoiceData
        {
            public string InvoiceNumber { get; set; }
            public string IssueDate { get; set; }
            public string SellerIdNum { get; set; }
            public string SellerName { get; set; }
            public string SellerAddr { get; set; }
            public string SellerTelNum { get; set; }
            public string SellerBankName { get; set; }
            public string SellerBankAccNum { get; set; }
            public string BuyerIdNum { get; set; }
            public string BuyerName { get; set; }
            public string BuyerAddr { get; set; }
            public string BuyerTelNum { get; set; }
            public string BuyerBankName { get; set; }
            public string BuyerBankAccNum { get; set; }
            public string TotalAmWithoutTax { get; set; }
            public string TotalTaxAm { get; set; }
            public string TotalAmount { get; set; }
            public string ItemName { get; set; }
            public string InvoiceType { get; set; }
            public string TaxBureauName { get; set; }
            public string FilePath { get; set; }
        }
    }
}
