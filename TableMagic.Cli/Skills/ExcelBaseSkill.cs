using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelBaseSkill : ISkill
{
    private readonly IExcelProvider _provider;
    private readonly SkillManager _skillManager;

    public ExcelBaseSkill(IExcelProvider provider, SkillManager skillManager) { _provider = provider; _skillManager = skillManager; }

    public string Name => "ExcelBase";
    public string Description => "Excel基础操作，获取工作表名称等信息，含批量调用和复合工具";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "get_worksheet_names",
                Description = "获取当前工作簿中所有工作表名称",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string>()
            },
            new()
            {
                Name = "get_open_workbooks",
                Description = "获取当前打开的所有工作簿列表",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>() }
                },
                RequiredParameters = new List<string>()
            },
            new()
            {
                Name = "get_workbook_metadata",
                Description = "获取工作簿的元数据信息",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string>()
            },
            new()
            {
                Name = "batch_call",
                Description = "在单次MCP请求中执行多个工具调用，减少网络往返开销",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "calls", new { type = "array", description = "工具调用数组，每项包含toolName和arguments", items = new { type = "object", properties = new { toolName = new { type = "string" }, arguments = new { type = "object" } } } } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "calls" }
            },
            new()
            {
                Name = "create_workbook_with_sheets",
                Description = "一次性创建工作簿及多个工作表",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名" } },
                            { "sheetNames", new { type = "array", description = "工作表名称数组（JSON数组，默认[\"Sheet1\"]）", items = new { type = "string" } } },
                            { "sheetNamePrefix", new { type = "string", description = "工作表名称前缀（设置后按前缀+两位编号创建）" } },
                            { "sheetCount", new { type = "integer", description = "使用前缀时创建的工作表数量" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "fileName" }
            },
            new()
            {
                Name = "write_table_data",
                Description = "一次性写入完整表格（表头+数据）",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名" } },
                            { "sheetName", new { type = "string", description = "工作表名称" } },
                            { "headers", new { type = "array", description = "表头（JSON数组）", items = new { type = "string" } } },
                            { "data", new { type = "array", description = "数据（JSON二维数组）", items = new { type = "array" } } },
                            { "startRow", new { type = "integer", description = "起始行（默认1）" } },
                            { "startColumn", new { type = "integer", description = "起始列（默认1）" } },
                            { "headerStyle", new { type = "string", description = "表头样式：bold/none（默认bold）" } },
                            { "autoFit", new { type = "boolean", description = "是否自动调整列宽（默认true）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "fileName", "sheetName", "headers", "data" }
            },
            new()
            {
                Name = "read_table_data",
                Description = "一次性读取完整表格（表头+数据）",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "fileName", new { type = "string", description = "工作簿文件名" } },
                            { "sheetName", new { type = "string", description = "工作表名称" } },
                            { "range", new { type = "string", description = "读取范围（如A1:D10，可选）" } },
                            { "includeHeaders", new { type = "boolean", description = "是否包含表头（默认true）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "fileName", "sheetName" }
            }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"]?.ToString() : null;

            return toolName switch
            {
                "get_worksheet_names" => SkillResult.Ok(string.Join(", ", _provider.GetWorksheetNames(fileName!))),
                "get_open_workbooks" => SkillResult.Ok(string.Join(", ", _provider.GetOpenWorkbooks())),
                "get_workbook_metadata" => SkillResult.Ok(_provider.GetWorkbookMetadata(fileName!)),
                "batch_call" => await ExecuteBatchCallAsync(arguments),
                "create_workbook_with_sheets" => ExecuteCreateWorkbookWithSheets(arguments),
                "write_table_data" => ExecuteWriteTableData(arguments),
                "read_table_data" => ExecuteReadTableData(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex)
        {
            return new SkillResult { Success = false, Error = ex.Message };
        }
    }

    private async Task<SkillResult> ExecuteBatchCallAsync(Dictionary<string, object> arguments)
    {
        if (!arguments.TryGetValue("calls", out var callsObj) || callsObj == null)
            return new SkillResult { Success = false, Error = "缺少calls参数" };

        var callsJson = JsonSerializer.Serialize(callsObj);
        var calls = JsonSerializer.Deserialize<List<BatchCallItem>>(callsJson, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        if (calls == null || calls.Count == 0)
            return new SkillResult { Success = false, Error = "calls为空" };

        var sb = new StringBuilder();
        int successCount = 0, failCount = 0;
        for (int i = 0; i < calls.Count; i++)
        {
            var call = calls[i];
            if (string.IsNullOrEmpty(call.ToolName))
            {
                sb.AppendLine($"[{i + 1}] 跳过：toolName为空");
                failCount++;
                continue;
            }
            var result = await _skillManager.ExecuteToolAsync(call.ToolName, call.Arguments ?? new Dictionary<string, object>());
            if (result.Success)
            {
                successCount++;
                sb.AppendLine($"[{i + 1}] {call.ToolName}: {result.Content}");
            }
            else
            {
                failCount++;
                sb.AppendLine($"[{i + 1}] {call.ToolName} 失败: {result.Error}");
            }
        }
        sb.AppendLine($"汇总: 共{calls.Count}个调用, 成功{successCount}, 失败{failCount}");
        return SkillResult.Ok(sb.ToString());
    }

    private SkillResult ExecuteCreateWorkbookWithSheets(Dictionary<string, object> arguments)
    {
        var fileName = arguments["fileName"]?.ToString()!;
        List<string> sheetNames;

        if (arguments.TryGetValue("sheetNamePrefix", out var prefixObj) && prefixObj?.ToString() is string prefix && !string.IsNullOrEmpty(prefix))
        {
            var count = arguments.TryGetValue("sheetCount", out var cntObj) && int.TryParse(cntObj?.ToString(), out var c) ? c : 1;
            sheetNames = Enumerable.Range(1, count).Select(i => $"{prefix}{i:D2}").ToList();
        }
        else if (arguments.TryGetValue("sheetNames", out var namesObj) && namesObj != null)
        {
            var json = JsonSerializer.Serialize(namesObj);
            sheetNames = JsonSerializer.Deserialize<List<string>>(json) ?? new List<string> { "Sheet1" };
        }
        else
        {
            sheetNames = new List<string> { "Sheet1" };
        }

        var path = _provider.CreateWorkbook(fileName, sheetNames[0]);
        for (int i = 1; i < sheetNames.Count; i++)
        {
            _provider.CreateWorksheet(fileName, sheetNames[i]);
        }
        return SkillResult.Ok($"已创建工作簿 {fileName}，包含工作表: {string.Join(", ", sheetNames)}");
    }

    private SkillResult ExecuteWriteTableData(Dictionary<string, object> arguments)
    {
        var fileName = arguments["fileName"]?.ToString()!;
        var sheetName = arguments["sheetName"]?.ToString()!;
        var startRow = arguments.TryGetValue("startRow", out var sr) && int.TryParse(sr?.ToString(), out var sri) ? sri : 1;
        var startCol = arguments.TryGetValue("startColumn", out var sc) && int.TryParse(sc?.ToString(), out var sci) ? sci : 1;
        var headerStyle = arguments.TryGetValue("headerStyle", out var hs) ? hs?.ToString() ?? "bold" : "bold";
        var autoFit = !arguments.TryGetValue("autoFit", out var af) || af?.ToString()?.ToLower() != "false";

        var headersJson = JsonSerializer.Serialize(arguments["headers"]);
        var headers = JsonSerializer.Deserialize<List<string>>(headersJson) ?? new List<string>();

        var dataJson = JsonSerializer.Serialize(arguments["data"]);
        var data = JsonSerializer.Deserialize<List<List<object>>>(dataJson) ?? new List<List<object>>();

        for (int c = 0; c < headers.Count; c++)
        {
            _provider.SetCellValue(fileName, sheetName, startRow, startCol + c, headers[c]);
        }

        if (headerStyle == "bold")
        {
            var headerRange = $"{(char)('A' + startCol - 1)}{startRow}:{(char)('A' + startCol + headers.Count - 2)}{startRow}";
            if (headers.Count > 26)
            {
                var endCol = startCol + headers.Count - 1;
                var colLetter = GetColumnLetter(endCol);
                headerRange = $"{GetColumnLetter(startCol)}{startRow}:{colLetter}{startRow}";
            }
            _provider.SetCellFormat(fileName, sheetName, headerRange, bold: true);
        }

        for (int r = 0; r < data.Count; r++)
        {
            for (int c = 0; c < data[r].Count; c++)
            {
                _provider.SetCellValue(fileName, sheetName, startRow + 1 + r, startCol + c, data[r][c]);
            }
        }

        if (autoFit)
        {
            for (int c = 0; c < headers.Count; c++)
            {
                _provider.SetColumnWidth(fileName, sheetName, startCol + c, 15);
            }
        }

        return SkillResult.Ok($"已写入表格数据到 {fileName}/{sheetName}，{headers.Count}列{data.Count}行数据");
    }

    private SkillResult ExecuteReadTableData(Dictionary<string, object> arguments)
    {
        var fileName = arguments["fileName"]?.ToString()!;
        var sheetName = arguments["sheetName"]?.ToString()!;
        var includeHeaders = !arguments.TryGetValue("includeHeaders", out var ih) || ih?.ToString()?.ToLower() != "false";

        string rangeAddress;
        if (arguments.TryGetValue("range", out var rng) && !string.IsNullOrEmpty(rng?.ToString()))
        {
            rangeAddress = rng.ToString()!;
        }
        else
        {
            var usedRange = _provider.GetUsedRange(fileName, sheetName);
            rangeAddress = string.IsNullOrEmpty(usedRange) ? "A1" : usedRange;
        }

        var values = _provider.GetRangeValues(fileName, sheetName, rangeAddress);
        int rows = values.GetLength(0);
        int cols = values.GetLength(1);

        var sb = new StringBuilder();
        int startRow = includeHeaders ? 0 : 1;
        for (int r = startRow; r < rows; r++)
        {
            var rowData = new string[cols];
            for (int c = 0; c < cols; c++)
            {
                rowData[c] = values[r, c]?.ToString() ?? "";
            }
            sb.AppendLine(string.Join(" | ", rowData));
        }
        return SkillResult.Ok(sb.ToString());
    }

    private static string GetColumnLetter(int column)
    {
        string letter = "";
        while (column > 0)
        {
            int mod = (column - 1) % 26;
            letter = Convert.ToChar('A' + mod) + letter;
            column = (column - 1) / 26;
        }
        return letter;
    }

    private class BatchCallItem
    {
        public string ToolName { get; set; } = "";
        public Dictionary<string, object>? Arguments { get; set; }
    }
}