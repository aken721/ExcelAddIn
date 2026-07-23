using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelApiSkill : ISkill
{
    private readonly IExcelProvider _provider;
    private static readonly HttpClient _http = new() { Timeout = TimeSpan.FromSeconds(30) };
    public ExcelApiSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelApi";
    public string Description => "API接口数据获取：REST API调用、认证";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "fetch_api_data", Description = "调用REST API获取数据并写入Excel",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "url", new { type = "string", description = "API地址（必需）" } },
                            { "method", new { type = "string", description = "请求方法：GET/POST（默认GET）" } },
                            { "headers", new { type = "string", description = "请求头（JSON格式，可选）" } },
                            { "body", new { type = "string", description = "请求体（JSON格式，可选）" } },
                            { "outputFileName", new { type = "string", description = "输出工作簿文件名（可选）" } },
                            { "outputSheetName", new { type = "string", description = "输出工作表名称（默认'API数据'）" } },
                            { "dataPath", new { type = "string", description = "JSON数据路径，如data.items（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "url" } },
            new() { Name = "test_api_connection", Description = "测试API连接是否可用",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "url", new { type = "string", description = "API地址" } },
                            { "method", new { type = "string", description = "请求方法（默认GET）" } },
                            { "headers", new { type = "string", description = "请求头（JSON格式，可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "url" } }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "fetch_api_data" => await FetchApiDataAsync(arguments),
                "test_api_connection" => await TestConnectionAsync(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private async Task<SkillResult> FetchApiDataAsync(Dictionary<string, object> args)
    {
        var url = GetStr(args, "url");
        var method = GetStr(args, "method") ?? "GET";
        var outputFn = GetStr(args, "outputFileName");
        var outputSn = GetStr(args, "outputSheetName") ?? "API数据";
        var dataPath = GetStr(args, "dataPath");

        var request = new HttpRequestMessage(new HttpMethod(method.ToUpper()), url);
        ApplyHeaders(request, GetStr(args, "headers"));

        if (method.ToUpper() == "POST" && args.ContainsKey("body"))
            request.Content = new StringContent(GetStr(args, "body") ?? "", Encoding.UTF8, "application/json");

        var response = await _http.SendAsync(request);
        var json = await response.Content.ReadAsStringAsync();

        if (!response.IsSuccessStatusCode)
            return new SkillResult { Success = false, Error = $"API请求失败: {response.StatusCode} - {json}" };

        var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;

        if (!string.IsNullOrEmpty(dataPath))
            foreach (var seg in dataPath.Split('.'))
                if (root.TryGetProperty(seg, out var prop)) root = prop;

        var fn = outputFn ?? "api_data.xlsx";
        if (root.ValueKind == JsonValueKind.Array)
        {
            var items = root.EnumerateArray().ToList();
            if (items.Count == 0) return SkillResult.Ok("API返回空数组");
            _provider.CreateWorkbook(fn, outputSn);
            var props = items[0].EnumerateObject().Select(p => p.Name).ToList();
            for (int c = 0; c < props.Count; c++) _provider.SetCellValue(fn, outputSn, 1, c + 1, props[c]);
            for (int r = 0; r < items.Count; r++)
            {
                var obj = items[r];
                for (int c = 0; c < props.Count; c++)
                {
                    if (obj.TryGetProperty(props[c], out var val))
                        _provider.SetCellValue(fn, outputSn, r + 2, c + 1, val.ValueKind == JsonValueKind.Number ? val.GetDouble().ToString() : val.ToString());
                }
            }
            _provider.SaveWorkbook(fn);
            return SkillResult.Ok($"API数据已写入，共 {items.Count} 条记录到 '{outputSn}'");
        }

        return SkillResult.Ok($"API响应:\n{json[..Math.Min(json.Length, 2000)]}");
    }

    private async Task<SkillResult> TestConnectionAsync(Dictionary<string, object> args)
    {
        var url = GetStr(args, "url");
        var method = GetStr(args, "method") ?? "GET";
        try
        {
            var request = new HttpRequestMessage(new HttpMethod(method.ToUpper()), url);
            ApplyHeaders(request, GetStr(args, "headers"));
            var response = await _http.SendAsync(request);
            return SkillResult.Ok($"API连接{'{'}测试{'}'}: {(response.IsSuccessStatusCode ? "成功" : "失败")} (HTTP {(int)response.StatusCode})");
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = $"API连接失败: {ex.Message}" }; }
    }

    private void ApplyHeaders(HttpRequestMessage request, string headersJson)
    {
        if (string.IsNullOrEmpty(headersJson)) return;
        var headers = JsonSerializer.Deserialize<Dictionary<string, string>>(headersJson);
        if (headers == null) return;
        foreach (var h in headers) request.Headers.TryAddWithoutValidation(h.Key, h.Value);
    }

    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
}