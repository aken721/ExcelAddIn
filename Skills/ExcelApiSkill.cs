using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelApiSkill : ISkill
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        public string Name => "ExcelApi";
        public string Description => "API接口数据提取技能，支持从REST API获取数据并写入Excel";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "fetch_api_data",
                    Description = "从API接口获取数据并写入Excel。支持GET/POST请求，支持API Key认证。当用户要求调用API、获取API数据时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "url", new { type = "string", description = "API接口URL" } },
                                { "method", new { type = "string", description = "请求方法：GET/POST，默认GET" } },
                                { "headers", new { type = "string", description = "请求头（JSON格式，可选）" } },
                                { "body", new { type = "string", description = "请求体（POST时使用，可选）" } },
                                { "apiKey", new { type = "string", description = "API密钥（可选）" } },
                                { "authType", new { type = "string", description = "认证类型：none/header/query/bearer，默认none" } },
                                { "outputSheetName", new { type = "string", description = "输出工作表名称（可选，默认'API_Data'）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "url" }
                },
                new SkillTool
                {
                    Name = "test_api_connection",
                    Description = "测试API连接是否正常。当用户要求测试API时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "url", new { type = "string", description = "API接口URL" } },
                                { "method", new { type = "string", description = "请求方法：GET/POST，默认GET" } },
                                { "apiKey", new { type = "string", description = "API密钥（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "url" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "fetch_api_data":
                        return await FetchApiDataAsync(arguments);
                    case "test_api_connection":
                        return await TestApiConnectionAsync(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelApiSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private async Task<SkillResult> FetchApiDataAsync(Dictionary<string, object> arguments)
        {
            var url = arguments["url"].ToString();
            var method = arguments.ContainsKey("method") ? arguments["method"].ToString().ToUpper() : "GET";
            var outputSheetName = arguments.ContainsKey("outputSheetName") 
                ? arguments["outputSheetName"].ToString() 
                : "API_Data";

            try
            {
                using (var request = new HttpRequestMessage())
                {
                    request.RequestUri = new Uri(url);
                    request.Method = method == "POST" ? HttpMethod.Post : HttpMethod.Get;

                    if (arguments.ContainsKey("headers"))
                    {
                        var headers = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(arguments["headers"].ToString());
                        foreach (var header in headers)
                        {
                            request.Headers.TryAddWithoutValidation(header.Key, header.Value);
                        }
                    }

                    if (arguments.ContainsKey("apiKey") && !string.IsNullOrEmpty(arguments["apiKey"]?.ToString()))
                    {
                        var apiKey = arguments["apiKey"].ToString();
                        var authType = arguments.ContainsKey("authType") ? arguments["authType"].ToString().ToLower() : "header";

                        switch (authType)
                        {
                            case "header":
                                request.Headers.Add("X-API-KEY", apiKey);
                                break;
                            case "bearer":
                                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                                break;
                            case "query":
                                var separator = url.Contains("?") ? "&" : "?";
                                request.RequestUri = new Uri($"{url}{separator}api_key={Uri.EscapeDataString(apiKey)}");
                                break;
                        }
                    }

                    if (method == "POST" && arguments.ContainsKey("body"))
                    {
                        var body = arguments["body"].ToString();
                        request.Content = new StringContent(body, Encoding.UTF8, "application/json");
                    }

                    var response = await _httpClient.SendAsync(request);
                    var content = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        return new SkillResult { Success = false, Error = $"API返回错误: {response.StatusCode} - {content}" };
                    }

                    var dataTable = ParseApiResponse(content);
                    WriteDataTableToExcel(dataTable, outputSheetName);

                    return new SkillResult 
                    { 
                        Success = true, 
                        Content = $"API数据获取成功，共 {dataTable.Rows.Count} 行数据已写入工作表 '{outputSheetName}'" 
                    };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = $"API请求失败：{ex.Message}" };
            }
        }

        private async Task<SkillResult> TestApiConnectionAsync(Dictionary<string, object> arguments)
        {
            var url = arguments["url"].ToString();
            var method = arguments.ContainsKey("method") ? arguments["method"].ToString().ToUpper() : "GET";

            try
            {
                using (var request = new HttpRequestMessage())
                {
                    request.RequestUri = new Uri(url);
                    request.Method = method == "POST" ? HttpMethod.Post : HttpMethod.Get;

                    if (arguments.ContainsKey("apiKey") && !string.IsNullOrEmpty(arguments["apiKey"]?.ToString()))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", arguments["apiKey"].ToString());
                    }

                    var response = await _httpClient.SendAsync(request);
                    var content = await response.Content.ReadAsStringAsync();

                    return new SkillResult 
                    { 
                        Success = response.IsSuccessStatusCode, 
                        Content = response.IsSuccessStatusCode 
                            ? $"API连接成功！状态码: {response.StatusCode}" 
                            : $"API连接失败！状态码: {response.StatusCode}\n响应: {content}" 
                    };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = $"API连接测试失败：{ex.Message}" };
            }
        }

        private System.Data.DataTable ParseApiResponse(string content)
        {
            var dataTable = new System.Data.DataTable();

            content = content.Trim();

            if (content.StartsWith("["))
            {
                var jsonArray = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JArray>(content);
                if (jsonArray.Count > 0)
                {
                    foreach (var prop in jsonArray[0].Children<Newtonsoft.Json.Linq.JProperty>())
                    {
                        dataTable.Columns.Add(prop.Name);
                    }

                    foreach (var item in jsonArray)
                    {
                        var row = dataTable.NewRow();
                        foreach (var prop in item.Children<Newtonsoft.Json.Linq.JProperty>())
                        {
                            row[prop.Name] = prop.Value?.ToString() ?? "";
                        }
                        dataTable.Rows.Add(row);
                    }
                }
            }
            else if (content.StartsWith("{"))
            {
                var jsonObj = Newtonsoft.Json.JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(content);
                
                var dataPath = FindDataArray(jsonObj);
                if (dataPath != null)
                {
                    var dataArray = dataPath as Newtonsoft.Json.Linq.JArray;
                    if (dataArray != null && dataArray.Count > 0)
                    {
                        foreach (var prop in dataArray[0].Children<Newtonsoft.Json.Linq.JProperty>())
                        {
                            dataTable.Columns.Add(prop.Name);
                        }

                        foreach (var item in dataArray)
                        {
                            var row = dataTable.NewRow();
                            foreach (var prop in item.Children<Newtonsoft.Json.Linq.JProperty>())
                            {
                                row[prop.Name] = prop.Value?.ToString() ?? "";
                            }
                            dataTable.Rows.Add(row);
                        }
                    }
                }
                else
                {
                    foreach (var prop in jsonObj.Properties())
                    {
                        dataTable.Columns.Add(prop.Name);
                    }
                    var row = dataTable.NewRow();
                    foreach (var prop in jsonObj.Properties())
                    {
                        row[prop.Name] = prop.Value?.ToString() ?? "";
                    }
                    dataTable.Rows.Add(row);
                }
            }

            return dataTable;
        }

        private Newtonsoft.Json.Linq.JToken FindDataArray(Newtonsoft.Json.Linq.JObject obj)
        {
            var arrayProps = new[] { "data", "results", "items", "records", "list", "rows" };
            
            foreach (var prop in arrayProps)
            {
                if (obj[prop] is Newtonsoft.Json.Linq.JArray array)
                    return array;
            }

            foreach (var prop in obj.Properties())
            {
                if (prop.Value is Newtonsoft.Json.Linq.JArray array)
                    return array;
                if (prop.Value is Newtonsoft.Json.Linq.JObject nested)
                {
                    var found = FindDataArray(nested);
                    if (found != null) return found;
                }
            }

            return null;
        }

        private void WriteDataTableToExcel(System.Data.DataTable dataTable, string sheetName)
        {
            var workbook = ThisAddIn.app.ActiveWorkbook;

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            Excel.Worksheet sheet;
            
            var existingNames = new List<string>();
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            
            string actualSheetName = sheetName;
            if (existingNames.Any(n => string.Equals(n, sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                int suffix = 2;
                while (existingNames.Any(n => string.Equals(n, $"{sheetName}_{suffix}", StringComparison.OrdinalIgnoreCase)))
                {
                    suffix++;
                }
                actualSheetName = $"{sheetName}_{suffix}";
            }
            
            sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            sheet.Name = actualSheetName;

            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                sheet.Cells[1, c + 1].Value = dataTable.Columns[c].ColumnName;
                sheet.Cells[1, c + 1].Font.Bold = true;
            }

            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                for (int c = 0; c < dataTable.Columns.Count; c++)
                {
                    sheet.Cells[r + 2, c + 1].Value = dataTable.Rows[r][c];
                }
            }

            sheet.UsedRange.Columns.AutoFit();
            sheet.Activate();

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;
        }
    }
}
