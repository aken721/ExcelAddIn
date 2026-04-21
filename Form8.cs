using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace TableMagic
{
    public partial class Form8 : Form
    {
        private bool _isCloudConnection = true; // 标识连接类型（true=云端，false=本地）
        private CancellationTokenSource _modelFetchCts = null; // 用于取消模型获取请求
        private System.Windows.Forms.Timer _debounceTimer = null; // 防抖定时器

        public Form8()
        {
            InitializeComponent();
            pbxPassword.BackgroundImage = Properties.Resources.eye_hide;
            txbKey.UseSystemPasswordChar = true;
            
            // 初始化 cbxLLM 下拉框选项
            InitializeLLMOptions();
            
            // 初始化防抖定时器
            _debounceTimer = new System.Windows.Forms.Timer();
            _debounceTimer.Interval = 1000; // 1秒防抖
            _debounceTimer.Tick += DebounceTimer_Tick;
            
            // 添加事件处理
            cbxLLM.SelectedIndexChanged += CbxLLM_SelectedIndexChanged;
        }

        // 初始化 LLM 提供商选项
        private void InitializeLLMOptions()
        {
            cbxLLM.Items.Clear();
            cbxLLM.Items.AddRange(new string[]
            {
                "DeepSeek",
                "Qwen", 
                "Kimi", 
                "GLM", 
                "MinMax", 
                "OpenRoute", 
                "Claude Code", 
                "Gemini", 
                "OpenAI", 
                "Ollama", 
                "vLLM", 
                "LM Studio", 
                "llama.cpp", 
                "自定义"
            });
            
            // 默认选择 DeepSeek
            cbxLLM.SelectedIndex = 0;
        }

        // LLM 提供商选择变化事件
        private void CbxLLM_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxLLM == null) return;

            string provider = cbxLLM.Text;
            string defaultUrl = GetDefaultApiUrl(provider);
            txbUrl.Text = defaultUrl;
            
            // 自定义提供商允许编辑URL
            txbUrl.ReadOnly = provider != "自定义";
            
            // 清除模型列表，准备获取新的模型
            cbxModel.Items.Clear();
            cbxModel.Text = "";
            
            // 无论是本地部署还是云端，都在选择服务商后请求一次可选模型列表
            if (provider != "自定义")
            {
                // 在UI线程上获取所需值，然后传递给后台任务
                string apiKey = txbKey.Text ?? "";
                string apiUrl = txbUrl.Text;
                Task.Run(() => FetchModelsAsync(provider, apiKey, apiUrl));
            }
        }

        // 根据模型提供商获取默认API地址
        private string GetDefaultApiUrl(string provider)
        {
            switch (provider)
            {
                case "DeepSeek":
                    return "https://api.deepseek.com/v1/chat/completions";
                case "Qwen":
                    return "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions";
                case "Kimi":
                    return "https://api.moonshot.cn/v1/chat/completions";
                case "GLM":
                    return "https://open.bigmodel.cn/api/paas/v4/chat/completions";
                case "MinMax":
                    return "https://api.minimax.chat/v1/text/chatcompletion_pro";
                case "OpenRoute":
                    return "https://openrouter.ai/api/v1/chat/completions";
                case "Claude Code":
                    return "https://api.anthropic.com/v1/messages";
                case "Gemini":
                    return "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent";
                case "OpenAI":
                    return "https://api.openai.com/v1/chat/completions";
                case "Ollama":
                    return "http://localhost:11434/v1/chat/completions";
                case "vLLM":
                    return "http://localhost:8000/v1/chat/completions";
                case "LM Studio":
                    return "http://localhost:1234/v1/chat/completions";
                case "llama.cpp":
                    return "http://localhost:8080/v1/chat/completions";
                case "自定义":
                    return "";
                default:
                    return "";
            }
        }

        private void Form8_Load(object sender, EventArgs e)
        {
            DecodeConfig();
            
            // 添加API Key输入变化事件，用于自动获取模型列表
            txbKey.TextChanged += TxbKey_TextChanged;
        }

        // API Key输入变化事件
        private void TxbKey_TextChanged(object sender, EventArgs e)
        {
            // 重启防抖定时器
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        // 防抖定时器触发事件
        private void DebounceTimer_Tick(object sender, EventArgs e)
        {
            _debounceTimer.Stop();
            // 在UI线程上获取所需值，然后传递给后台任务
            string provider = cbxLLM.Text;
            string apiKey = txbKey.Text;
            string apiUrl = txbUrl.Text;
            Task.Run(() => FetchModelsAsync(provider, apiKey, apiUrl));
        }

        // 获取模型列表
        private async Task FetchModelsAsync(string provider, string apiKey, string apiUrl)
        {
            // 本地部署的模型提供商不需要API Key
            string[] localProviders = { "Ollama", "vLLM", "LM Studio", "llama.cpp" };
            bool isLocalProvider = localProviders.Contains(provider);
            
            // 自定义提供商或云端模型没有API Key时，不获取模型列表
            if (provider == "自定义" || (!isLocalProvider && string.IsNullOrEmpty(apiKey)) || string.IsNullOrEmpty(apiUrl))
                return;

            // 取消之前的请求
            if (_modelFetchCts != null)
            {
                _modelFetchCts.Cancel();
                _modelFetchCts.Dispose();
            }

            _modelFetchCts = new CancellationTokenSource();
            CancellationToken token = _modelFetchCts.Token;

            try
            {
                // 在UI线程上更新状态
                this.Invoke(new Action(() =>
                {
                    lblResult.Text = "正在获取模型列表...";
                    cbxModel.Items.Clear();
                    cbxModel.Text = "";
                }));

                var (models, errorType, errorMessage) = await GetModelsAsync(provider, apiKey, apiUrl, token);

                if (token.IsCancellationRequested)
                    return;

                // 在UI线程上更新控件
                this.Invoke(new Action(() =>
                {
                    if (models.Count > 0)
                    {
                        cbxModel.Items.AddRange(models.ToArray());
                        cbxModel.SelectedIndex = 0;
                        lblResult.Text = $"成功获取 {models.Count} 个模型";
                    }
                    else if (!string.IsNullOrEmpty(errorType))
                    {
                        // 根据错误类型显示不同的提示信息
                        switch (errorType)
                        {
                            case "api_key_error":
                                lblResult.Text = "需要填写正确的API Key后才能选择模型";
                                break;
                            case "network_error":
                            case "timeout":
                                lblResult.Text = "API地址无法访问，请检查API地址是否正确";
                                break;
                            case "endpoint_not_found":
                                lblResult.Text = "API地址不正确，请检查API地址";
                                break;
                            default:
                                lblResult.Text = errorMessage ?? "未获取到模型列表";
                                break;
                        }
                    }
                    else
                    {
                        lblResult.Text = "未获取到模型列表";
                    }
                }));
            }
            catch (Exception ex)
            {
                if (!token.IsCancellationRequested)
                {
                    this.Invoke(new Action(() =>
                    {
                        lblResult.Text = $"获取模型列表失败: {ex.Message}";
                    }));
                }
            }
            finally
            {
                if (_modelFetchCts != null)
                {
                    _modelFetchCts.Dispose();
                    _modelFetchCts = null;
                }
            }
        }

        // 根据提供商获取模型列表
        private async Task<(List<string> models, string errorType, string errorMessage)> GetModelsAsync(string provider, string apiKey, string apiUrl, CancellationToken token)
        {
            List<string> models = new List<string>();
            string errorType = null;
            string errorMessage = null;

            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(15);

                    // 添加认证头（云端API需要）
                    if (provider != "Ollama" && provider != "vLLM" && provider != "LM Studio" && provider != "llama.cpp")
                    {
                        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                    }

                    string modelsEndpoint = GetModelsEndpoint(provider, apiUrl);
                    if (string.IsNullOrEmpty(modelsEndpoint))
                    {
                        System.Diagnostics.Debug.WriteLine($"[模型获取] 未找到模型端点，provider: {provider}");
                        errorType = "endpoint_not_found";
                        errorMessage = "未找到模型列表API端点";
                        return (models, errorType, errorMessage);
                    }

                    System.Diagnostics.Debug.WriteLine($"[模型获取] 请求端点: {modelsEndpoint}");
                    
                    var response = await client.GetAsync(modelsEndpoint, token);
                    System.Diagnostics.Debug.WriteLine($"[模型获取] 响应状态码: {response.StatusCode}");
                    
                    if (!response.IsSuccessStatusCode)
                    {
                        string errorContent = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"[模型获取] 请求失败: {response.StatusCode}, 错误: {errorContent}");
                        
                        // 根据状态码和错误内容判断错误类型
                        if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized || 
                            errorContent.Contains("api_key") || errorContent.Contains("apiKey") ||
                            errorContent.Contains("invalid") || errorContent.Contains("unauthorized") ||
                            errorContent.Contains("authentication") || errorContent.Contains("Authorization"))
                        {
                            errorType = "api_key_error";
                            errorMessage = "API Key 无效或未填写，请检查API Key是否正确";
                        }
                        else if (response.StatusCode == System.Net.HttpStatusCode.NotFound)
                        {
                            errorType = "endpoint_not_found";
                            errorMessage = "API地址不正确，请检查API地址";
                        }
                        else
                        {
                            errorType = "api_error";
                            errorMessage = $"API错误: {response.StatusCode}";
                        }
                        
                        return (models, errorType, errorMessage);
                    }

                    string responseContent = await response.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine($"[模型获取] 响应内容: {responseContent.Substring(0, Math.Min(500, responseContent.Length))}...");
                    
                    models = ParseModelsResponse(provider, responseContent);
                    System.Diagnostics.Debug.WriteLine($"[模型获取] 解析到 {models.Count} 个模型");
                }
            }
            catch (HttpRequestException ex)
            {
                System.Diagnostics.Debug.WriteLine($"[模型获取] 网络异常: {ex.Message}");
                errorType = "network_error";
                errorMessage = "无法连接到API地址，请检查网络连接和API地址是否正确";
            }
            catch (TaskCanceledException)
            {
                System.Diagnostics.Debug.WriteLine($"[模型获取] 请求超时");
                errorType = "timeout";
                errorMessage = "请求超时，请检查API地址是否正确或网络是否畅通";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[模型获取] 异常: {ex.GetType().Name} - {ex.Message}");
                if (ex.InnerException != null)
                {
                    System.Diagnostics.Debug.WriteLine($"[模型获取] 内部异常: {ex.InnerException.Message}");
                }
                errorType = "unknown_error";
                errorMessage = $"未知错误: {ex.Message}";
            }

            return (models, errorType, errorMessage);
        }

        // 获取模型列表的API端点
        private string GetModelsEndpoint(string provider, string apiUrl)
        {
            switch (provider)
            {
                case "DeepSeek":
                    return "https://api.deepseek.com/v1/models";
                case "Qwen":
                    return "https://dashscope.aliyuncs.com/compatible-mode/v1/models";
                case "Kimi":
                    return "https://api.moonshot.cn/v1/models";
                case "OpenAI":
                    return "https://api.openai.com/v1/models";
                case "OpenRoute":
                    return "https://openrouter.ai/api/v1/models";
                case "GLM":
                    return "https://open.bigmodel.cn/api/paas/v4/models";
                case "MinMax":
                    return "https://api.minimax.chat/v1/models";
                case "Claude Code":
                    return "https://api.anthropic.com/v1/models";
                case "Gemini":
                    return "https://generativelanguage.googleapis.com/v1beta/models";
                case "Ollama":
                    return "http://localhost:11434/api/tags";
                case "vLLM":
                case "LM Studio":
                case "llama.cpp":
                    return apiUrl.Replace("/chat/completions", "/models");
                default:
                    return "";
            }
        }

        // 解析模型列表响应（使用动态解析，更加健壮）
        private List<string> ParseModelsResponse(string provider, string responseContent)
        {
            List<string> models = new List<string>();

            try
            {
                using (var jsonDoc = System.Text.Json.JsonDocument.Parse(responseContent))
                {
                    var root = jsonDoc.RootElement;

                    switch (provider)
                    {
                        case "DeepSeek":
                        case "Qwen":
                        case "Kimi":
                        case "OpenAI":
                        case "OpenRoute":
                        case "GLM":
                        case "vLLM":
                        case "LM Studio":
                        case "llama.cpp":
                            // 标准OpenAI格式: {"data": [{"id": "model-name", ...}, ...]}
                            if (root.TryGetProperty("data", out var dataArray) && dataArray.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var item in dataArray.EnumerateArray())
                                {
                                    if (item.TryGetProperty("id", out var idProp))
                                    {
                                        models.Add(idProp.GetString());
                                    }
                                }
                            }
                            break;

                        case "MinMax":
                            // MinMax格式: {"models": [{"model": "model-name", ...}, ...]}
                            if (root.TryGetProperty("models", out var minMaxModels) && minMaxModels.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var item in minMaxModels.EnumerateArray())
                                {
                                    // 优先使用 model 字段，其次使用 name 字段
                                    if (item.TryGetProperty("model", out var modelProp))
                                    {
                                        models.Add(modelProp.GetString());
                                    }
                                    else if (item.TryGetProperty("name", out var nameProp))
                                    {
                                        models.Add(nameProp.GetString());
                                    }
                                }
                            }
                            break;

                        case "Ollama":
                            // Ollama格式: {"models": [{"name": "model-name", ...}, ...]}
                            System.Diagnostics.Debug.WriteLine($"[模型获取] Ollama响应内容: {responseContent}");
                            if (root.TryGetProperty("models", out var ollamaModels) && ollamaModels.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var item in ollamaModels.EnumerateArray())
                                {
                                    // 优先使用 name 字段，其次使用 model 字段
                                    if (item.TryGetProperty("name", out var nameProp))
                                    {
                                        models.Add(nameProp.GetString());
                                    }
                                    else if (item.TryGetProperty("model", out var modelProp))
                                    {
                                        models.Add(modelProp.GetString());
                                    }
                                }
                            }
                            System.Diagnostics.Debug.WriteLine($"[模型获取] Ollama解析结果: {models.Count} 个模型");
                            break;

                        default:
                            // 尝试通用解析：查找常见的模型列表字段
                            if (root.TryGetProperty("data", out var genericData) && genericData.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var item in genericData.EnumerateArray())
                                {
                                    if (item.TryGetProperty("id", out var idProp))
                                        models.Add(idProp.GetString());
                                    else if (item.TryGetProperty("name", out var nameProp))
                                        models.Add(nameProp.GetString());
                                    else if (item.TryGetProperty("model", out var modelProp))
                                        models.Add(modelProp.GetString());
                                }
                            }
                            else if (root.TryGetProperty("models", out var genericModels) && genericModels.ValueKind == JsonValueKind.Array)
                            {
                                foreach (var item in genericModels.EnumerateArray())
                                {
                                    if (item.TryGetProperty("name", out var nameProp))
                                        models.Add(nameProp.GetString());
                                    else if (item.TryGetProperty("model", out var modelProp))
                                        models.Add(modelProp.GetString());
                                    else if (item.TryGetProperty("id", out var idProp))
                                        models.Add(idProp.GetString());
                                }
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[模型获取] 解析异常: {ex.Message}");
                // 解析失败，返回空列表
            }

            return models;
        }

        // 模型响应相关类
        public class OpenAIModelsResponse
        {
            public List<OpenAIModel> data { get; set; }
        }

        public class OpenAIModel
        {
            public string id { get; set; }
            public string type { get; set; }
            public int created { get; set; }
            public string owned_by { get; set; }
        }

        public class MinMaxModelsResponse
        {
            public List<MinMaxModel> models { get; set; }
        }

        public class MinMaxModel
        {
            public string model { get; set; }
            public string name { get; set; }
            public string description { get; set; }
        }

        public class OllamaModelsResponse
        {
            public List<OllamaModel> models { get; set; }
        }

        public class OllamaModel
        {
            public string name { get; set; }
            public string model { get; set; }
            public OllamaModelDetails details { get; set; }
        }

        public class OllamaModelDetails
        {
            public string parent_model { get; set; }
            public string format { get; set; }
            public string family { get; set; }
            public List<string> families { get; set; }
            public string parameter_size { get; set; }
            public string quantization_level { get; set; }
        }
        //重绘选项页布局
        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            //调整选项卡文字方向
            SolidBrush _Brush = new SolidBrush(Color.Black);//单色画刷
            RectangleF _TabTextArea = (RectangleF)tabControl1.GetTabRect(e.Index);//绘制区域
            StringFormat _sf = new StringFormat();//封装文本布局格式信息
            _sf.LineAlignment = StringAlignment.Center;
            _sf.Alignment = StringAlignment.Center;
            e.Graphics.DrawString(tabControl1.Controls[e.Index].Text, SystemInformation.MenuFont, _Brush, _TabTextArea, _sf);
        }

        private async Task<string> GetAIResponse(string apiKey, string model)
        {
            string apiUrl = txbUrl.Text;

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", apiKey);

                // 构造请求体
                var requestBody = new
                {
                    model = model, // 选择的模型
                    messages = new[]
                    {
                        new { role = "user", content = "Hello!" }
                    }
                };

                try
                {
                    var response = await client.PostAsJsonAsync(apiUrl, requestBody);
                    var responseContent = await response.Content.ReadAsStringAsync();

                    if (response.IsSuccessStatusCode)
                    {
                        return "连接成功";
                    }
                    else
                    {
                        // 返回错误信息和错误代码
                        return $"测试失败，错误代码：{response.StatusCode}，错误信息：{responseContent}";
                    }
                }
                catch (HttpRequestException ex)
                {
                    // 返回异常信息
                    return $"测试失败，发生异常：{ex.Message}";
                }
            }
        }


        private async void btnTest_Click(object sender, EventArgs e)
        {
            string provider = cbxLLM.Text;
            
            lblResult.Text = "开始检测......";
            txbKey.ReadOnly = true;
            cbxModel.Enabled = false;
            cbxEnterKey.Enabled = false;
            btnTest.Enabled = false;
            btnSave.Enabled = false;
            btnQuit.Enabled = false;
            this.Refresh();

            string result;

            if (!_isCloudConnection || provider == "Ollama" || provider == "vLLM" || provider == "LM Studio" || provider == "llama.cpp")
            {
                // 本地API连接测试
                string apiUrl = txbUrl.Text;
                string modelsUrl = apiUrl;

                if (apiUrl.EndsWith("/v1/chat/completions"))
                {
                    modelsUrl = apiUrl.Replace("/v1/chat/completions", "/v1/models");
                }
                else if (apiUrl.EndsWith("/chat/completions"))
                {
                    modelsUrl = apiUrl.Replace("/chat/completions", "/models");
                }
                else if (provider == "Ollama")
                {
                    modelsUrl = "http://localhost:11434/api/tags";
                }

                try
                {
                    using (var client = new HttpClient())
                    {
                        client.Timeout = TimeSpan.FromSeconds(10);
                        var response = await client.GetAsync(modelsUrl);
                        if (response.IsSuccessStatusCode)
                        {
                            result = "本地连接成功";
                        }
                        else
                        {
                            result = $"本地连接测试失败，状态码: {response.StatusCode}";
                        }
                    }
                }
                catch (Exception ex)
                {
                    result = $"本地连接测试异常: {ex.Message}";
                }
            }
            else
            {
                // 云端API连接测试
                string model = cbxModel.Text;

                if (string.IsNullOrEmpty(model))
                {
                    lblResult.Text = "请选择模型";
                    txbKey.ReadOnly = false;
                    btnTest.Enabled = true;
                    cbxModel.Enabled = true;
                    cbxEnterKey.Enabled = true;
                    btnSave.Enabled = true;
                    btnQuit.Enabled = true;
                    return;
                }
                result = await TestCloudConnection(txbKey.Text, model, provider);
            }

            if (!string.IsNullOrEmpty(result))
            {
                lblResult.Text = result;
            }

            txbKey.ReadOnly = false;
            btnTest.Enabled = true;
            cbxModel.Enabled = true;
            cbxEnterKey.Enabled = true;
            btnSave.Enabled = true;
            btnQuit.Enabled = true;
        }

        // 测试云端连接
        private async Task<string> TestCloudConnection(string apiKey, string model, string provider)
        {
            string apiUrl = txbUrl.Text;

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                // 根据提供商构造请求体
                var requestBody = GetRequestBodyForProvider(provider, model);

                try
                {
                    var response = await client.PostAsJsonAsync(apiUrl, requestBody);
                    var responseContent = await response.Content.ReadAsStringAsync();

                    if (response.IsSuccessStatusCode)
                    {
                        return "连接成功";
                    }
                    else
                    {
                        // 返回错误信息和错误代码
                        return $"测试失败，错误代码：{response.StatusCode}，错误信息：{responseContent}";
                    }
                }
                catch (HttpRequestException ex)
                {
                    // 返回异常信息
                    return $"测试失败，发生异常：{ex.Message}";
                }
            }
        }

        // 根据提供商获取请求体
        private object GetRequestBodyForProvider(string provider, string model)
        {
            switch (provider)
            {
                case "Claude Code":
                    // Anthropic Claude格式
                    return new
                    {
                        model = model,
                        messages = new[]
                        {
                            new { role = "user", content = "Hello, Claude!" }
                        },
                        max_tokens = 100
                    };
                case "Gemini":
                    // Google Gemini格式
                    return new
                    {
                        contents = new[]
                        {
                            new {
                                parts = new[]
                                {
                                    new { text = "Hello, Gemini!" }
                                }
                            }
                        }
                    };
                case "DeepSeek":
                default:
                    // 标准OpenAI格式
                    return new
                    {
                        model = model,
                        messages = new[]
                        {
                            new { role = "user", content = "Hello!" }
                        },
                        max_tokens = 100
                    };
            }
        }
        private const string KeyFilePath = "encryption.key"; // 保存密钥和IV的文件路径
        private const string ConfigFilePath = "config.encrypted"; // 保存加密配置信息的文件路径

        //加密保存配置信息
        private void btnSave_Click(object sender, EventArgs e)
        {
            string provider = cbxLLM.Text;
            
            string apiKey = txbKey.Text.Trim();
            string model = cbxModel.Text;
            string apiUrl = txbUrl.Text.Trim();
            string enterMode = "0";
            
            switch (cbxEnterKey.SelectedIndex)
            {
                case 0:
                    enterMode = "0";
                    break;
                case 1:
                    enterMode = "1";
                    break;
                case 2:
                    enterMode = "2";
                    break;
            }

            // 确定连接类型
            string connectionType = _isCloudConnection ? "cloud" : "local";
            string isCloudModel = _isCloudConnection ? "true" : "false";

            if (string.IsNullOrEmpty(apiKey) && _isCloudConnection && provider != "Ollama" && provider != "vLLM" && provider != "LM Studio" && provider != "llama.cpp")
            {
                lblResult.Text = "API KEY不能为空";
                return;
            }
            if (string.IsNullOrEmpty(apiUrl))
            {
                lblResult.Text = "API地址不能为空";
                return;
            }
            if (string.IsNullOrEmpty(model))
            {
                lblResult.Text = "请选择模型";
                return;
            }

            string content = $"api-key^{apiKey};model^{model};api-url^{apiUrl};enter-mode^{enterMode};connection-type^{connectionType};is-cloud-model^{isCloudModel};timeout-minutes^{txbWaitingTime.Text.Trim()};provider^{provider}";

            try
            {
                // 获取或生成密钥和IV
                (byte[] key, byte[] iv) = GetOrCreateEncryptionKey();

                // 加密文本内容
                string encryptedContent = EncryptString(content, key, iv);

                // 将加密内容写入文件
                File.WriteAllText(ConfigFilePath, encryptedContent);

                lblResult.Text = "保存成功";
                this.Refresh();
                Thread.Sleep(1000);
                this.Dispose();
            }
            catch (Exception ex)
            {
                lblResult.Text = $"发生错误：{ex.Message}";
            }
        }

        //读取配置信息
        private void DecodeConfig()
        {
            if (!File.Exists(ConfigFilePath))
            {
                lblResult.Text = "未初始化配置";
                txbKey.Text = "";
                cbxEnterKey.SelectedIndex = 0;
                txbWaitingTime.Text = "5";
                return;
            }

            try
            {
                // 获取密钥和IV
                (byte[] key, byte[] iv) = GetEncryptionKey();

                // 读取加密内容
                string encryptedContent = File.ReadAllText(ConfigFilePath);

                // 解密文本内容
                string decryptedContent = DecryptString(encryptedContent, key, iv);

                // 解析配置信息
                var parts = decryptedContent.Split(';');
                string apiKey = parts[0].Split('^')[1];
                string model = parts[1].Split('^')[1];
                string apiUrl = parts[2].Split('^')[1];
                string enterMode = parts[3].Split('^')[1];
                
                // 读取连接类型
                string connectionType = "cloud"; // 默认为云端
                if (parts.Length >= 5)
                {
                    connectionType = parts[4].Split('^')[1];
                }

                // 读取模型提供商
                string provider = "Qwen"; // 默认为Qwen
                if (parts.Length >= 8)
                {
                    provider = parts[7].Split('^')[1];
                }

                // 设置连接类型标志
                _isCloudConnection = (connectionType == "cloud");

                // 设置UI控件值
                txbKey.Text = apiKey;
                txbUrl.Text = apiUrl;
                
                // 设置模型提供商
                int providerIndex = cbxLLM.FindStringExact(provider);
                if (providerIndex >= 0)
                {
                    cbxLLM.SelectedIndex = providerIndex;
                }
                else
                {
                    cbxLLM.SelectedIndex = 0; // 默认选择第一个
                }

                // 设置模型
                cbxModel.Text = model;

                // 设置回车模式
                switch (enterMode)
                {
                    case "0":
                        cbxEnterKey.SelectedIndex = 0;
                        break;
                    case "1":
                        cbxEnterKey.SelectedIndex = 1;
                        break;
                    case "2":
                        cbxEnterKey.SelectedIndex = 2;
                        break;
                }

                // 设置等待时间
                if (parts.Length >= 7)
                {
                    string timeoutMinutes = parts[6].Split('^')[1];
                    if (!string.IsNullOrEmpty(timeoutMinutes))
                    {
                        txbWaitingTime.Text = timeoutMinutes;
                    }
                    else
                    {
                        txbWaitingTime.Text = "5";
                    }
                }
                else
                {
                    txbWaitingTime.Text = "5";
                }

                // 如果有API Key，尝试获取模型列表
                if (!string.IsNullOrEmpty(apiKey))
                {
                    // 使用已定义的变量，传递给后台任务
                    Task.Run(() => FetchModelsAsync(provider, apiKey, apiUrl));
                }
            }
            catch (Exception ex)
            {
                lblResult.Text = $"发生错误：{ex.Message}";
            }
        }


        private void btnQuit_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        // 获取或生成密钥和IV
        private (byte[], byte[]) GetOrCreateEncryptionKey()
        {
            if (File.Exists(KeyFilePath))
            {
                // 从文件中读取密钥和IV
                string[] lines = File.ReadAllLines(KeyFilePath);
                return (Convert.FromBase64String(lines[0]), Convert.FromBase64String(lines[1]));
            }
            else
            {
                // 自动生成密钥和IV
                using (Aes aesAlg = Aes.Create())
                {
                    aesAlg.GenerateKey();
                    aesAlg.GenerateIV();

                    byte[] key = aesAlg.Key;
                    byte[] iv = aesAlg.IV;

                    // 将密钥和IV保存到文件
                    File.WriteAllLines(KeyFilePath, new[] { Convert.ToBase64String(key), Convert.ToBase64String(iv) });

                    return (key, iv);
                }
            }
        }

        // 从文件中获取密钥和IV
        private (byte[], byte[]) GetEncryptionKey()
        {
            if (!File.Exists(KeyFilePath))
            {
                throw new FileNotFoundException("密钥文件不存在，请先进行加密操作。");
            }

            string[] lines = File.ReadAllLines(KeyFilePath);
            return (Convert.FromBase64String(lines[0]), Convert.FromBase64String(lines[1]));
        }

        // 加密字符串
        private string EncryptString(string plainText, byte[] key, byte[] iv)
        {
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = key;
                aesAlg.IV = iv;

                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
                byte[] encryptedBytes = encryptor.TransformFinalBlock(plainTextBytes, 0, plainTextBytes.Length);

                return Convert.ToBase64String(encryptedBytes);
            }
        }

        // 解密字符串
        private string DecryptString(string cipherText, byte[] key, byte[] iv)
        {
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = key;
                aesAlg.IV = iv;

                ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                byte[] cipherTextBytes = Convert.FromBase64String(cipherText);
                byte[] decryptedBytes = decryptor.TransformFinalBlock(cipherTextBytes, 0, cipherTextBytes.Length);

                return Encoding.UTF8.GetString(decryptedBytes);
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            if (File.Exists("encryption.key")) File.Delete("encryption.key");

            if (File.Exists("config.encrypted")) File.Delete("config.encrypted");

            txbKey.Text = "";
            cbxModel.SelectedIndex = 0;
            cbxEnterKey.SelectedIndex = 0;
            lblResult.Text = "";
            txbUrl.Text = ""; // 清空 API 地址，让用户自行选择
            txbUrl.ReadOnly = false;
            _isCloudConnection = false;
            txbKey.Enabled = false;
        }

        private void pbxPassword_Click(object sender, EventArgs e)
        {
            if (txbKey.UseSystemPasswordChar)
            {
                pbxPassword.BackgroundImage = Properties.Resources.eye_open;
                txbKey.UseSystemPasswordChar = false;
            }
            else
            {
                pbxPassword.BackgroundImage = Properties.Resources.eye_hide;
                txbKey.UseSystemPasswordChar = true;
            }
        }



        // 验证是否为本地API地址
        private bool IsLocalApiUrl(string url)
        {
            try
            {
                Uri uri = new Uri(url);
                string host = uri.Host.ToLower();

                // 检查localhost
                if (host == "localhost" || host == "127.0.0.1")
                    return true;

                // 检查192.168.*.*
                if (host.StartsWith("192.168."))
                    return true;

                // 检查10.0.0.0-10.255.255.255
                if (host.StartsWith("10."))
                    return true;

                // 检查172.16.0.0-172.31.255.255
                string[] parts = host.Split('.');
                if (parts.Length == 4 && parts[0] == "172")
                {
                    if (int.TryParse(parts[1], out int second))
                    {
                        if (second >= 16 && second <= 31)
                            return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        // 从本地API获取模型列表
        private async Task<string[]> GetLocalModels(string apiUrl)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(10);

                    // 构建models端点URL
                    string modelsUrl;

                    // 简化逻辑：直接替换路径部分
                    if (apiUrl.Contains("/v1/chat/completions"))
                    {
                        modelsUrl = apiUrl.Replace("/v1/chat/completions", "/v1/models");
                    }
                    else if (apiUrl.Contains("/chat/completions"))
                    {
                        modelsUrl = apiUrl.Replace("/chat/completions", "/models");
                    }
                    else
                    {
                        // 如果URL不包含标准路径,在基础URL后添加
                        Uri baseUri = new Uri(apiUrl);
                        string baseUrl = $"{baseUri.Scheme}://{baseUri.Authority}";
                        modelsUrl = baseUrl.TrimEnd('/') + "/v1/models";
                    }

                    System.Diagnostics.Debug.WriteLine($"原始API URL: {apiUrl}");
                    System.Diagnostics.Debug.WriteLine($"获取模型列表URL: {modelsUrl}");

                    var response = await client.GetAsync(modelsUrl);

                    System.Diagnostics.Debug.WriteLine($"响应状态码: {response.StatusCode}");

                    if (response.IsSuccessStatusCode)
                    {
                        var content = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"响应内容: {content.Substring(0, Math.Min(500, content.Length))}...");

                        // 简单解析JSON获取模型列表
                        var models = new System.Collections.Generic.List<string>();

                        // 使用System.Text.Json解析
                        using (var doc = System.Text.Json.JsonDocument.Parse(content))
                        {
                            // 标准OpenAI格式: { "data": [...] }
                            if (doc.RootElement.TryGetProperty("data", out var data))
                            {
                                foreach (var model in data.EnumerateArray())
                                {
                                    if (model.TryGetProperty("id", out var id))
                                    {
                                        models.Add(id.GetString());
                                    }
                                }
                            }
                            // Ollama格式: { "models": [...] }
                            else if (doc.RootElement.TryGetProperty("models", out var ollamaModels))
                            {
                                foreach (var model in ollamaModels.EnumerateArray())
                                {
                                    if (model.TryGetProperty("name", out var name))
                                    {
                                        models.Add(name.GetString());
                                    }
                                    else if (model.TryGetProperty("model", out var modelName))
                                    {
                                        models.Add(modelName.GetString());
                                    }
                                }
                            }
                            // 直接是数组格式
                            else if (doc.RootElement.ValueKind == System.Text.Json.JsonValueKind.Array)
                            {
                                foreach (var model in doc.RootElement.EnumerateArray())
                                {
                                    if (model.ValueKind == System.Text.Json.JsonValueKind.String)
                                    {
                                        models.Add(model.GetString());
                                    }
                                    else if (model.TryGetProperty("id", out var id))
                                    {
                                        models.Add(id.GetString());
                                    }
                                    else if (model.TryGetProperty("name", out var name))
                                    {
                                        models.Add(name.GetString());
                                    }
                                }
                            }
                        }

                        System.Diagnostics.Debug.WriteLine($"获取到 {models.Count} 个模型");
                        
                        // 如果标准端点没有获取到模型，尝试Ollama的/api/tags端点
                        if (models.Count == 0)
                        {
                            System.Diagnostics.Debug.WriteLine("尝试Ollama /api/tags 端点...");
                            Uri baseUri = new Uri(apiUrl);
                            string ollamaTagsUrl = $"{baseUri.Scheme}://{baseUri.Authority}/api/tags";
                            
                            try
                            {
                                var ollamaResponse = await client.GetAsync(ollamaTagsUrl);
                                if (ollamaResponse.IsSuccessStatusCode)
                                {
                                    var ollamaContent = await ollamaResponse.Content.ReadAsStringAsync();
                                    System.Diagnostics.Debug.WriteLine($"Ollama响应: {ollamaContent.Substring(0, Math.Min(500, ollamaContent.Length))}...");
                                    
                                    using (var ollamaDoc = System.Text.Json.JsonDocument.Parse(ollamaContent))
                                    {
                                        if (ollamaDoc.RootElement.TryGetProperty("models", out var ollamaModelList))
                                        {
                                            foreach (var model in ollamaModelList.EnumerateArray())
                                            {
                                                if (model.TryGetProperty("name", out var name))
                                                {
                                                    models.Add(name.GetString());
                                                }
                                            }
                                        }
                                    }
                                    System.Diagnostics.Debug.WriteLine($"从Ollama获取到 {models.Count} 个模型");
                                }
                            }
                            catch (Exception ollamaEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"Ollama端点请求失败: {ollamaEx.Message}");
                            }
                        }
                        
                        return models.ToArray();
                    }
                    else
                    {
                        var errorContent = await response.Content.ReadAsStringAsync();
                        System.Diagnostics.Debug.WriteLine($"获取模型列表失败: {response.StatusCode}");
                        System.Diagnostics.Debug.WriteLine($"错误内容: {errorContent}");
                        
                        // 如果/v1/models失败，尝试Ollama的/api/tags端点
                        System.Diagnostics.Debug.WriteLine("尝试Ollama /api/tags 端点...");
                        Uri baseUri = new Uri(apiUrl);
                        string ollamaTagsUrl = $"{baseUri.Scheme}://{baseUri.Authority}/api/tags";
                        
                        try
                        {
                            var ollamaResponse = await client.GetAsync(ollamaTagsUrl);
                            if (ollamaResponse.IsSuccessStatusCode)
                            {
                                var ollamaContent = await ollamaResponse.Content.ReadAsStringAsync();
                                var models = new System.Collections.Generic.List<string>();
                                
                                using (var ollamaDoc = System.Text.Json.JsonDocument.Parse(ollamaContent))
                                {
                                    if (ollamaDoc.RootElement.TryGetProperty("models", out var ollamaModelList))
                                    {
                                        foreach (var model in ollamaModelList.EnumerateArray())
                                        {
                                            if (model.TryGetProperty("name", out var name))
                                            {
                                                models.Add(name.GetString());
                                            }
                                        }
                                    }
                                }
                                System.Diagnostics.Debug.WriteLine($"从Ollama获取到 {models.Count} 个模型");
                                return models.ToArray();
                            }
                        }
                        catch (Exception ollamaEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"Ollama端点请求失败: {ollamaEx.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 如果获取失败,返回空数组
                System.Diagnostics.Debug.WriteLine($"获取模型列表异常: {ex.GetType().Name} - {ex.Message}");
                if (ex.InnerException != null)
                {
                    System.Diagnostics.Debug.WriteLine($"内部异常: {ex.InnerException.Message}");
                }
            }

            return new string[0];
        }

        // API地址文本框内容改变事件
        private async void txbUrl_TextChanged(object sender, EventArgs e)
        {
            string url = txbUrl.Text.Trim();

            if (string.IsNullOrEmpty(url))
                return;

            // 取消之前的请求
            _modelFetchCts?.Cancel();
            _modelFetchCts = new CancellationTokenSource();
            var currentCts = _modelFetchCts;

            // 停止之前的防抖定时器
            _debounceTimer?.Stop();
            _debounceTimer?.Dispose();

            // 检查是否为本地API
            if (IsLocalApiUrl(url))
            {
                _isCloudConnection = false;
                // 本地连接时禁用API Key输入框
                txbKey.Enabled = false;
                lblResult.Text = "检测到本地API地址,正在获取本地模型列表...";

                // 使用防抖定时器，延迟500ms后再获取模型列表
                _debounceTimer = new System.Windows.Forms.Timer();
                _debounceTimer.Interval = 500;
                _debounceTimer.Tick += async (s, args) =>
                {
                    _debounceTimer.Stop();

                    // 检查是否已被取消
                    if (currentCts.IsCancellationRequested)
                        return;

                    // 保存当前选中的模型名称
                    string currentModel = cbxModel.Text;

                    // 获取本地模型列表
                    var models = await GetLocalModels(url);

                    // 再次检查是否已被取消（异步操作完成后）
                    if (currentCts.IsCancellationRequested)
                        return;

                    if (models.Length > 0)
                    {
                        cbxModel.Items.Clear();
                        foreach (var model in models)
                        {
                            cbxModel.Items.Add(model);
                        }

                        // 恢复之前选中的模型（如果存在）
                        if (!string.IsNullOrEmpty(currentModel) && cbxModel.Items.Contains(currentModel))
                        {
                            cbxModel.Text = currentModel;
                        }
                        else if (cbxModel.Items.Count > 0)
                        {
                            cbxModel.SelectedIndex = 0;
                        }

                        lblResult.Text = $"已获取 {models.Length} 个本地模型";
                    }
                    else
                    {
                        lblResult.Text = "无法获取本地模型列表,请确保API服务正在运行";
                    }
                };
                _debounceTimer.Start();
            }
            else
            {
                // 其他云端 API 地址（包括 DeepSeek、OpenAI 等）
                _isCloudConnection = true;
                // 云端连接时启用 API Key 输入框
                txbKey.Enabled = true;
                lblResult.Text = "云端 API 地址";
            }
        }

        // API 地址文本框双击事件（清空内容以便用户输入自定义地址）
        private void txbUrl_DoubleClick(object sender, EventArgs e)
        {
            txbUrl.Clear();
            txbUrl.Focus();
        }
    }
}
