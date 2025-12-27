using System;
using System.Drawing;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace ExcelAddIn
{
    public partial class Form8 : Form
    {
        int i = 0;
        private bool _isCloudConnection = true; // 标识连接类型（true=云端，false=本地）
        private CancellationTokenSource _modelFetchCts = null; // 用于取消模型获取请求
        private System.Windows.Forms.Timer _debounceTimer = null; // 防抖定时器

        public Form8()
        {
            InitializeComponent();
            pbxPassword.BackgroundImage = Properties.Resources.eye_hide;
            txbKey.UseSystemPasswordChar = true;
            txbUrl.Text = @"https://api.deepseek.com/v1/chat/completions";
            txbUrl.ReadOnly = true;
        }

        private void Form8_Load(object sender, EventArgs e)
        {
            DecodeConfig();
            i = 0;
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

        private async Task<string> GetDeepSeekResponse(string apiKey, string model)
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
                        new { role = "user", content = "Hello, DeepSeek!" }
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
            lblResult.Text = "开始检测......";
            txbKey.ReadOnly = true;
            cbxModel.Enabled = false;
            cbxEnterKey.Enabled = false;
            btnTest.Enabled = false;
            btnSave.Enabled = false;
            btnQuit.Enabled = false;
            this.Refresh();

            string result;

            if (!_isCloudConnection)
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
                string model;
                switch (cbxModel.Text)
                {
                    case "deepseek-v3":
                        model = "deepseek-chat";
                        break;
                    case "deepseek-r1":
                        model = "deepseek-reasoner";
                        break;
                    default:
                        model = cbxModel.Text;
                        break;
                }

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
                result = await GetDeepSeekResponse(txbKey.Text, model);
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
        private const string KeyFilePath = "encryption.key"; // 保存密钥和IV的文件路径
        private const string ConfigFilePath = "config.encrypted"; // 保存加密配置信息的文件路径

        //加密保存配置信息
        private void btnSave_Click(object sender, EventArgs e)
        {
            string apiKey = txbKey.Text.Trim();
            string model = string.Empty;

            // 判断是否为本地模型
            if (!_isCloudConnection)
            {
                // 本地模型直接保存模型名称
                model = cbxModel.Text;
            }
            else
            {
                // DeepSeek云端模型按原有逻辑
                switch (cbxModel.SelectedIndex)
                {
                    case 0:
                        model = "deepseek-chat";
                        break;
                    case 1:
                        model = "deepseek-reasoner";
                        break;
                }
            }

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

            if (string.IsNullOrEmpty(apiKey) && _isCloudConnection)
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

            // 验证远程API地址
            if (_isCloudConnection && apiUrl != "https://api.deepseek.com/v1/chat/completions")
            {
                lblResult.Text = "远程连接地址必须是: https://api.deepseek.com/v1/chat/completions";
                return;
            }

            string content = $"api-key^{apiKey};model^{model};api-url^{apiUrl};enter-mode^{enterMode};connection-type^{connectionType};is-cloud-model^{isCloudModel}";

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
                cbxModel.SelectedIndex = 0;
                cbxEnterKey.SelectedIndex = 0;
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
                txbKey.Text = decryptedContent.Split(';')[0].Split('^')[1];
                string model = decryptedContent.Split(';')[1].Split('^')[1];
                txbUrl.Text = decryptedContent.Split(';')[2].Split('^')[1];

                // 读取连接类型(如果配置文件中有的话)
                string connectionType = "cloud"; // 默认为云端
                var parts = decryptedContent.Split(';');
                if (parts.Length >= 5)
                {
                    connectionType = parts[4].Split('^')[1];
                }

                // 设置连接类型标志
                _isCloudConnection = (connectionType == "cloud");

                // 判断是否为本地API
                if (IsLocalApiUrl(txbUrl.Text))
                {
                    _isCloudConnection = false;
                    // 本地模型直接显示模型名称
                    cbxModel.Text = model;
                }
                else
                {
                    _isCloudConnection = true;
                    // DeepSeek云端模型
                    switch (model)
                    {
                        case "deepseek-chat":
                            cbxModel.Text = "deepseek-v3";
                            break;
                        case "deepseek-reasoner":
                            cbxModel.Text = "deepseek-r1";
                            break;
                        default:
                            // 其他情况也显示
                            cbxModel.Text = model;
                            break;
                    }
                }

                switch (decryptedContent.Split(';')[3].Split('^')[1])
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
            txbUrl.Text = @"https://api.deepseek.com/v1/chat/completions";
            txbUrl.ReadOnly = true;
            _isCloudConnection = true;
            txbKey.Enabled = true;
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

        //apiUrl文本框保护机制，双击3次，激活文本框可编辑
        private void txbUrl_DoubleClick(object sender, EventArgs e)
        {
            if (i < 2)
            {
                i++;
                lblResult.Text = $"警告{i}：请不要轻易修改api地址";
            }
            else if (i == 2)
            {
                i++;
                lblResult.Text = $"警告{i}：修改api地址可能导致访问失败，除非你知道必须修改它了，那就再双击我1次试试";

            }
            else
            {
                txbUrl.ReadOnly = false;
                lblResult.Text = "请输入api地址";
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
            else if (url == "https://api.deepseek.com/v1/chat/completions")
            {
                // DeepSeek官方API
                _isCloudConnection = true;
                // 云端连接时启用API Key输入框
                txbKey.Enabled = true;
                cbxModel.Items.Clear();
                cbxModel.Items.Add("deepseek-v3");
                cbxModel.Items.Add("deepseek-r1");
                cbxModel.SelectedIndex = 0;
                lblResult.Text = "DeepSeek官方API";
            }
            else
            {
                // 其他远程地址,必须是DeepSeek官方地址
                _isCloudConnection = true;
                // 云端连接时启用API Key输入框
                txbKey.Enabled = true;
                lblResult.Text = "远程连接地址必须是: https://api.deepseek.com/v1/chat/completions";
            }
        }
    }
}
