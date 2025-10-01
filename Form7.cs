using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http.Json;
using System.Text.Json;
using System.Security.Cryptography;
using System.IO;



namespace ExcelAddIn
{
    public partial class Form7 : Form
    {
        // 初始化 HttpClient（推荐使用 IHttpClientFactory 生产环境）
        private static readonly HttpClient _httpClient = new HttpClient();

        private string _apiKey = string.Empty;           //api key变量
        private string _model = string.Empty;           //模型变量
        private string _apiUrl = string.Empty;         //api接口地址变量
        private string _enterMode = string.Empty;     //回车模式变量

        public Form7()
        {
            InitializeComponent();
            

            // 强制使用 TLS 1.2+ 协议
            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls12 |
                System.Net.SecurityProtocolType.Tls13;

            flowLayoutPanelChat.AutoScroll = true;
            flowLayoutPanelChat.AutoSize = false;

            // 创建自定义右键菜单
            ContextMenuStrip customContextMenu = new ContextMenuStrip();

            // 添加菜单项
            ToolStripMenuItem cutItem = new ToolStripMenuItem("剪切", null, Cut_Click);
            ToolStripMenuItem copyItem = new ToolStripMenuItem("复制", null, Copy_Click);
            ToolStripMenuItem pasteItem = new ToolStripMenuItem("粘贴", null, Paste_Click);
            ToolStripMenuItem selectAllItem = new ToolStripMenuItem("全选", null, SelectAll_Click);
            ToolStripMenuItem clearItem = new ToolStripMenuItem("清空", null, Clear_Click);

            // 将菜单项添加到上下文菜单
            customContextMenu.Items.Add(cutItem);
            customContextMenu.Items.Add(copyItem);
            customContextMenu.Items.Add(pasteItem);
            customContextMenu.Items.Add(selectAllItem);
            customContextMenu.Items.Add(clearItem);

            // 设置richTextBoxInput为多行输入框
            richTextBoxInput.Multiline = true;
            richTextBoxInput.ScrollBars = RichTextBoxScrollBars.Vertical;
            // 将自定义上下文菜单绑定到 RichTextBox
            richTextBoxInput.ContextMenuStrip = customContextMenu; 
        }

        private void Form7_Load(object sender, EventArgs e)
        {
            DecodeConfig();
        }

        private void Cut_Click(object sender, EventArgs e)
        {
            richTextBoxInput.Cut(); // 调用复制功能
        }
        private void Copy_Click(object sender, EventArgs e)
        {
            richTextBoxInput.Copy(); // 调用复制功能
        }

        private void Paste_Click(object sender, EventArgs e)
        {
            richTextBoxInput.Paste(); // 调用粘贴功能
        }

        private void SelectAll_Click(object sender, EventArgs e)
        {
            richTextBoxInput.SelectAll(); // 调用粘贴功能
        }

        private void Clear_Click(object sender, EventArgs e)
        {
            richTextBoxInput.Clear(); // 清空 RichTextBox 内容
        }       

        private async void send_button_Click(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(_apiKey)||string.IsNullOrEmpty(_model))
            {
                prompt_label.Text = "没有获取到API KEY或选择模型，请先使用配置功能进行配置";
                return;
            }
            string userInput = richTextBoxInput.Text.Trim();
            if (string.IsNullOrEmpty(userInput))
            {
                prompt_label.Text = "请输入问题！";
                return;
            }

            try
            {
                // 添加用户消息
                AddChatItem(userInput, true);                
                prompt_label.Text = "思考中...";
                richTextBoxInput.Clear();
                send_button.Enabled = false;

                // 调用DeepSeek API
                var response = await GetDeepSeekResponse(userInput);

                // 添加AI回复
                AddChatItem(response, false);
                prompt_label.Text = "";
            }
            catch (HttpRequestException ex)
            {
                prompt_label.Text = $"网络错误: {ex.Message}";
            }
            catch (JsonException ex)
            {
                prompt_label.Text = $"解析响应失败: {ex.Message}";
            }
            catch (Exception ex)
            {
                prompt_label.Text = $"未知错误: {ex.Message}";
            }
            finally
            {
                send_button.Enabled = true;
                richTextBoxInput.Clear();
            }
        }

        // 对话历史记录
        private List<ChatMessage> _chatHistory = new List<ChatMessage>();

        // 消息模型类
        public class ChatMessage
        {
            public string Role { get; set; }
            public string Content { get; set; }
        }

        //获取对话请求
        private async Task<string> GetDeepSeekResponse(string userInput)
        {
            string apiKey = _apiKey;
            string apiUrl = _apiUrl;

            // 将用户消息加入历史
            _chatHistory.Add(new ChatMessage
            {
                Role = "user",
                Content = userInput
            });

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", apiKey);

                // 使用完整对话历史
                var requestBody = new
                {
                    model = _model,
                    messages = _chatHistory.Select(m => new
                    {
                        role = m.Role,
                        content = m.Content
                    }),
                    temperature = 0.7,
                    max_tokens = 1000
                };

                var response = await client.PostAsJsonAsync(apiUrl, requestBody);
                var responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException($"HTTP Error: {response.StatusCode}");
                }

                var jsonResponse = JsonSerializer.Deserialize<DeepSeekResponse>(responseContent);
                var aiResponse = jsonResponse?.choices[0].message.content.Trim();

                // 新增：将AI回复加入历史
                if (!string.IsNullOrEmpty(aiResponse))
                {
                    _chatHistory.Add(new ChatMessage
                    {
                        Role = "assistant",
                        Content = aiResponse
                    });
                }

                return aiResponse ?? string.Empty;
            }
        }

        // 清空对话历史的方法
        private void btnNewChat_Click(object sender, EventArgs e)
        {
            _chatHistory.Clear();
            flowLayoutPanelChat.Controls.Clear();
            prompt_label.Text = "新对话已开始";
        }

        // DeepSeek API响应模型
        public class DeepSeekResponse
        {
            public Choice[] choices { get; set; }

            public class Choice
            {
                public Message message { get; set; }
            }

            public class Message
            {
                public string role { get; set; }
                public string content { get; set; }
            }
        }

        private void AddChatItem(string text, bool isUser)
        {
            int scrollBarWidth = SystemInformation.VerticalScrollBarWidth;
            int availableWidth = flowLayoutPanelChat.ClientSize.Width - scrollBarWidth - 20;

            RichTextBox richTextBox = new RichTextBox { Text = text, BorderStyle = BorderStyle.None, ReadOnly = true, WordWrap = true, Padding = new Padding(5), ContextMenuStrip = CreateMessageContextMenu(isUser) };

            using (Graphics g = richTextBox.CreateGraphics())
            {
                SizeF textSize = g.MeasureString(text, richTextBox.Font, availableWidth, StringFormat.GenericTypographic);
                richTextBox.Width = Math.Min((int)Math.Ceiling(textSize.Width) + richTextBox.Padding.Horizontal, availableWidth);
                richTextBox.Height = (int)Math.Ceiling(textSize.Height) + richTextBox.Padding.Vertical;
            }

            if (isUser)
            {
                richTextBox.BackColor = Color.LightBlue;
                richTextBox.Tag = "user_message";
                int rtbLeftMargin = flowLayoutPanelChat.ClientSize.Width - richTextBox.Width - 10 - scrollBarWidth;
                if (rtbLeftMargin < 10) rtbLeftMargin = 10;
                richTextBox.Margin = new Padding(rtbLeftMargin, 5, 10, 0);
                flowLayoutPanelChat.Controls.Add(richTextBox);

                FlowLayoutPanel buttonPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = false, Size = new Size(44, 20), BackColor = Color.Transparent, Padding = new Padding(0), Tag = "user_button_panel" };
                Button btnEdit = new Button { Text = "✎", Size = new Size(20, 20), Margin = new Padding(1, 0, 1, 0), FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI Symbol", 7), Cursor = Cursors.Hand };
                btnEdit.Click += (s, e) => { richTextBoxInput.Text = text; richTextBoxInput.Focus(); richTextBoxInput.SelectAll(); };
                buttonPanel.Controls.Add(btnEdit);
                Button btnResend = new Button { Text = "↻", Size = new Size(20, 20), Margin = new Padding(1, 0, 1, 0), FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI Symbol", 7), Cursor = Cursors.Hand };
                btnResend.Click += (s, e) => { richTextBoxInput.Text = text; send_button_Click(null, EventArgs.Empty); };
                buttonPanel.Controls.Add(btnResend);
                int btnLeftMargin = flowLayoutPanelChat.ClientSize.Width - buttonPanel.Width - 10 - scrollBarWidth;
                if (btnLeftMargin < 10) btnLeftMargin = 10;
                buttonPanel.Margin = new Padding(btnLeftMargin, 2, 10, 15);
                flowLayoutPanelChat.Controls.Add(buttonPanel);
                flowLayoutPanelChat.ScrollControlIntoView(buttonPanel);
            }
            else
            {
                richTextBox.BackColor = Color.LightGreen;
                richTextBox.Margin = new Padding(10, 5, 10, 5);
                flowLayoutPanelChat.Controls.Add(richTextBox);
                flowLayoutPanelChat.ScrollControlIntoView(richTextBox);
            }
        }

        private ContextMenuStrip CreateMessageContextMenu(bool isUserMessage)
        {
            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem copyItem = new ToolStripMenuItem("复制");
            copyItem.Click += (s, e) => { if (menu.SourceControl is RichTextBox rtb) { Clipboard.SetText(rtb.SelectionLength > 0 ? rtb.SelectedText : rtb.Text); } };
            menu.Items.Add(copyItem);
            if (isUserMessage)
            {
                ToolStripMenuItem deleteItem = new ToolStripMenuItem("删除");
                deleteItem.Click += (s, e) => { if (menu.SourceControl is RichTextBox rtb && flowLayoutPanelChat.Controls.Contains(rtb)) { flowLayoutPanelChat.Controls.Remove(rtb); rtb.Dispose(); } };
                menu.Items.Add(deleteItem);
            }
            return menu;
        }


        // 创建右键上下文菜单
        private ContextMenuStrip CreateContextMenu(bool isUserMessage)
        {
            ContextMenuStrip menu = new ContextMenuStrip();

            // 复制菜单项（新增选中判断）
            ToolStripMenuItem copyItem = new ToolStripMenuItem("复制");
            copyItem.Click += (sender, e) =>
            {
                if (menu.SourceControl is RichTextBox rtb)
                {
                    // 判断是否有选中文本
                    if (rtb.SelectionLength > 0)
                    {
                        Clipboard.SetText(rtb.SelectedText);
                    }
                    else
                    {
                        Clipboard.SetText(rtb.Text);
                    }
                }
            };

            // 删除菜单项（仅用户消息）
            ToolStripMenuItem deleteItem = null;
            if (isUserMessage)
            {
                deleteItem = new ToolStripMenuItem("删除");
                deleteItem.Click += (sender, e) =>
                {
                    if (menu.SourceControl is RichTextBox rtb &&
                        flowLayoutPanelChat.Controls.Contains(rtb))
                    {
                        flowLayoutPanelChat.Controls.Remove(rtb);
                        rtb.Dispose();
                    }
                };
            }
            // 添加菜单项
            menu.Items.Add(copyItem);
            if (deleteItem != null) menu.Items.Add(deleteItem);

            // 样式设置
            menu.RenderMode = ToolStripRenderMode.Professional;
            menu.BackColor = Color.White;
            menu.Font = new Font("微软雅黑", 9f);

            return menu;
        }

        private void settingsMenuItem_Click(object sender, EventArgs e)
        {
            Form8 form8=new Form8();
            form8.FormClosed += Form8_FormClosed;
            form8.ShowDialog();
        }

        private void Form8_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form7_Load(this, new EventArgs());
        }

        private void exitMenuItem_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        private const string KeyFilePath = "encryption.key"; // 保存密钥和IV的文件路径
        private const string ConfigFilePath = "config.encrypted"; // 保存加密配置信息的文件路径

        //读取配置信息
        private void DecodeConfig()
        {
            if (!File.Exists(ConfigFilePath))
            {
                prompt_label.Text = "配置文件不存在,请先进入设置进行API KEY配置";
                _apiKey = string.Empty;
                _model = string.Empty;
                _apiUrl = string.Empty;
                _enterMode = string.Empty;
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
                _apiKey = decryptedContent.Split(';')[0].Split('^')[1];
                _model = decryptedContent.Split(';')[1].Split('^')[1];
                _apiUrl = decryptedContent.Split(';')[2].Split('^')[1];
                _enterMode= decryptedContent.Split(';')[3].Split('^')[1];
                prompt_label.Text = "可以开始对话了！";
            }
            catch (Exception ex)
            {
                prompt_label.Text = $"发生错误：{ex.Message}";
                _apiKey = string.Empty;
                _model = string.Empty;
                _apiUrl= string.Empty;
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

        private void richTextBoxInput_KeyDown(object sender, KeyEventArgs e)
        {
            switch (_enterMode)
            {
                case "0":
                    if (e.KeyCode == Keys.Enter)
                    {
                        if (e.Shift)
                        {
                            // 手动添加换行符
                            richTextBoxInput.AppendText(Environment.NewLine);
                        }
                        else
                        {
                            // 触发发送操作
                            send_button_Click(null, EventArgs.Empty);
                        }

                        // 阻止默认行为
                        e.Handled = true;          // 标记事件已处理
                        e.SuppressKeyPress = true; // 阻止控件处理按键（避免“叮”声或其他默认行为）
                    }
                    break;
                case "1":
                    if (e.KeyCode == Keys.Enter)
                    {                        
                        richTextBoxInput.AppendText(Environment.NewLine);                       

                        // 阻止默认行为
                        e.Handled = true;          // 标记事件已处理
                        e.SuppressKeyPress = true; // 阻止控件处理按键（避免“叮”声或其他默认行为）
                    }
                    break;
                case "2":
                    if (e.KeyCode == Keys.Enter)
                    {
                        if (e.Control)
                        {
                            // 触发发送操作
                            send_button_Click(null, EventArgs.Empty);

                        }
                        else
                        {
                            // 手动添加换行符
                            richTextBoxInput.AppendText(Environment.NewLine);                            
                        }

                        // 阻止默认行为
                        e.Handled = true;          // 标记事件已处理
                        e.SuppressKeyPress = true; // 阻止控件处理按键（避免“叮”声或其他默认行为）
                    }
                    break;
            }
        }
    }
}

