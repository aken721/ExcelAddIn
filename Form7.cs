using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;



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
        private bool _isCloudConnection = true;       //是否为云端连接（true=云端，false=本地）
        private bool _usePromptEngineering = false;   //是否使用Prompt Engineering模式（本地模型不支持function calling时自动启用）
        private bool _isOllamaApi = false;            //是否为Ollama API（用于添加Ollama特有参数）
        private int _timeoutMinutes = 5;            //请求超时时间，默认5分钟

        private bool _isPromptEngineeringChecked = false;  // 是否勾选了"优先Prompt Engineering"

        private ExcelMcp _excelMcp = null;  // Excel MCP实例
        private string _activeWorkbook = string.Empty;  // 当前活跃的工作簿
        private string _activeWorksheet = string.Empty;  // 当前活跃的工作表

        // 缓存MCP工具定义，避免重复创建
        private List<object> _cachedMcpTools = null;
        
        // 跟踪当前会话中已执行的一次性工具（如create_chart），防止递归时重复执行
        private HashSet<string> _executedOneTimeTools = new HashSet<string>();
        // 一次性工具列表（这些工具在一次用户请求中只应执行一次）
        private static readonly HashSet<string> _oneTimeTools = new HashSet<string> 
        { 
            "create_chart", "create_table", "create_workbook", "create_worksheet", 
            "create_named_range", "save_workbook", "save_workbook_as" 
        };
        
        // 两阶段工具调用：是否启用工具分组模式（用于减少小模型的处理负担）
        private bool _useToolGrouping = true;
        
        // 工具分组定义（用于原生Function Calling的两阶段调用）
        private static readonly Dictionary<string, (string Description, string[] Tools)> _nativeToolGroups = new Dictionary<string, (string Description, string[] Tools)>
        {
            ["cell_rw"] = (
                "单元格读写：读取/写入单元格值、公式、批量操作、查找替换、统计",
                new[] { "set_cell_value", "get_cell_value", "set_range_values", "get_range_values", "set_formula", "get_formula", "validate_formula", "clear_range", "copy_range", "get_current_selection", "get_used_range", "get_last_row", "get_last_column", "get_range_statistics", "find_value", "find_and_replace" }
            ),
            ["format"] = (
                "格式设置：字体、颜色、边框、合并单元格、对齐、条件格式、数字格式",
                new[] { "set_cell_format", "set_border", "set_number_format", "merge_cells", "unmerge_cells", "set_cell_text_wrap", "set_cell_indent", "set_cell_orientation", "set_cell_shrink_to_fit", "apply_conditional_formatting" }
            ),
            ["row_col"] = (
                "行列操作：行高、列宽、插入/删除行列、自动调整、隐藏/显示",
                new[] { "set_row_height", "set_column_width", "insert_rows", "insert_columns", "delete_rows", "delete_columns", "autofit_columns", "autofit_rows", "set_row_visible", "set_column_visible" }
            ),
            ["sheet"] = (
                "工作表操作：创建/删除/重命名/复制/移动工作表、冻结窗格",
                new[] { "get_worksheet_names", "create_worksheet", "rename_worksheet", "delete_worksheet", "copy_worksheet", "move_worksheet", "set_worksheet_visible", "get_worksheet_index", "freeze_panes", "unfreeze_panes" }
            ),
            ["workbook"] = (
                "工作簿操作：创建/打开/保存/关闭工作簿、获取文件信息",
                new[] { "create_workbook", "open_workbook", "save_workbook", "save_workbook_as", "close_workbook", "get_workbook_metadata", "get_current_excel_info", "get_excel_files", "delete_excel_file" }
            ),
            ["data"] = (
                "数据处理：排序、筛选、去重、数据验证、创建表格和图表",
                new[] { "sort_range", "set_auto_filter", "remove_duplicates", "set_data_validation", "get_validation_rules", "create_table", "get_table_names", "create_chart" }
            ),
            ["named"] = (
                "命名区域：创建/删除/查询命名区域",
                new[] { "create_named_range", "delete_named_range", "get_named_ranges", "get_named_range_address" }
            ),
            ["link"] = (
                "批注和超链接：添加/删除批注、内部跳转、外部链接",
                new[] { "add_comment", "get_comment", "delete_comment", "add_hyperlink", "set_hyperlink_formula", "delete_hyperlink" }
            )
        };

        // 检测是否为小模型（参数量小于3B的模型）
        // 小模型处理Function Calling很慢，建议直接使用Prompt Engineering模式
        private bool IsSmallModel(string modelName)
        {
            if (string.IsNullOrEmpty(modelName)) return false;
            
            string nameLower = modelName.ToLower();
            
            // 检测常见的小模型标识
            // 格式通常是: model:0.5b, model:1b, model:1.5b, model:2b 等
            var smallPatterns = new[] { 
                ":0.", ":1b", ":1.5b", ":2b", 
                "-0.", "-1b", "-1.5b", "-2b",
                "0.5b", "0.6b", "1b", "1.5b", "2b",
                "tiny", "mini", "small"
            };
            
            foreach (var pattern in smallPatterns)
            {
                if (nameLower.Contains(pattern))
                {
                    return true;
                }
            }
            
            return false;
        }

        // 日志文件路径（使用用户文档目录，确保可写入）
        private static string _logFilePath = null;
        
        // 获取日志文件路径
        private static string GetLogFilePath()
        {
            if (_logFilePath == null)
            {
                try
                {
                    // 优先使用插件安装目录
                    string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                    if (!string.IsNullOrEmpty(assemblyPath))
                    {
                        string dir = Path.GetDirectoryName(assemblyPath);
                        if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                        {
                            _logFilePath = Path.Combine(dir, "aiDialog.txt");
                            // 测试是否可写
                            try
                            {
                                File.AppendAllText(_logFilePath, "");
                                return _logFilePath;
                            }
                            catch { }
                        }
                    }
                }
                catch { }
                
                // 备用：使用用户文档目录
                try
                {
                    string docPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    _logFilePath = Path.Combine(docPath, "ExcelAddIn_aiDialog.txt");
                }
                catch
                {
                    // 最后备用：使用临时目录
                    _logFilePath = Path.Combine(Path.GetTempPath(), "ExcelAddIn_aiDialog.txt");
                }
            }
            return _logFilePath;
        }

        // 写入日志的方法（追加模式，不删除历史记录）
        private void WriteLog(string category, string message)
        {
            try
            {
                string logPath = GetLogFilePath();
                string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string logEntry = $"[{timestamp}] [{category}]\n{message}\n{"".PadRight(80, '-')}\n";
                File.AppendAllText(logPath, logEntry, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"写入日志失败: {ex.Message}");
            }
        }

        // 获取简化的请求体用于日志记录（排除tools定义和系统提示词中的工具说明）
        private string GetSimplifiedRequestBodyForLog(Dictionary<string, object> requestBody)
        {
            try
            {
                var simplifiedBody = new Dictionary<string, object>();
                
                foreach (var kvp in requestBody)
                {
                    if (kvp.Key == "tools")
                    {
                        // 只记录工具数量，不记录完整定义
                        if (kvp.Value is List<object> toolsList)
                        {
                            simplifiedBody["tools"] = $"[已省略 {toolsList.Count} 个工具定义]";
                        }
                        else
                        {
                            simplifiedBody["tools"] = "[已省略工具定义]";
                        }
                    }
                    else if (kvp.Key == "messages")
                    {
                        // 简化消息列表，只保留用户消息内容
                        var simplifiedMessages = new List<object>();
                        if (kvp.Value is List<object> messages)
                        {
                            foreach (var msg in messages)
                            {
                                var msgDict = msg as dynamic;
                                if (msgDict != null)
                                {
                                    string role = "";
                                    string content = "";
                                    
                                    // 使用反射获取属性
                                    var roleProperty = msg.GetType().GetProperty("role");
                                    var contentProperty = msg.GetType().GetProperty("content");
                                    
                                    if (roleProperty != null)
                                        role = roleProperty.GetValue(msg)?.ToString() ?? "";
                                    if (contentProperty != null)
                                        content = contentProperty.GetValue(msg)?.ToString() ?? "";
                                    
                                    if (role == "system")
                                    {
                                        // 系统提示词只记录前100个字符
                                        simplifiedMessages.Add(new
                                        {
                                            role = role,
                                            content = content.Length > 100 
                                                ? content.Substring(0, 100) + $"... [已省略，共{content.Length}字符]" 
                                                : content
                                        });
                                    }
                                    else
                                    {
                                        // 其他消息保持原样
                                        simplifiedMessages.Add(msg);
                                    }
                                }
                            }
                        }
                        simplifiedBody["messages"] = simplifiedMessages;
                    }
                    else
                    {
                        // 其他字段保持原样
                        simplifiedBody[kvp.Key] = kvp.Value;
                    }
                }
                
                return JsonSerializer.Serialize(simplifiedBody, new JsonSerializerOptions { WriteIndented = true });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"简化请求体失败: {ex.Message}");
                // 如果简化失败，返回原始序列化结果
                return JsonSerializer.Serialize(requestBody, new JsonSerializerOptions { WriteIndented = true });
            }
        }

        // 初始化日志文件（不清空，只添加会话分隔符）
        private void InitLog()
        {
            try
            {
                string logPath = GetLogFilePath();
                var sb = new StringBuilder();
                sb.AppendLine();
                sb.AppendLine("".PadRight(80, '='));
                sb.AppendLine($"=== 新会话开始 ===");
                sb.AppendLine($"时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"日志路径: {logPath}");
                sb.AppendLine($"模型: {(string.IsNullOrEmpty(_model) ? "未配置" : _model)}");
                sb.AppendLine($"API地址: {(string.IsNullOrEmpty(_apiUrl) ? "未配置" : _apiUrl)}");
                sb.AppendLine($"连接类型: {(_isCloudConnection ? "云端" : "本地")}");
                sb.AppendLine($"Prompt Engineering模式: {_usePromptEngineering}");
                sb.AppendLine($"Ollama API: {_isOllamaApi}");
                sb.AppendLine("".PadRight(80, '='));
                sb.AppendLine();
                
                File.AppendAllText(logPath, sb.ToString(), Encoding.UTF8);
                
                // 在调试输出中显示日志路径
                System.Diagnostics.Debug.WriteLine($"AI对话日志路径: {logPath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"初始化日志失败: {ex.Message}");
            }
        }

        // 安全更新prompt_label的方法（确保在UI线程上执行）
        private void SafeUpdatePromptLabel(string text)
        {
            if (prompt_label.InvokeRequired)
            {
                prompt_label.Invoke(new Action(() => prompt_label.Text = text));
            }
            else
            {
                prompt_label.Text = text;
            }
        }

        public Form7()
        {
            InitializeComponent();


            // 强制使用 TLS 1.2+ 协议
            System.Net.ServicePointManager.SecurityProtocol =
                System.Net.SecurityProtocolType.Tls12 |
                System.Net.SecurityProtocolType.Tls13;

            flowLayoutPanelChat.AutoScroll = true;
            flowLayoutPanelChat.AutoSize = false;
            flowLayoutPanelChat.FlowDirection = FlowDirection.TopDown;
            flowLayoutPanelChat.WrapContents = false;
            // 确保滚动条能正常显示
            flowLayoutPanelChat.HorizontalScroll.Enabled = false;
            flowLayoutPanelChat.HorizontalScroll.Visible = false;
            flowLayoutPanelChat.VerticalScroll.Enabled = true;
            flowLayoutPanelChat.VerticalScroll.Visible = true;

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

            // 初始化时默认不勾选"优先Prompt Engineering"
            checkBoxPromptEngineering.Checked = false;
        }

        private async void Form7_Load(object sender, EventArgs e)
        {
            // 显示加载提示
            prompt_label.Text = "正在初始化...";

            // 立即设置默认勾选"使用MCP"，提升响应速度
            checkBoxUseMcp.Checked = true;

            // 并行执行所有初始化任务，提升加载速度
            var configTask = Task.Run(() => DecodeConfig());

            var mcpTask = Task.Run(() =>
            {
                try
                {
                    _excelMcp = new ExcelMcp("./excel_files");
                }
                catch (Exception ex)
                {
                    // 异常信息将在最后统一处理
                    System.Diagnostics.Debug.WriteLine($"初始化Excel MCP失败: {ex.Message}");
                }
            });

            var excelInfoTask = Task.Run(() =>
            {
                try
                {
                    if (ThisAddIn.app != null && ThisAddIn.app.ActiveWorkbook != null)
                    {
                        var activeWorkbook = ThisAddIn.app.ActiveWorkbook;
                        _activeWorkbook = activeWorkbook.Name;

                        if (ThisAddIn.app.ActiveSheet != null)
                        {
                            Microsoft.Office.Interop.Excel.Worksheet activeSheet = ThisAddIn.app.ActiveSheet;
                            _activeWorksheet = activeSheet.Name;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 如果获取失败，不影响程序运行
                    System.Diagnostics.Debug.WriteLine($"获取活跃工作簿失败: {ex.Message}");
                }
            });

            // 等待所有任务完成
            await Task.WhenAll(configTask, mcpTask, excelInfoTask);

            // 配置加载完成后，初始化日志文件（此时配置信息已可用）
            InitLog();

            // 所有任务完成后，在UI线程统一更新界面
            if (_excelMcp == null)
            {
                prompt_label.Text = "初始化Excel MCP失败，请重新打开窗口";
            }
            else if (!File.Exists(ConfigFilePath))
            {
                prompt_label.Text = "配置文件不存在，请先进入设置进行API KEY配置";
            }
            else if (string.IsNullOrEmpty(_apiKey) || string.IsNullOrEmpty(_model))
            {
                prompt_label.Text = "请先进入设置配置API KEY";
            }
            else
            {
                prompt_label.Text = "可以开始对话了！";
            }

            // 更新模型信息标签
            UpdateModelInfoLabel();

            // 根据连接类型设置"优先提示工程"复选框的状态
            UpdatePromptEngineeringCheckBoxState();

            // 添加窗体大小变化事件，用于调整对话框宽度
            this.Resize += Form7_Resize;
        }

        // 窗体大小变化时重新计算对话框宽度
        private void Form7_Resize(object sender, EventArgs e)
        {
            // 遍历所有对话行，更新宽度
            int scrollBarWidth = SystemInformation.VerticalScrollBarWidth;
            int availableWidth = flowLayoutPanelChat.ClientSize.Width - scrollBarWidth - 20;

            foreach (Control control in flowLayoutPanelChat.Controls)
            {
                if (control is Panel rowPanel && (rowPanel.Tag?.ToString() == "user_row" || rowPanel.Tag?.ToString() == "model_row" || rowPanel.Tag?.ToString() == "thinking_row"))
                {
                    // 更新行容器宽度
                    rowPanel.Width = availableWidth;

                    // 查找对话框并更新位置
                    foreach (Control child in rowPanel.Controls)
                    {
                        if (child is Panel chatBubble && (chatBubble.Tag?.ToString() == "user_container" || chatBubble.Tag?.ToString() == "model_container" || chatBubble.Tag?.ToString() == "thinking_placeholder"))
                        {
                            if (chatBubble.Tag?.ToString() == "user_container")
                            {
                                // 用户消息靠右
                                int newLeft = availableWidth - chatBubble.Width;
                                chatBubble.Location = new Point(newLeft, chatBubble.Location.Y);

                                // 更新按钮位置
                                foreach (Control sibling in rowPanel.Controls)
                                {
                                    if (sibling is Panel btnPanel && sibling.Tag?.ToString() == "user_button_panel")
                                    {
                                        btnPanel.Location = new Point(newLeft - btnPanel.Width - 5, btnPanel.Location.Y);
                                    }
                                }
                            }
                            // model_container 和 thinking_placeholder 保持靠左，不需要调整
                        }
                    }
                }
            }
        }

        // 更新模型信息标签
        private void UpdateModelInfoLabel()
        {
            if (string.IsNullOrEmpty(_model))
            {
                labelModelInfo.Text = "未配置模型";
                labelModelInfo.ForeColor = Color.Gray;
            }
            else
            {
                string apiType = _isCloudConnection ? "云端" : "本地";
                labelModelInfo.Text = $"{_model} ({apiType})";
                labelModelInfo.ForeColor = _isCloudConnection ? Color.DodgerBlue : Color.Green;
            }
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

        // 用于存储思考中占位符的引用
        private Panel _thinkingPlaceholder = null;

        private async void send_button_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_apiKey) || string.IsNullOrEmpty(_model))
            {
                prompt_label.Text = "没有获取到API KEY或选择模型，请先使用配置功能进行配置";
                WriteLog("发送失败", "API KEY或模型未配置");
                return;
            }
            string userInput = richTextBoxInput.Text.Trim();
            if (string.IsNullOrEmpty(userInput))
            {
                prompt_label.Text = "请输入问题！";
                return;
            }

            // 记录用户输入
            WriteLog("用户输入", $"内容: {userInput}\n当前模型: {_model}\nAPI地址: {_apiUrl}\n连接类型: {(_isCloudConnection ? "云端" : "本地")}\n用户勾选'优先提示工程': {_isPromptEngineeringChecked}\nPrompt Engineering模式: {_usePromptEngineering}");

            // 清空已执行的一次性工具记录（每次新请求重新开始）
            _executedOneTimeTools.Clear();

            try
            {
                // 添加用户消息
                AddChatItem(userInput, true);
                prompt_label.Text = "思考中...";
                richTextBoxInput.Clear();
                send_button.Enabled = false;

                // 添加思考中占位符
                AddThinkingPlaceholder();

                // 调用DeepSeek API
                var response = await GetDeepSeekResponse(userInput);

                // 移除思考中占位符
                RemoveThinkingPlaceholder();

                // 添加AI回复
                AddChatItem(response, false);
                prompt_label.Text = "";
                
                WriteLog("对话完成", $"AI回复长度: {response?.Length ?? 0}字符");
            }
            catch (TaskCanceledException ex) when (ex.InnerException is TimeoutException)
            {
                RemoveThinkingPlaceholder();
                prompt_label.Text = "请求超时：模型响应时间过长，请稍后重试或尝试更小的模型";
                WriteLog("异常-超时", $"TaskCanceledException(Timeout): {ex.Message}\n内部异常: {ex.InnerException?.Message}");
            }
            catch (TaskCanceledException ex)
            {
                RemoveThinkingPlaceholder();
                prompt_label.Text = "请求已取消：可能是网络问题或模型响应超时，请重试";
                WriteLog("异常-取消", $"TaskCanceledException: {ex.Message}\n堆栈: {ex.StackTrace}");
            }
            catch (OperationCanceledException ex)
            {
                RemoveThinkingPlaceholder();
                prompt_label.Text = "操作已取消：请检查网络连接后重试";
                WriteLog("异常-操作取消", $"OperationCanceledException: {ex.Message}\n堆栈: {ex.StackTrace}");
            }
            catch (HttpRequestException ex)
            {
                RemoveThinkingPlaceholder();
                prompt_label.Text = $"网络错误: {ex.Message}";
                WriteLog("异常-网络错误", $"HttpRequestException: {ex.Message}\n堆栈: {ex.StackTrace}");
            }
            catch (JsonException ex)
            {
                RemoveThinkingPlaceholder();
                prompt_label.Text = $"解析响应失败: {ex.Message}";
                WriteLog("异常-JSON解析", $"JsonException: {ex.Message}\n堆栈: {ex.StackTrace}");
            }
            catch (Exception ex)
            {
                RemoveThinkingPlaceholder();
                // 检查是否是取消相关的异常
                if (ex.Message.Contains("取消") || ex.Message.Contains("cancel") || ex.Message.Contains("Cancel"))
                {
                    prompt_label.Text = "请求已取消：模型响应时间过长或网络问题，请重试";
                    WriteLog("异常-取消相关", $"Exception: {ex.Message}\n堆栈: {ex.StackTrace}");
                }
                else
                {
                    prompt_label.Text = $"未知错误: {ex.Message}";
                    WriteLog("异常-未知错误", $"Exception: {ex.GetType().Name}: {ex.Message}\n堆栈: {ex.StackTrace}");
                }
            }
            finally
            {
                send_button.Enabled = true;
                richTextBoxInput.Clear();
            }
        }

        // 添加思考中占位符
        private void AddThinkingPlaceholder()
        {
            flowLayoutPanelChat.SuspendLayout();
            try
            {
                int scrollBarWidth = SystemInformation.VerticalScrollBarWidth;
                int availableWidth = flowLayoutPanelChat.ClientSize.Width - scrollBarWidth - 20;
                int cornerRadius = 12;

                // 创建占位符面板
                Panel chatBubble = new Panel
                {
                    Size = new Size(80, 36),
                    BackColor = Color.LightGreen,
                    Tag = "thinking_placeholder"
                };

                // 设置圆角
                System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                path.AddArc(0, 0, cornerRadius, cornerRadius, 180, 90);
                path.AddArc(chatBubble.Width - cornerRadius, 0, cornerRadius, cornerRadius, 270, 90);
                path.AddArc(chatBubble.Width - cornerRadius, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 0, 90);
                path.AddArc(0, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 90, 90);
                path.CloseAllFigures();
                chatBubble.Region = new Region(path);

                // 添加"......"文本
                Label thinkingLabel = new Label
                {
                    Text = "......",
                    AutoSize = false,
                    Size = new Size(76, 32),
                    Location = new Point(2, 2),
                    TextAlign = ContentAlignment.MiddleCenter,
                    BackColor = Color.LightGreen,
                    Font = new Font("微软雅黑", 12, FontStyle.Bold)
                };
                chatBubble.Controls.Add(thinkingLabel);

                // 创建行容器
                Panel rowPanel = new Panel
                {
                    Size = new Size(availableWidth, 36),
                    BackColor = Color.Transparent,
                    Tag = "thinking_row"
                };

                chatBubble.Location = new Point(0, 0);
                rowPanel.Controls.Add(chatBubble);
                rowPanel.Margin = new Padding(10, 5, 10, 10);

                flowLayoutPanelChat.Controls.Add(rowPanel);
                flowLayoutPanelChat.ScrollControlIntoView(rowPanel);

                _thinkingPlaceholder = rowPanel;
            }
            finally
            {
                flowLayoutPanelChat.ResumeLayout(true);
            }
        }

        // 移除思考中占位符
        private void RemoveThinkingPlaceholder()
        {
            flowLayoutPanelChat.SuspendLayout();
            try
            {
                if (_thinkingPlaceholder != null && flowLayoutPanelChat.Controls.Contains(_thinkingPlaceholder))
                {
                    flowLayoutPanelChat.Controls.Remove(_thinkingPlaceholder);
                    _thinkingPlaceholder.Dispose();
                    _thinkingPlaceholder = null;
                }
            }
            finally
            {
                flowLayoutPanelChat.ResumeLayout(true);
            }
        }

        // 对话历史记录
        private List<ChatMessage> _chatHistory = new List<ChatMessage>();

        // 消息模型类
        public class ChatMessage
        {
            public string Role { get; set; }
            public string Content { get; set; }
            public List<ToolCall> ToolCalls { get; set; }  // 工具调用
            public string ToolCallId { get; set; }  // 工具调用ID（用于工具响应）
        }

        // 工具调用类
        public class ToolCall
        {
            public string Id { get; set; }
            public string Type { get; set; }
            public FunctionCall Function { get; set; }
        }

        public class FunctionCall
        {
            public string Name { get; set; }
            public string Arguments { get; set; }
        }

        // 工具分组定义：组名 -> (关键词列表, 工具列表)
        private static readonly Dictionary<string, (string[] Keywords, string[] Tools)> _toolGroups = new Dictionary<string, (string[] Keywords, string[] Tools)>
        {
            ["单元格读写"] = (
                new[] { "写入", "输入", "设置值", "读取", "获取值", "单元格", "公式", "清除", "复制", "选中", "范围", "查找", "替换", "统计", "最后一行", "最后一列" },
                new[] { "set_cell_value", "get_cell_value", "set_range_values", "get_range_values", "set_formula", "get_formula", "validate_formula", "clear_range", "copy_range", "get_current_selection", "get_used_range", "get_last_row", "get_last_column", "get_range_statistics", "find_value", "find_and_replace" }
            ),
            ["格式设置"] = (
                new[] { "格式", "颜色", "字体", "背景", "加粗", "斜体", "边框", "合并", "对齐", "居中", "换行", "缩进", "旋转", "条件格式", "数字格式" },
                new[] { "set_cell_format", "set_border", "set_number_format", "merge_cells", "unmerge_cells", "set_cell_text_wrap", "set_cell_indent", "set_cell_orientation", "set_cell_shrink_to_fit", "apply_conditional_formatting" }
            ),
            ["行列操作"] = (
                new[] { "行高", "列宽", "插入行", "插入列", "删除行", "删除列", "自动列宽", "自动行高", "隐藏行", "隐藏列", "显示行", "显示列" },
                new[] { "set_row_height", "set_column_width", "insert_rows", "insert_columns", "delete_rows", "delete_columns", "autofit_columns", "autofit_rows", "set_row_visible", "set_column_visible" }
            ),
            ["工作表操作"] = (
                new[] { "工作表", "表名", "创建表", "新建表", "重命名", "删除表", "复制表", "移动表", "隐藏表", "显示表", "冻结", "取消冻结", "sheet", "切换", "激活", "跳转到" },
                new[] { "activate_worksheet", "get_worksheet_names", "create_worksheet", "rename_worksheet", "delete_worksheet", "copy_worksheet", "move_worksheet", "set_worksheet_visible", "get_worksheet_index", "freeze_panes", "unfreeze_panes" }
            ),
            ["工作簿操作"] = (
                new[] { "工作簿", "文件", "新建", "打开", "保存", "另存为", "关闭", "excel文件" },
                new[] { "create_workbook", "open_workbook", "save_workbook", "save_workbook_as", "close_workbook", "get_workbook_metadata", "get_current_excel_info", "get_excel_files", "delete_excel_file" }
            ),
            ["数据处理"] = (
                new[] { "排序", "筛选", "去重", "删除重复", "数据验证", "表格", "图表", "chart" },
                new[] { "sort_range", "set_auto_filter", "remove_duplicates", "set_data_validation", "get_validation_rules", "create_table", "get_table_names", "create_chart" }
            ),
            ["命名区域"] = (
                new[] { "命名区域", "命名范围", "名称管理" },
                new[] { "create_named_range", "delete_named_range", "get_named_ranges", "get_named_range_address" }
            ),
            ["批注超链接"] = (
                new[] { "批注", "注释", "超链接", "链接", "跳转" },
                new[] { "add_comment", "get_comment", "delete_comment", "add_hyperlink", "set_hyperlink_formula", "delete_hyperlink" }
            )
        };

        // 工具详细说明
        private static readonly Dictionary<string, string> _toolDetails = new Dictionary<string, string>
        {
            // 单元格读写 - 所有工具都支持可选的sheetName参数来指定目标工作表
            ["set_cell_value"] = "设置单元格值。参数: row(int), column(int), value(string), sheetName(可选,指定工作表名)",
            ["get_cell_value"] = "获取单元格值。参数: row(int), column(int), sheetName(可选)",
            ["set_range_values"] = "批量设置值。参数: rangeAddress(如\"A1:C3\"), data(JSON二维数组), sheetName(可选)",
            ["get_range_values"] = "获取区域值。参数: rangeAddress, sheetName(可选)",
            ["set_formula"] = "设置公式。参数: cellAddress, formula, sheetName(可选)",
            ["get_formula"] = "获取公式。参数: cellAddress, sheetName(可选)",
            ["validate_formula"] = "验证公式语法。参数: formula",
            ["clear_range"] = "清除范围。参数: rangeAddress, clearType(all/contents/formats), sheetName(可选)",
            ["copy_range"] = "复制范围。参数: sourceRange, targetRange, sheetName(可选)",
            ["get_current_selection"] = "获取当前选中单元格。无参数",
            ["get_used_range"] = "获取已使用范围。参数: sheetName(可选)",
            ["get_last_row"] = "获取最后有数据的行。参数: columnIndex(可选), sheetName(可选)",
            ["get_last_column"] = "获取最后有数据的列。参数: rowIndex(可选), sheetName(可选)",
            ["get_range_statistics"] = "获取范围统计。参数: rangeAddress, sheetName(可选)",
            ["find_value"] = "查找值。参数: searchValue, sheetName(可选)",
            ["find_and_replace"] = "查找替换。参数: findValue, replaceValue, sheetName(可选)",
            // 格式设置
            ["set_cell_format"] = "设置单元格格式。参数: rangeAddress(如\"A1\"或\"F9\"), backgroundColor(背景色,如\"#FFFF00\"黄色), fontColor(字体颜色), bold, italic, fontSize, sheetName(可选)",
            ["set_border"] = "设置边框。参数: rangeAddress, borderType(all/outline), lineStyle(continuous/dash/dot), sheetName(可选)",
            ["set_number_format"] = "数字格式。参数: rangeAddress, formatCode, sheetName(可选)",
            ["merge_cells"] = "合并单元格。参数: rangeAddress",
            ["unmerge_cells"] = "取消合并。参数: rangeAddress",
            ["set_cell_text_wrap"] = "自动换行。参数: rangeAddress, wrap(bool)",
            ["set_cell_indent"] = "缩进。参数: rangeAddress, indentLevel(int)",
            ["set_cell_orientation"] = "文字旋转。参数: rangeAddress, degrees(-90到90)",
            ["set_cell_shrink_to_fit"] = "缩小填充。参数: rangeAddress, shrink(bool)",
            ["apply_conditional_formatting"] = "条件格式。参数: rangeAddress, formatType, criteria",
            // 行列操作
            ["set_row_height"] = "设置行高。参数: rowNumber(int), height(double)",
            ["set_column_width"] = "设置列宽。参数: columnNumber(int), width(double)",
            ["insert_rows"] = "插入行。参数: rowIndex, count",
            ["insert_columns"] = "插入列。参数: columnIndex, count",
            ["delete_rows"] = "删除行。参数: rowIndex, count",
            ["delete_columns"] = "删除列。参数: columnIndex, count",
            ["autofit_columns"] = "自动列宽。参数: rangeAddress",
            ["autofit_rows"] = "自动行高。参数: rangeAddress",
            ["set_row_visible"] = "显示/隐藏行。参数: rowIndex, visible(bool)",
            ["set_column_visible"] = "显示/隐藏列。参数: columnIndex, visible(bool)",
            // 工作表操作
            // 工作表操作
            ["activate_worksheet"] = "激活/切换到指定工作表（在该表上进行后续操作前必须先激活）。参数: sheetName",
            ["get_worksheet_names"] = "获取所有表名。无参数",
            ["create_worksheet"] = "创建表。参数: sheetName",
            ["rename_worksheet"] = "重命名表。参数: oldSheetName, newSheetName",
            ["delete_worksheet"] = "删除表。参数: sheetName",
            ["copy_worksheet"] = "复制表。参数: sourceSheetName, targetSheetName",
            ["move_worksheet"] = "移动表。参数: sheetName, position(int)",
            ["set_worksheet_visible"] = "显示/隐藏表。参数: sheetName, visible(bool)",
            ["get_worksheet_index"] = "获取表索引。参数: sheetName",
            ["freeze_panes"] = "冻结窗格。参数: row, column",
            ["unfreeze_panes"] = "取消冻结。无参数",
            // 工作簿操作
            ["create_workbook"] = "创建工作簿。参数: fileName",
            ["open_workbook"] = "打开工作簿。参数: fileName",
            ["save_workbook"] = "保存工作簿。无参数",
            ["save_workbook_as"] = "另存为。参数: fileName, newFileName",
            ["close_workbook"] = "关闭工作簿。参数: fileName(可选)",
            ["get_workbook_metadata"] = "获取工作簿信息。无参数",
            ["get_current_excel_info"] = "获取当前Excel信息。无参数",
            ["get_excel_files"] = "获取文件列表。无参数",
            ["delete_excel_file"] = "删除文件。参数: fileName",
            // 数据处理
            ["sort_range"] = "排序。参数: rangeAddress, sortColumnIndex, ascending(bool)",
            ["set_auto_filter"] = "自动筛选。参数: rangeAddress",
            ["remove_duplicates"] = "删除重复。参数: rangeAddress, columnIndices(JSON数组)",
            ["set_data_validation"] = "数据验证。参数: rangeAddress, validationType, formula1",
            ["get_validation_rules"] = "获取验证规则。参数: rangeAddress",
            ["create_table"] = "创建表格。参数: rangeAddress, tableName",
            ["get_table_names"] = "获取表格名。无参数",
            ["create_chart"] = "创建图表。参数: dataRange(必需), chartType(可选,默认column), title(可选)",
            // 命名区域
            ["create_named_range"] = "创建命名区域。参数: rangeName, rangeAddress",
            ["delete_named_range"] = "删除命名区域。参数: rangeName",
            ["get_named_ranges"] = "获取所有命名区域。无参数",
            ["get_named_range_address"] = "获取命名区域地址。参数: rangeName",
            // 批注和超链接
            ["add_comment"] = "添加批注。参数: cellAddress, commentText",
            ["get_comment"] = "获取批注。参数: cellAddress",
            ["delete_comment"] = "删除批注。参数: cellAddress",
            ["add_hyperlink"] = "添加外部链接。参数: cellAddress, url, displayText",
            ["set_hyperlink_formula"] = "添加内部跳转。参数: cellAddress, targetLocation(如\"Sheet2!A1\"), displayText",
            ["delete_hyperlink"] = "删除超链接。参数: cellAddress"
        };

        // 根据用户输入选择相关的工具组
        private List<string> SelectRelevantToolGroups(string userInput)
        {
            var selectedGroups = new List<string>();
            string inputLower = userInput.ToLower();

            foreach (var group in _toolGroups)
            {
                foreach (var keyword in group.Value.Keywords)
                {
                    if (inputLower.Contains(keyword.ToLower()))
                    {
                        if (!selectedGroups.Contains(group.Key))
                        {
                            selectedGroups.Add(group.Key);
                        }
                        break;
                    }
                }
            }

            // 如果选择了"数据处理"组（图表、排序等），必须同时包含"单元格读写"组（find_value、get_range_values）
            if (selectedGroups.Contains("数据处理") && !selectedGroups.Contains("单元格读写"))
            {
                selectedGroups.Insert(0, "单元格读写"); // 插入到最前面，强调先查找
            }
            
            // 如果用户提到分析、报告等，也需要单元格读写组
            if ((inputLower.Contains("分析") || inputLower.Contains("报告") || inputLower.Contains("变化") || inputLower.Contains("趋势")) 
                && !selectedGroups.Contains("单元格读写"))
            {
                selectedGroups.Insert(0, "单元格读写");
            }

            // 如果没有匹配到任何组，默认返回"单元格读写"组（最常用）
            if (selectedGroups.Count == 0)
            {
                selectedGroups.Add("单元格读写");
            }

            return selectedGroups;
        }

        // 生成Prompt Engineering模式的系统提示词（用于不支持原生Function Calling的本地模型）
        private string GetPromptEngineeringSystemPrompt(string userInput = null)
        {
            var sb = new StringBuilder();
            
            // 获取当前环境信息
            string currentCell = "A1";
            int currentRow = 1;
            int currentCol = 1;
            string selectionAddress = "A1";
            try
            {
                if (ThisAddIn.app?.Selection != null)
                {
                    Microsoft.Office.Interop.Excel.Range selection = ThisAddIn.app.Selection;
                    currentCell = selection.Address.Replace("$", "");
                    selectionAddress = currentCell;
                    currentRow = selection.Row;
                    currentCol = selection.Column;
                }
            }
            catch { }

            // 更新活跃工作表信息
            try
            {
                if (ThisAddIn.app?.ActiveWorkbook != null)
                {
                    _activeWorkbook = ThisAddIn.app.ActiveWorkbook.Name;
                    if (ThisAddIn.app.ActiveSheet != null)
                    {
                        Microsoft.Office.Interop.Excel.Worksheet ws = ThisAddIn.app.ActiveSheet;
                        _activeWorksheet = ws.Name;
                    }
                }
            }
            catch { }

            string sheetName = string.IsNullOrEmpty(_activeWorksheet) ? "Sheet1" : _activeWorksheet;
            string colLetter = GetColumnLetter(currentCol);

            // 极简提示词，强调直接输出工具调用
            sb.AppendLine("你是Excel工具调用助手。收到指令后，直接输出工具调用JSON，不要解释。");
            sb.AppendLine();
            sb.AppendLine($"当前：工作表=\"{sheetName}\"，选中区域={selectionAddress}");
            sb.AppendLine();
            sb.AppendLine("输出格式：");
            sb.AppendLine("<tool_calls>");
            sb.AppendLine("[{\"name\": \"工具名\", \"arguments\": {参数}}]");
            sb.AppendLine("</tool_calls>");
            sb.AppendLine();

            // 根据用户输入智能选择工具组
            List<string> relevantGroups;
            if (!string.IsNullOrEmpty(userInput))
            {
                relevantGroups = SelectRelevantToolGroups(userInput);
            }
            else
            {
                relevantGroups = new List<string> { "单元格读写", "格式设置", "工作表操作" };
            }

            sb.AppendLine("可用工具：");
            foreach (var groupName in relevantGroups)
            {
                if (_toolGroups.TryGetValue(groupName, out var groupInfo))
                {
                    foreach (var tool in groupInfo.Tools)
                    {
                        if (_toolDetails.TryGetValue(tool, out var detail))
                        {
                            sb.AppendLine($"- {tool}: {detail}");
                        }
                    }
                }
            }
            sb.AppendLine();
            
            // 检测用户意图，提供针对性指导
            string inputLower = userInput?.ToLower() ?? "";
            bool wantsChart = inputLower.Contains("图表") || inputLower.Contains("折线") || inputLower.Contains("曲线") || 
                             inputLower.Contains("柱状") || inputLower.Contains("饼图") || inputLower.Contains("chart");
            bool wantsAnalysis = inputLower.Contains("分析") || inputLower.Contains("报告") || inputLower.Contains("变化");
            bool wantsRead = inputLower.Contains("读取") || inputLower.Contains("获取") || inputLower.Contains("查看") || inputLower.Contains("是多少");
            bool hasSelectedRange = inputLower.Contains("选中") || inputLower.Contains("选择") || inputLower.Contains("当前区域");
            
            if (wantsChart && hasSelectedRange)
            {
                // 用户要基于选中区域创建图表
                sb.AppendLine("📊 图表创建任务（选中区域）：");
                sb.AppendLine($"直接用选中区域创建图表：");
                sb.AppendLine("<tool_calls>");
                sb.AppendLine($"[{{\"name\": \"create_chart\", \"arguments\": {{\"dataRange\": \"{selectionAddress}\", \"chartType\": \"line\", \"title\": \"数据图表\"}}}}]");
                sb.AppendLine("</tool_calls>");
            }
            else if (wantsChart || wantsAnalysis || wantsRead)
            {
                // 用户要创建图表/分析/读取数据，必须先查找
                sb.AppendLine("⚠️ 重要：必须先查找数据位置，禁止编造数据！");
                sb.AppendLine();
                sb.AppendLine("标准流程：");
                sb.AppendLine("第1步：用find_value查找用户提到的关键词位置");
                sb.AppendLine("<tool_calls>");
                sb.AppendLine("[{\"name\": \"find_value\", \"arguments\": {\"searchValue\": \"用户提到的关键词\"}}]");
                sb.AppendLine("</tool_calls>");
                sb.AppendLine();
                sb.AppendLine("第2步：根据find_value返回的位置，用get_range_values读取数据");
                if (wantsChart)
                {
                    sb.AppendLine("第3步：用create_chart创建图表（dataRange填实际数据范围）");
                }
                sb.AppendLine();
                sb.AppendLine("每次只执行一步，等待结果后再继续。");
            }
            
            sb.AppendLine();
            sb.AppendLine("规则：");
            sb.AppendLine("1. 直接输出<tool_calls>JSON");
            sb.AppendLine("2. 禁止编造数据，必须从Excel读取");
            sb.AppendLine("3. 每次只输出一个工具调用，等待结果");

            return sb.ToString();
        }
        
        // 将列号转换为字母（1=A, 2=B, 3=C...）
        private string GetColumnLetter(int columnNumber)
        {
            string result = "";
            while (columnNumber > 0)
            {
                columnNumber--;
                result = (char)('A' + columnNumber % 26) + result;
                columnNumber /= 26;
            }
            return result;
        }

        // 解析Prompt Engineering模式下AI响应中的工具调用
        private List<PromptToolCall> ParsePromptToolCalls(string response)
        {
            var toolCalls = new List<PromptToolCall>();

            try
            {
                // 格式1: 处理 <tool_calls>...</tool_calls> 块（标准格式）
                int searchStart = 0;
                while (true)
                {
                    int startIndex = response.IndexOf("<tool_calls>", searchStart);
                    if (startIndex == -1) break;
                    
                    int endIndex = response.IndexOf("</tool_calls>", startIndex);
                    if (endIndex == -1)
                    {
                        // 没有闭合标签，尝试找下一个<tool_calls>或字符串结尾
                        int nextStart = response.IndexOf("<tool_calls>", startIndex + 12);
                        if (nextStart == -1)
                        {
                            // 没有下一个，取到字符串结尾
                            endIndex = response.Length;
                        }
                        else
                        {
                            endIndex = nextStart;
                        }
                    }
                    
                    string jsonContent = response.Substring(startIndex + 12, endIndex - startIndex - 12).Trim();
                    // 移除可能的闭合标签
                    jsonContent = jsonContent.Replace("</tool_calls>", "").Trim();
                    
                    if (!string.IsNullOrEmpty(jsonContent))
                    {
                        var parsed = ParseJsonToolCalls(jsonContent);
                        toolCalls.AddRange(parsed);
                    }
                    
                    searchStart = endIndex;
                    if (searchStart >= response.Length) break;
                }
                
                if (toolCalls.Count > 0)
                {
                    return toolCalls;
                }

                // 格式2: 处理多个连续的JSON数组 [{...}][{...}]
                // 这种情况是模型没有用标签包裹，直接输出多个JSON数组
                int jsonArrayStart = response.IndexOf("[{");
                if (jsonArrayStart != -1)
                {
                    // 提取所有JSON数组
                    string remaining = response.Substring(jsonArrayStart);
                    int pos = 0;
                    while (pos < remaining.Length)
                    {
                        int arrayStart = remaining.IndexOf("[", pos);
                        if (arrayStart == -1) break;
                        
                        int bracketCount = 0;
                        int arrayEnd = -1;
                        for (int i = arrayStart; i < remaining.Length; i++)
                        {
                            if (remaining[i] == '[') bracketCount++;
                            else if (remaining[i] == ']') bracketCount--;
                            if (bracketCount == 0)
                            {
                                arrayEnd = i;
                                break;
                            }
                        }
                        
                        if (arrayEnd != -1)
                        {
                            string jsonContent = remaining.Substring(arrayStart, arrayEnd - arrayStart + 1).Trim();
                            if (jsonContent.Contains("\"name\""))
                            {
                                var parsed = ParseJsonToolCalls(jsonContent);
                                toolCalls.AddRange(parsed);
                            }
                            pos = arrayEnd + 1;
                        }
                        else
                        {
                            break;
                        }
                    }
                    
                    if (toolCalls.Count > 0)
                    {
                        return toolCalls;
                    }
                }

                // 格式3: ```json ... ``` 代码块格式
                int codeBlockStart = response.IndexOf("```json");
                if (codeBlockStart == -1)
                {
                    codeBlockStart = response.IndexOf("```");
                }
                if (codeBlockStart != -1)
                {
                    int codeBlockEnd = response.IndexOf("```", codeBlockStart + 3);
                    if (codeBlockEnd != -1)
                    {
                        int contentStart = response.IndexOf('\n', codeBlockStart);
                        if (contentStart != -1 && contentStart < codeBlockEnd)
                        {
                            string codeContent = response.Substring(contentStart + 1, codeBlockEnd - contentStart - 1).Trim();
                            int jsonStart = codeContent.IndexOf('[');
                            if (jsonStart != -1)
                            {
                                int bracketCount = 0;
                                int jsonEnd = -1;
                                for (int i = jsonStart; i < codeContent.Length; i++)
                                {
                                    if (codeContent[i] == '[') bracketCount++;
                                    else if (codeContent[i] == ']') bracketCount--;
                                    if (bracketCount == 0)
                                    {
                                        jsonEnd = i;
                                        break;
                                    }
                                }
                                if (jsonEnd != -1)
                                {
                                    string jsonContent = codeContent.Substring(jsonStart, jsonEnd - jsonStart + 1).Trim();
                                    if (jsonContent.Contains("\"name\""))
                                    {
                                        return ParseJsonToolCalls(jsonContent);
                                    }
                                }
                            }
                        }
                    }
                }

                // 格式3: 直接是JSON数组 [{...}]
                int directJsonStart = response.IndexOf("[{");
                if (directJsonStart != -1)
                {
                    int bracketCount = 0;
                    int jsonEnd = -1;
                    for (int i = directJsonStart; i < response.Length; i++)
                    {
                        if (response[i] == '[') bracketCount++;
                        else if (response[i] == ']') bracketCount--;
                        if (bracketCount == 0)
                        {
                            jsonEnd = i;
                            break;
                        }
                    }
                    if (jsonEnd != -1)
                    {
                        string jsonContent = response.Substring(directJsonStart, jsonEnd - directJsonStart + 1).Trim();
                        if (jsonContent.Contains("\"name\""))
                        {
                            return ParseJsonToolCalls(jsonContent);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"解析工具调用失败: {ex.Message}");
            }

            return toolCalls;
        }

        // 解析JSON格式的工具调用
        private List<PromptToolCall> ParseJsonToolCalls(string jsonContent)
        {
            var toolCalls = new List<PromptToolCall>();
            WriteLog("JSON解析", $"原始JSON内容:\n{jsonContent}");

            try
            {
                // 清理JSON内容
                jsonContent = jsonContent.Trim();
                
                // 替换中文引号为英文引号
                jsonContent = jsonContent.Replace("\u201c", "\"").Replace("\u201d", "\"");
                jsonContent = jsonContent.Replace("\u2018", "'").Replace("\u2019", "'");
                
                // 修复不完整的JSON数组（缺少闭合的]）
                if (jsonContent.StartsWith("[") && !jsonContent.EndsWith("]"))
                {
                    // 计算括号数量
                    int openBrackets = jsonContent.Count(c => c == '[');
                    int closeBrackets = jsonContent.Count(c => c == ']');
                    int openBraces = jsonContent.Count(c => c == '{');
                    int closeBraces = jsonContent.Count(c => c == '}');
                    
                    // 补全缺少的闭合括号
                    for (int i = 0; i < openBraces - closeBraces; i++)
                        jsonContent += "}";
                    for (int i = 0; i < openBrackets - closeBrackets; i++)
                        jsonContent += "]";
                    
                    WriteLog("JSON修复", $"补全闭合括号后:\n{jsonContent}");
                }
                
                // 替换全角字符
                jsonContent = jsonContent.Replace("\uff1a", ":").Replace("\uff0c", ",");
                
                WriteLog("JSON解析", $"清理后的JSON:\n{jsonContent}");
                System.Diagnostics.Debug.WriteLine($"清理后的JSON: {jsonContent}");
                
                // 处理多个JSON数组连续的情况（如 [...][...]）
                // 只取第一个完整的JSON数组
                if (jsonContent.StartsWith("["))
                {
                    int bracketCount = 0;
                    int firstArrayEnd = -1;
                    for (int i = 0; i < jsonContent.Length; i++)
                    {
                        if (jsonContent[i] == '[') bracketCount++;
                        else if (jsonContent[i] == ']') bracketCount--;
                        
                        if (bracketCount == 0)
                        {
                            firstArrayEnd = i;
                            break;
                        }
                    }
                    
                    if (firstArrayEnd != -1 && firstArrayEnd < jsonContent.Length - 1)
                    {
                        // 有多余内容，只取第一个数组
                        jsonContent = jsonContent.Substring(0, firstArrayEnd + 1);
                        WriteLog("JSON解析", $"截取第一个数组后:\n{jsonContent}");
                    }
                }

                using (var doc = JsonDocument.Parse(jsonContent))
                {
                    if (doc.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var element in doc.RootElement.EnumerateArray())
                        {
                            try
                            {
                                // 标准格式: {"name": "xxx", "arguments": {...}}
                                if (element.ValueKind == JsonValueKind.Object)
                                {
                                    var toolCall = new PromptToolCall
                                    {
                                        Id = Guid.NewGuid().ToString(),
                                        Name = element.GetProperty("name").GetString(),
                                        ArgumentsJson = element.TryGetProperty("arguments", out var args) ? args.GetRawText() : "{}"
                                    };
                                    toolCalls.Add(toolCall);
                                    WriteLog("JSON解析成功", $"工具: {toolCall.Name}, 参数: {toolCall.ArgumentsJson}");
                                    System.Diagnostics.Debug.WriteLine($"成功解析工具: {toolCall.Name}, 参数: {toolCall.ArgumentsJson}");
                                }
                                // 错误格式: ["tool_name", {...}] - 尝试修复
                                else if (element.ValueKind == JsonValueKind.Array)
                                {
                                    var arr = element.EnumerateArray().ToArray();
                                    if (arr.Length >= 2 && arr[0].ValueKind == JsonValueKind.String)
                                    {
                                        var toolCall = new PromptToolCall
                                        {
                                            Id = Guid.NewGuid().ToString(),
                                            Name = arr[0].GetString(),
                                            ArgumentsJson = arr[1].ValueKind == JsonValueKind.Object ? arr[1].GetRawText() : "{}"
                                        };
                                        toolCalls.Add(toolCall);
                                        WriteLog("JSON解析成功(修复数组格式)", $"工具: {toolCall.Name}, 参数: {toolCall.ArgumentsJson}");
                                        System.Diagnostics.Debug.WriteLine($"成功解析工具(修复数组格式): {toolCall.Name}, 参数: {toolCall.ArgumentsJson}");
                                    }
                                }
                            }
                            catch (Exception innerEx)
                            {
                                WriteLog("JSON解析失败", $"解析单个工具调用失败: {innerEx.Message}, 元素: {element.GetRawText()}");
                                System.Diagnostics.Debug.WriteLine($"解析单个工具调用失败: {innerEx.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"JSON解析失败: {ex.Message}, 内容: {jsonContent}");
                WriteLog("JSON解析异常", $"错误: {ex.Message}\n内容: {jsonContent}");
                
                // 尝试使用正则表达式提取工具调用
                try
                {
                    // 提取工具名称
                    var nameRegex = new System.Text.RegularExpressions.Regex(@"""name""\s*:\s*""([^""]+)""");
                    var nameMatch = nameRegex.Match(jsonContent);
                    
                    if (nameMatch.Success)
                    {
                        string toolName = nameMatch.Groups[1].Value;
                        
                        // 提取arguments部分（支持嵌套大括号）
                        int argsStart = jsonContent.IndexOf("\"arguments\"");
                        if (argsStart != -1)
                        {
                            int braceStart = jsonContent.IndexOf('{', argsStart);
                            if (braceStart != -1)
                            {
                                int braceCount = 0;
                                int braceEnd = -1;
                                for (int i = braceStart; i < jsonContent.Length; i++)
                                {
                                    if (jsonContent[i] == '{') braceCount++;
                                    else if (jsonContent[i] == '}') braceCount--;
                                    if (braceCount == 0)
                                    {
                                        braceEnd = i;
                                        break;
                                    }
                                }
                                
                                if (braceEnd != -1)
                                {
                                    string argsJson = jsonContent.Substring(braceStart, braceEnd - braceStart + 1);
                                    toolCalls.Add(new PromptToolCall
                                    {
                                        Id = Guid.NewGuid().ToString(),
                                        Name = toolName,
                                        ArgumentsJson = argsJson
                                    });
                                    WriteLog("正则提取成功", $"工具: {toolName}, 参数: {argsJson}");
                                    System.Diagnostics.Debug.WriteLine($"正则提取工具: {toolName}, 参数: {argsJson}");
                                }
                                else
                                {
                                    // 无法找到完整的arguments，使用空对象
                                    toolCalls.Add(new PromptToolCall
                                    {
                                        Id = Guid.NewGuid().ToString(),
                                        Name = toolName,
                                        ArgumentsJson = "{}"
                                    });
                                    WriteLog("正则提取(无参数)", $"工具: {toolName}");
                                    System.Diagnostics.Debug.WriteLine($"正则提取工具(无参数): {toolName}");
                                }
                            }
                        }
                        else
                        {
                            // 没有arguments字段
                            toolCalls.Add(new PromptToolCall
                            {
                                Id = Guid.NewGuid().ToString(),
                                Name = toolName,
                                ArgumentsJson = "{}"
                            });
                            WriteLog("正则提取(无arguments)", $"工具: {toolName}");
                            System.Diagnostics.Debug.WriteLine($"正则提取工具(无arguments): {toolName}");
                        }
                    }
                }
                catch (Exception regexEx)
                {
                    WriteLog("正则提取失败", $"错误: {regexEx.Message}");
                    System.Diagnostics.Debug.WriteLine($"正则提取也失败: {regexEx.Message}");
                }
            }

            return toolCalls;
        }

        // 从AI响应中移除工具调用标签，获取纯文本内容
        private string RemoveToolCallTags(string response)
        {
            // 尝试移除 <tool_calls>...</tool_calls>
            int startIndex = response.IndexOf("<tool_calls>");
            int endIndex = response.IndexOf("</tool_calls>");

            if (startIndex != -1 && endIndex != -1)
            {
                string before = response.Substring(0, startIndex).Trim();
                string after = response.Substring(endIndex + 13).Trim();
                return (before + " " + after).Trim();
            }

            // 尝试移除 ```json ... ``` 代码块
            int codeBlockStart = response.IndexOf("```json");
            if (codeBlockStart == -1)
            {
                codeBlockStart = response.IndexOf("```");
            }
            if (codeBlockStart != -1)
            {
                int codeBlockEnd = response.IndexOf("```", codeBlockStart + 3);
                if (codeBlockEnd != -1)
                {
                    string before = response.Substring(0, codeBlockStart).Trim();
                    string after = (codeBlockEnd + 3 < response.Length) ? response.Substring(codeBlockEnd + 3).Trim() : "";
                    return (before + " " + after).Trim();
                }
            }

            // 尝试移除 tool_calls\n[...]
            startIndex = response.IndexOf("tool_calls");
            if (startIndex != -1)
            {
                int jsonStart = response.IndexOf('[', startIndex);
                if (jsonStart != -1)
                {
                    int bracketCount = 0;
                    int jsonEnd = -1;
                    for (int i = jsonStart; i < response.Length; i++)
                    {
                        if (response[i] == '[') bracketCount++;
                        else if (response[i] == ']') bracketCount--;
                        if (bracketCount == 0)
                        {
                            jsonEnd = i;
                            break;
                        }
                    }
                    if (jsonEnd != -1)
                    {
                        string before = response.Substring(0, startIndex).Trim();
                        string after = (jsonEnd + 1 < response.Length) ? response.Substring(jsonEnd + 1).Trim() : "";
                        return (before + " " + after).Trim();
                    }
                }
            }

            return response;
        }

        // Prompt Engineering模式下的工具调用类
        private class PromptToolCall
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string ArgumentsJson { get; set; }  // 存储JSON字符串而不是JsonElement
        }

        // 获取工具组选择器（第一阶段：让模型选择需要的工具组）
        private List<object> GetToolGroupSelector()
        {
            return new List<object>
            {
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "select_tool_groups",
                        description = "根据用户需求选择需要使用的工具组。必须先调用此工具选择工具组，然后才能使用具体工具。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                groups = new
                                {
                                    type = "array",
                                    items = new
                                    {
                                        type = "string",
                                        @enum = _nativeToolGroups.Keys.ToArray()
                                    },
                                    description = $"选择需要的工具组ID列表。可选值：\n" + string.Join("\n", _nativeToolGroups.Select(g => $"- {g.Key}: {g.Value.Description}"))
                                }
                            },
                            required = new[] { "groups" }
                        }
                    }
                }
            };
        }

        // 根据选中的工具组获取具体工具定义
        private List<object> GetToolsByGroups(List<string> groupIds)
        {
            var tools = new List<object>();
            var allTools = GetMcpTools();
            var selectedToolNames = new HashSet<string>();

            // 收集所有选中组的工具名称
            foreach (var groupId in groupIds)
            {
                if (_nativeToolGroups.TryGetValue(groupId, out var groupInfo))
                {
                    foreach (var toolName in groupInfo.Tools)
                    {
                        selectedToolNames.Add(toolName);
                    }
                }
            }

            // 从完整工具列表中筛选
            foreach (var tool in allTools)
            {
                try
                {
                    var json = JsonSerializer.Serialize(tool);
                    using var doc = JsonDocument.Parse(json);
                    var funcName = doc.RootElement.GetProperty("function").GetProperty("name").GetString();
                    if (selectedToolNames.Contains(funcName))
                    {
                        tools.Add(tool);
                    }
                }
                catch { }
            }

            return tools;
        }

        // 根据用户输入智能预选工具组（减少第一阶段的必要性）
        private List<string> PreSelectToolGroups(string userInput)
        {
            var selected = new List<string>();
            string inputLower = userInput.ToLower();

            // 关键词映射
            var keywordMap = new Dictionary<string, string[]>
            {
                ["cell_rw"] = new[] { "写入", "输入", "设置值", "读取", "获取", "单元格", "公式", "清除", "复制", "范围", "查找", "替换", "统计", "最后", "区域" },
                ["format"] = new[] { "格式", "颜色", "字体", "背景", "加粗", "斜体", "边框", "合并", "对齐", "居中", "换行", "条件格式" },
                ["row_col"] = new[] { "行高", "列宽", "插入行", "插入列", "删除行", "删除列", "隐藏", "显示" },
                ["sheet"] = new[] { "工作表", "表名", "创建表", "新建表", "重命名", "删除表", "复制表", "冻结", "sheet" },
                ["workbook"] = new[] { "工作簿", "文件", "新建", "打开", "保存", "关闭" },
                ["data"] = new[] { "排序", "筛选", "去重", "验证", "表格", "图表", "chart", "折线", "柱形", "饼图", "曲线", "柱状", "散点", "面积", "雷达", "生成图", "创建图", "画图", "可视化", "分析" },
                ["named"] = new[] { "命名区域", "命名范围" },
                ["link"] = new[] { "批注", "注释", "超链接", "链接", "跳转" }
            };

            foreach (var kv in keywordMap)
            {
                foreach (var keyword in kv.Value)
                {
                    if (inputLower.Contains(keyword))
                    {
                        if (!selected.Contains(kv.Key))
                        {
                            selected.Add(kv.Key);
                        }
                        break;
                    }
                }
            }

            // 默认包含单元格读写（最常用）
            if (selected.Count == 0)
            {
                selected.Add("cell_rw");
            }

            return selected;
        }

        // 获取MCP工具定义（带缓存优化）
        private List<object> GetMcpTools()
        {
            // 如果已缓存，直接返回
            if (_cachedMcpTools != null)
            {
                return _cachedMcpTools;
            }

            // 首次调用时创建并缓存
            _cachedMcpTools = new List<object>
            {
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "create_workbook",
                        description = "创建一个新的Excel工作簿文件",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（包含.xlsx扩展名）" },
                                sheetName = new { type = "string", description = "初始工作表名称，默认为Sheet1" }
                            },
                            required = new[] { "fileName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "open_workbook",
                        description = "打开一个已存在的Excel工作簿文件",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "要打开的工作簿文件名" }
                            },
                            required = new[] { "fileName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_cell_value",
                        description = "设置Excel工作表中指定单元格的值。如果未指定工作簿或工作表名称，将使用当前活跃的工作簿和工作表。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" },
                                sheetName = new { type = "string", description = "工作表名称（可选，默认使用当前活跃工作表）" },
                                row = new { type = "integer", description = "行号（从1开始）" },
                                column = new { type = "integer", description = "列号（从1开始）" },
                                value = new { type = "string", description = "要设置的值" }
                            },
                            required = new[] { "row", "column", "value" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_cell_value",
                        description = "获取Excel工作表中指定单元格的值。如果未指定工作簿或工作表名称，将使用当前活跃的工作簿和工作表。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" },
                                sheetName = new { type = "string", description = "工作表名称（可选，默认使用当前活跃工作表）" },
                                row = new { type = "integer", description = "行号（从1开始）" },
                                column = new { type = "integer", description = "列号（从1开始）" }
                            },
                            required = new[] { "row", "column" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "save_workbook",
                        description = "保存Excel工作簿。如果未指定工作簿名称，将保存当前活跃的工作簿。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_worksheet_names",
                        description = "获取工作簿中所有工作表的名称列表。如果未指定工作簿名称，将使用当前活跃的工作簿。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "close_workbook",
                        description = "关闭已打开的Excel工作簿（自动保存）。如果未指定工作簿名称，将关闭当前活跃的工作簿。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "save_workbook_as",
                        description = "将工作簿另存为新文件",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "当前工作簿文件名" },
                                newFileName = new { type = "string", description = "新文件名" }
                            },
                            required = new[] { "fileName", "newFileName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "create_worksheet",
                        description = "在工作簿中创建新的工作表。如果未指定工作簿名称，将在当前活跃工作簿中创建。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" },
                                sheetName = new { type = "string", description = "新工作表的名称" }
                            },
                            required = new[] { "sheetName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "rename_worksheet",
                        description = "重命名工作表。如果未指定工作簿名称，将在当前活跃工作簿中操作。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" },
                                oldSheetName = new { type = "string", description = "原工作表名称" },
                                newSheetName = new { type = "string", description = "新工作表名称" }
                            },
                            required = new[] { "oldSheetName", "newSheetName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "delete_worksheet",
                        description = "删除工作表。如果未指定工作簿名称，将在当前活跃工作簿中操作。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" },
                                sheetName = new { type = "string", description = "要删除的工作表名称" }
                            },
                            required = new[] { "sheetName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_range_values",
                        description = "设置单元格区域的值（批量设置）。如果未指定工作簿或工作表名称，将使用当前活跃的工作簿和工作表。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" },
                                sheetName = new { type = "string", description = "工作表名称（可选，默认使用当前活跃工作表）" },
                                rangeAddress = new { type = "string", description = "单元格区域地址，如'A1:C3'" },
                                data = new { type = "string", description = "JSON格式的二维数组数据，如'[[1,2,3],[4,5,6]]'" }
                            },
                            required = new[] { "rangeAddress", "data" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_range_values",
                        description = "获取单元格区域的值（批量获取）。如果未指定工作簿或工作表名称，将使用当前活跃的工作簿和工作表。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" },
                                sheetName = new { type = "string", description = "工作表名称（可选，默认使用当前活跃工作表）" },
                                rangeAddress = new { type = "string", description = "单元格区域地址，如'A1:C3'" }
                            },
                            required = new[] { "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_formula",
                        description = "设置单元格的公式。如果未指定工作簿或工作表名称，将使用当前活跃的工作簿和工作表。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" },
                                sheetName = new { type = "string", description = "工作表名称（可选，默认使用当前活跃工作表）" },
                                cellAddress = new { type = "string", description = "单元格地址，如'A1'" },
                                formula = new { type = "string", description = "Excel公式，如'=SUM(A1:A10)'" }
                            },
                            required = new[] { "cellAddress", "formula" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_formula",
                        description = "获取单元格的公式。如果未指定工作簿或工作表名称，将使用当前活跃的工作簿和工作表。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" },
                                sheetName = new { type = "string", description = "工作表名称（可选，默认使用当前活跃工作表）" },
                                cellAddress = new { type = "string", description = "单元格地址，如'A1'" }
                            },
                            required = new[] { "cellAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_excel_files",
                        description = "获取excel_files目录下所有Excel文件列表",
                        parameters = new
                        {
                            type = "object",
                            properties = new { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "delete_excel_file",
                        description = "删除Excel文件（文件必须已关闭）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "要删除的文件名" }
                            },
                            required = new[] { "fileName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_current_excel_info",
                        description = "获取当前Excel应用程序中打开的工作簿和活跃工作表信息",
                        parameters = new
                        {
                            type = "object",
                            properties = new { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_cell_format",
                        description = "设置单元格或区域的格式（字体颜色、背景色、对齐方式等）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格区域地址，如'A1'或'A1:C3'" },
                                fontColor = new { type = "string", description = "字体颜色（可选），如'红色'、'#FF0000'" },
                                backgroundColor = new { type = "string", description = "背景色（可选），如'黄色'、'#FFFF00'" },
                                fontSize = new { type = "integer", description = "字号（可选），如12" },
                                bold = new { type = "boolean", description = "是否加粗（可选）" },
                                italic = new { type = "boolean", description = "是否斜体（可选）" },
                                horizontalAlignment = new { type = "string", description = "水平对齐（可选）：left/center/right" },
                                verticalAlignment = new { type = "string", description = "垂直对齐（可选）：top/center/bottom" }
                            },
                            required = new[] { "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_border",
                        description = "设置单元格或区域的边框",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格区域地址，如'A1:C3'" },
                                borderType = new { type = "string", description = "边框类型：all(全部)/outline(外框)/horizontal(横线)/vertical(竖线)" },
                                lineStyle = new { type = "string", description = "线型（可选）：continuous(实线)/dash(虚线)/dot(点线)" }
                            },
                            required = new[] { "rangeAddress", "borderType" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "merge_cells",
                        description = "合并单元格",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "要合并的单元格区域，如'A1:C3'" }
                            },
                            required = new[] { "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "unmerge_cells",
                        description = "取消合并单元格",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "要取消合并的单元格区域，如'A1:C3'" }
                            },
                            required = new[] { "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_row_height",
                        description = "设置行高",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rowNumber = new { type = "integer", description = "行号" },
                                height = new { type = "number", description = "行高（磅）" }
                            },
                            required = new[] { "rowNumber", "height" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_column_width",
                        description = "设置列宽",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                columnNumber = new { type = "integer", description = "列号（A=1, B=2...）" },
                                width = new { type = "number", description = "列宽（字符）" }
                            },
                            required = new[] { "columnNumber", "width" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "insert_rows",
                        description = "在指定位置插入行",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rowIndex = new { type = "integer", description = "插入位置的行号" },
                                count = new { type = "integer", description = "插入的行数（默认1）" }
                            },
                            required = new[] { "rowIndex" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "insert_columns",
                        description = "在指定位置插入列",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                columnIndex = new { type = "integer", description = "插入位置的列号" },
                                count = new { type = "integer", description = "插入的列数（默认1）" }
                            },
                            required = new[] { "columnIndex" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "delete_rows",
                        description = "删除指定的行",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rowIndex = new { type = "integer", description = "起始行号" },
                                count = new { type = "integer", description = "删除的行数（默认1）" }
                            },
                            required = new[] { "rowIndex" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "delete_columns",
                        description = "删除指定的列",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                columnIndex = new { type = "integer", description = "起始列号" },
                                count = new { type = "integer", description = "删除的列数（默认1）" }
                            },
                            required = new[] { "columnIndex" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "copy_worksheet",
                        description = "复制工作表",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sourceSheetName = new { type = "string", description = "源工作表名称" },
                                targetSheetName = new { type = "string", description = "目标工作表名称" }
                            },
                            required = new[] { "sourceSheetName", "targetSheetName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "copy_range",
                        description = "复制单元格范围",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                sourceRange = new { type = "string", description = "源范围地址（如'A1:C3'）" },
                                targetRange = new { type = "string", description = "目标范围地址（如'E1'）" }
                            },
                            required = new[] { "sourceRange", "targetRange" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "clear_range",
                        description = "清除范围内容或格式",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格范围地址" },
                                clearType = new { type = "string", description = "清除类型：all(全部)/contents(内容)/formats(格式)" }
                            },
                            required = new[] { "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_workbook_metadata",
                        description = "获取工作簿元数据信息",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                includeRanges = new { type = "boolean", description = "是否包含范围信息（默认false）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_data_validation",
                        description = "设置数据验证规则（下拉列表、数值限制等）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格范围地址" },
                                validationType = new { type = "string", description = "验证类型：whole/decimal/list/date/time/textlength/custom" },
                                operatorType = new { type = "string", description = "操作符：between/equal/greater/less等" },
                                formula1 = new { type = "string", description = "公式1或列表值" },
                                formula2 = new { type = "string", description = "公式2（范围时使用）" },
                                inputMessage = new { type = "string", description = "输入提示" },
                                errorMessage = new { type = "string", description = "错误提示" }
                            },
                            required = new[] { "rangeAddress", "validationType" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_validation_rules",
                        description = "获取单元格范围的数据验证规则",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格范围地址（可选）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_number_format",
                        description = "设置单元格数字格式",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格范围地址" },
                                formatCode = new { type = "string", description = "格式代码（如'0.00','#,##0','yyyy-mm-dd'）" }
                            },
                            required = new[] { "rangeAddress", "formatCode" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "apply_conditional_formatting",
                        description = "应用条件格式（色阶、数据条、图标集等）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格范围地址" },
                                ruleType = new { type = "string", description = "规则类型：cellvalue/colorscale/databar/iconset/expression" },
                                formula1 = new { type = "string", description = "公式或条件值" },
                                formula2 = new { type = "string", description = "公式2（可选）" },
                                color1 = new { type = "string", description = "颜色1（可选）" },
                                color2 = new { type = "string", description = "颜色2（可选）" },
                                color3 = new { type = "string", description = "颜色3（可选）" }
                            },
                            required = new[] { "rangeAddress", "ruleType" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "create_chart",
                        description = "创建图表（折线图、柱状图、饼图等）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                chartType = new { type = "string", description = "图表类型：line/bar/column/pie/scatter/area/radar（默认column）" },
                                dataRange = new { type = "string", description = "数据源范围（如'A1:D10'），也可用rangeAddress" },
                                chartPosition = new { type = "string", description = "图表位置（如'F1'，可选，默认在数据右侧）" },
                                title = new { type = "string", description = "图表标题（可选）" },
                                width = new { type = "integer", description = "图表宽度（默认400）" },
                                height = new { type = "integer", description = "图表高度（默认300）" }
                            },
                            required = new[] { "dataRange" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "create_table",
                        description = "创建Excel原生表格(ListObject)",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "表格数据范围" },
                                tableName = new { type = "string", description = "表格名称" },
                                hasHeaders = new { type = "boolean", description = "是否包含标题行（默认true）" },
                                tableStyle = new { type = "string", description = "表格样式（默认TableStyleMedium2）" }
                            },
                            required = new[] { "rangeAddress", "tableName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_table_names",
                        description = "获取工作表中所有表格名称",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "validate_formula",
                        description = "验证Excel公式语法是否正确",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                formula = new { type = "string", description = "要验证的公式（如'=SUM(A1:A10)'）" }
                            },
                            required = new[] { "formula" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "find_value",
                        description = "在工作表中查找指定值的所有位置",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                searchValue = new { type = "string", description = "要查找的值" },
                                matchCase = new { type = "boolean", description = "是否区分大小写（默认false）" }
                            },
                            required = new[] { "searchValue" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "find_and_replace",
                        description = "在工作表中查找并替换所有匹配的值",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                findWhat = new { type = "string", description = "要查找的值" },
                                replaceWith = new { type = "string", description = "替换后的值" },
                                matchCase = new { type = "boolean", description = "是否区分大小写（默认false）" }
                            },
                            required = new[] { "findWhat", "replaceWith" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "freeze_panes",
                        description = "冻结窗格（冻结指定行和列之前的部分）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                row = new { type = "integer", description = "冻结行号（在此行之前的行将被冻结）" },
                                column = new { type = "integer", description = "冻结列号（在此列之前的列将被冻结）" }
                            },
                            required = new[] { "row", "column" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "unfreeze_panes",
                        description = "取消冻结窗格",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "autofit_columns",
                        description = "自动调整列宽以适应内容",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "要调整的范围地址（如'A:A'或'A1:C10'）" }
                            },
                            required = new[] { "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "autofit_rows",
                        description = "自动调整行高以适应内容",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "要调整的范围地址（如'1:1'或'A1:C10'）" }
                            },
                            required = new[] { "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_column_visible",
                        description = "设置列的可见性（隐藏或显示列）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                columnIndex = new { type = "integer", description = "列号（A=1, B=2...）" },
                                visible = new { type = "boolean", description = "是否可见（true显示，false隐藏）" }
                            },
                            required = new[] { "columnIndex", "visible" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_row_visible",
                        description = "设置行的可见性（隐藏或显示行）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rowIndex = new { type = "integer", description = "行号" },
                                visible = new { type = "boolean", description = "是否可见（true显示，false隐藏）" }
                            },
                            required = new[] { "rowIndex", "visible" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "add_comment",
                        description = "为单元格添加批注",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                cellAddress = new { type = "string", description = "单元格地址（如'A1'）" },
                                commentText = new { type = "string", description = "批注文本" }
                            },
                            required = new[] { "cellAddress", "commentText" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "delete_comment",
                        description = "删除单元格的批注",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                cellAddress = new { type = "string", description = "单元格地址（如'A1'）" }
                            },
                            required = new[] { "cellAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_comment",
                        description = "获取单元格的批注内容",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                cellAddress = new { type = "string", description = "单元格地址（如'A1'）" }
                            },
                            required = new[] { "cellAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "add_hyperlink",
                        description = "为单元格添加超链接对象（非公式方式）。适用于外部链接，如网址（会用浏览器打开）、本地文件路径、网络文件路径等。不适用于工作簿内部跳转。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                cellAddress = new { type = "string", description = "单元格地址（如'A1'）" },
                                url = new { type = "string", description = "链接地址：网址（如'https://www.baidu.com'）或本地/网络文件路径（如'C:\\Documents\\file.xlsx'）" },
                                displayText = new { type = "string", description = "显示文本（可选）" }
                            },
                            required = new[] { "cellAddress", "url" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_hyperlink_formula",
                        description = "使用HYPERLINK公式为单元格设置超链接。适用于工作簿内部跳转（如跳转到其他工作表的某个单元格），此类链接在Excel内打开，不会打开浏览器。",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                cellAddress = new { type = "string", description = "要设置公式的单元格地址（如'A1'）" },
                                targetLocation = new { type = "string", description = "目标位置，格式为'工作表名!单元格地址'，如'Sheet2!A1'、'销售数据!B5'" },
                                displayText = new { type = "string", description = "显示文本，如'跳转到Sheet2'、'查看详情'" }
                            },
                            required = new[] { "cellAddress", "targetLocation", "displayText" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "delete_hyperlink",
                        description = "删除单元格的超链接",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                cellAddress = new { type = "string", description = "单元格地址（如'A1'）" }
                            },
                            required = new[] { "cellAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_used_range",
                        description = "获取工作表中已使用的单元格范围",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_range_statistics",
                        description = "获取单元格范围的统计信息（总和、平均值、最大值、最小值等）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格范围地址（如'A1:A10'）" }
                            },
                            required = new[] { "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_last_row",
                        description = "获取指定列中最后一个有数据的行号",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                columnIndex = new { type = "integer", description = "列号（默认为1，即A列）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_last_column",
                        description = "获取指定行中最后一个有数据的列号",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rowIndex = new { type = "integer", description = "行号（默认为1）" }
                            },
                            required = new string[] { }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "sort_range",
                        description = "对单元格范围进行排序",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "要排序的范围地址（如'A1:C10'）" },
                                sortColumnIndex = new { type = "integer", description = "排序依据的列索引（相对于范围的列，1表示第一列）" },
                                ascending = new { type = "boolean", description = "是否升序排列（true升序，false降序，默认true）" }
                            },
                            required = new[] { "rangeAddress", "sortColumnIndex" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_auto_filter",
                        description = "为范围设置自动筛选",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "要筛选的范围地址（如'A1:C10'）" },
                                columnIndex = new { type = "integer", description = "筛选列索引（可选，0表示不筛选）" },
                                criteria = new { type = "string", description = "筛选条件（可选）" }
                            },
                            required = new[] { "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "remove_duplicates",
                        description = "删除范围中的重复行",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "要处理的范围地址（如'A1:C10'）" },
                                columnIndices = new { type = "string", description = "用于判断重复的列索引数组（JSON格式，如'[1,2]'表示第1和第2列）" }
                            },
                            required = new[] { "rangeAddress", "columnIndices" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "move_worksheet",
                        description = "移动工作表到指定位置",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "要移动的工作表名称" },
                                position = new { type = "integer", description = "目标位置（1表示第一个位置）" }
                            },
                            required = new[] { "sheetName", "position" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_worksheet_visible",
                        description = "设置工作表的可见性（隐藏或显示工作表）",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称" },
                                visible = new { type = "boolean", description = "是否可见（true显示，false隐藏）" }
                            },
                            required = new[] { "sheetName", "visible" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_worksheet_index",
                        description = "获取工作表在工作簿中的位置索引",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称" }
                            },
                            required = new[] { "sheetName" }
                        }
                    }
                },
                // 命名区域工具
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "create_named_range",
                        description = "创建命名区域，使公式更易读。例如将A2:A100命名为'销售额'，之后可以使用=SUM(销售额)代替=SUM(A2:A100)",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeName = new { type = "string", description = "命名区域的名称，如'销售额'、'成本'" },
                                rangeAddress = new { type = "string", description = "区域地址，如'A2:A100'、'B1:D10'" }
                            },
                            required = new[] { "rangeName", "rangeAddress" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "delete_named_range",
                        description = "删除命名区域",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                rangeName = new { type = "string", description = "要删除的命名区域的名称" }
                            },
                            required = new[] { "rangeName" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_named_ranges",
                        description = "获取工作簿中所有命名区域的列表",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" }
                            }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_named_range_address",
                        description = "获取命名区域的引用地址",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                rangeName = new { type = "string", description = "命名区域的名称" }
                            },
                            required = new[] { "rangeName" }
                        }
                    }
                },
                // 单元格格式增强工具
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_cell_text_wrap",
                        description = "设置单元格文本自动换行，适用于长文本内容",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格或区域地址，如'A1'或'A1:C10'" },
                                wrap = new { type = "boolean", description = "true=自动换行，false=不换行" }
                            },
                            required = new[] { "rangeAddress", "wrap" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_cell_indent",
                        description = "设置单元格的缩进级别，用于层级显示",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格或区域地址" },
                                indentLevel = new { type = "integer", description = "缩进级别（0-15）" }
                            },
                            required = new[] { "rangeAddress", "indentLevel" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_cell_orientation",
                        description = "设置单元格文本的旋转角度，常用于表头设计",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格或区域地址" },
                                degrees = new { type = "integer", description = "旋转角度（-90到90），正数逆时针，负数顺时针" }
                            },
                            required = new[] { "rangeAddress", "degrees" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "set_cell_shrink_to_fit",
                        description = "设置单元格缩小字体以适应单元格宽度",
                        parameters = new
                        {
                            type = "object",
                            properties = new
                            {
                                fileName = new { type = "string", description = "工作簿文件名（可选）" },
                                sheetName = new { type = "string", description = "工作表名称（可选）" },
                                rangeAddress = new { type = "string", description = "单元格或区域地址" },
                                shrink = new { type = "boolean", description = "true=缩小字体填充，false=不缩小" }
                            },
                            required = new[] { "rangeAddress", "shrink" }
                        }
                    }
                },
                new
                {
                    type = "function",
                    function = new
                    {
                        name = "get_current_selection",
                        description = "获取当前选中的单元格或区域的信息（地址、行号、列号、值等）",
                        parameters = new
                        {
                            type = "object",
                            properties = new { }
                        }
                    }
                }
            };

            return _cachedMcpTools;
        }

        // 工具名称规范化：将模型可能输出的变体名称映射到正确的工具名
        private string NormalizeToolName(string toolName)
        {
            if (string.IsNullOrEmpty(toolName)) return toolName;

            // 转换为小写进行匹配
            var lowerName = toolName.ToLower().Trim();

            // 常见变体映射
            var aliases = new Dictionary<string, string>
            {
                // 复数形式 -> 单数形式
                { "create_worksheets", "create_worksheet" },
                { "delete_worksheets", "delete_worksheet" },
                { "rename_worksheets", "rename_worksheet" },
                { "copy_worksheets", "copy_worksheet" },
                { "move_worksheets", "move_worksheet" },
                { "create_workbooks", "create_workbook" },
                { "open_workbooks", "open_workbook" },
                { "save_workbooks", "save_workbook" },
                { "close_workbooks", "close_workbook" },
                { "insert_row", "insert_rows" },
                { "insert_column", "insert_columns" },
                { "delete_row", "delete_rows" },
                { "delete_column", "delete_columns" },
                // 其他可能的变体
                { "set_value", "set_cell_value" },
                { "get_value", "get_cell_value" },
                { "setcellvalue", "set_cell_value" },
                { "getcellvalue", "get_cell_value" },
                { "createworksheet", "create_worksheet" },
                { "createworkbook", "create_workbook" },
                { "renameworksheet", "rename_worksheet" },
                { "rename_sheet", "rename_worksheet" },
                { "renamesheet", "rename_worksheet" },
                { "sheet_rename", "rename_worksheet" },
                // 更多工作表操作变体
                { "deleteworksheet", "delete_worksheet" },
                { "delete_sheet", "delete_worksheet" },
                { "copyworksheet", "copy_worksheet" },
                { "copy_sheet", "copy_worksheet" },
                { "moveworksheet", "move_worksheet" },
                { "move_sheet", "move_worksheet" },
            };

            if (aliases.TryGetValue(lowerName, out var normalized))
            {
                System.Diagnostics.Debug.WriteLine($"工具名称规范化: {toolName} -> {normalized}");
                return normalized;
            }

            return toolName;
        }

        // 执行MCP工具调用（确保在UI线程上执行Excel COM操作）
        private string ExecuteMcpTool(string toolName, JsonElement arguments)
        {
            // 如果不在UI线程上，需要切换到UI线程执行
            if (this.InvokeRequired)
            {
                string result = null;
                this.Invoke(new Action(() =>
                {
                    result = ExecuteMcpToolInternal(toolName, arguments);
                }));
                return result;
            }
            return ExecuteMcpToolInternal(toolName, arguments);
        }

        // 实际执行MCP工具调用的内部方法
        private string ExecuteMcpToolInternal(string toolName, JsonElement arguments)
        {
            try
            {
                // 工具名称规范化：处理模型可能输出的变体名称
                toolName = NormalizeToolName(toolName);

                // 辅助方法：获取文件名（如果未提供，使用活跃工作簿）
                string GetFileName()
                {
                    if (arguments.TryGetProperty("fileName", out var fileNameProp))
                    {
                        var fn = fileNameProp.GetString();
                        if (!string.IsNullOrEmpty(fn)) return fn;
                    }
                    if (string.IsNullOrEmpty(_activeWorkbook))
                        throw new Exception("未指定工作簿名称且没有活跃的工作簿");
                    return _activeWorkbook;
                }

                // 辅助方法：获取工作表名（如果未提供，使用活跃工作表）
                string GetSheetName()
                {
                    if (arguments.TryGetProperty("sheetName", out var sheetNameProp))
                    {
                        var sn = sheetNameProp.GetString();
                        if (!string.IsNullOrEmpty(sn)) return sn;
                    }
                    if (string.IsNullOrEmpty(_activeWorksheet))
                        throw new Exception("未指定工作表名称且没有活跃的工作表");
                    return _activeWorksheet;
                }

                // 辅助方法：获取当前Excel工作簿
                Microsoft.Office.Interop.Excel.Workbook GetCurrentWorkbook(string fileName = null)
                {
                    if (ThisAddIn.app == null)
                        throw new Exception("Excel应用程序未初始化");

                    var targetFileName = fileName ?? _activeWorkbook;
                    if (string.IsNullOrEmpty(targetFileName))
                        throw new Exception("未指定工作簿且没有活跃工作簿");

                    // 查找指定的工作簿
                    foreach (Microsoft.Office.Interop.Excel.Workbook wb in ThisAddIn.app.Workbooks)
                    {
                        if (wb.Name == targetFileName)
                            return wb;
                    }

                    throw new Exception($"未找到工作簿: {targetFileName}");
                }

                // 辅助方法：获取工作表
                Microsoft.Office.Interop.Excel.Worksheet GetWorksheet(string fileName = null, string sheetName = null)
                {
                    var workbook = GetCurrentWorkbook(fileName);
                    var targetSheetName = sheetName ?? _activeWorksheet;

                    if (string.IsNullOrEmpty(targetSheetName))
                        throw new Exception("未指定工作表且没有活跃工作表");

                    foreach (Microsoft.Office.Interop.Excel.Worksheet ws in workbook.Worksheets)
                    {
                        if (ws.Name == targetSheetName)
                            return ws;
                    }

                    throw new Exception($"未找到工作表: {targetSheetName}");
                }

                switch (toolName)
                {
                    case "create_workbook":
                        {
                            var fileName = arguments.GetProperty("fileName").GetString();
                            var sheetName = arguments.TryGetProperty("sheetName", out var sheet) ? sheet.GetString() : "Sheet1";

                            // 使用ExcelMcp创建独立文件
                            var result = _excelMcp.CreateWorkbook(fileName, sheetName);

                            // 注意：这里创建的是独立文件，不会在当前Excel中打开
                            return $"成功创建工作簿文件: {result}（保存在excel_files目录）";
                        }

                    case "open_workbook":
                        {
                            var fileName = arguments.GetProperty("fileName").GetString();

                            // 使用Excel应用程序打开文件
                            var filePath = System.IO.Path.Combine(_excelMcp.GetType().GetField("_excelFilesPath",
                                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                                ?.GetValue(_excelMcp)?.ToString() ?? "./excel_files", fileName);

                            if (!System.IO.File.Exists(filePath))
                                throw new Exception($"文件不存在: {filePath}");

                            var wb = ThisAddIn.app.Workbooks.Open(filePath);
                            _activeWorkbook = wb.Name;

                            if (wb.Worksheets.Count > 0)
                            {
                                Microsoft.Office.Interop.Excel.Worksheet ws = wb.Worksheets[1];
                                _activeWorksheet = ws.Name;
                            }

                            return $"成功打开工作簿: {fileName}，当前活跃工作簿已设置为 {_activeWorkbook}" +
                                   (!string.IsNullOrEmpty(_activeWorksheet) ? $"，活跃工作表为 {_activeWorksheet}" : "");
                        }

                    case "close_workbook":
                        {
                            var fileName = GetFileName();
                            var workbook = GetCurrentWorkbook(fileName);
                            workbook.Close(true);

                            if (_activeWorkbook == fileName)
                            {
                                _activeWorkbook = string.Empty;
                                _activeWorksheet = string.Empty;
                            }

                            return $"成功关闭工作簿: {fileName}";
                        }

                    case "save_workbook":
                        {
                            var fileName = GetFileName();
                            var workbook = GetCurrentWorkbook(fileName);
                            workbook.Save();
                            return $"成功保存工作簿: {fileName}";
                        }

                    case "save_workbook_as":
                        {
                            var fileName = arguments.GetProperty("fileName").GetString();
                            var newFileName = arguments.GetProperty("newFileName").GetString();
                            var workbook = GetCurrentWorkbook(fileName);

                            var newFilePath = System.IO.Path.Combine(
                                System.IO.Path.GetDirectoryName(workbook.FullName), newFileName);

                            workbook.SaveAs(newFilePath);

                            if (_activeWorkbook == fileName)
                                _activeWorkbook = newFileName;

                            return $"成功将工作簿 {fileName} 另存为 {newFileName}";
                        }

                    case "activate_worksheet":
                        {
                            var fileName = GetFileName();
                            var sheetName = arguments.GetProperty("sheetName").GetString();
                            var workbook = GetCurrentWorkbook(fileName);

                            // 查找并激活指定工作表
                            Microsoft.Office.Interop.Excel.Worksheet targetSheet = null;
                            foreach (Microsoft.Office.Interop.Excel.Worksheet ws in workbook.Worksheets)
                            {
                                if (ws.Name == sheetName)
                                {
                                    targetSheet = ws;
                                    break;
                                }
                            }

                            if (targetSheet == null)
                                throw new Exception($"未找到工作表: {sheetName}");

                            targetSheet.Activate();
                            _activeWorksheet = sheetName;

                            return $"成功激活工作表: {sheetName}，后续操作将在此表上执行";
                        }

                    case "create_worksheet":
                        {
                            var fileName = GetFileName();
                            var sheetName = arguments.GetProperty("sheetName").GetString();
                            var workbook = GetCurrentWorkbook(fileName);

                            Microsoft.Office.Interop.Excel.Worksheet newSheet = workbook.Worksheets.Add();
                            newSheet.Name = sheetName;

                            _activeWorksheet = sheetName;

                            return $"成功创建工作表: {sheetName}，当前活跃工作表已设置为 {sheetName}";
                        }

                    case "rename_worksheet":
                        {
                            var fileName = GetFileName();
                            var oldSheetName = arguments.GetProperty("oldSheetName").GetString();
                            var newSheetName = arguments.GetProperty("newSheetName").GetString();

                            var worksheet = GetWorksheet(fileName, oldSheetName);
                            worksheet.Name = newSheetName;

                            if (_activeWorksheet == oldSheetName)
                                _activeWorksheet = newSheetName;

                            return $"成功将工作表 {oldSheetName} 重命名为 {newSheetName}";
                        }

                    case "delete_worksheet":
                        {
                            var fileName = GetFileName();
                            var sheetName = arguments.GetProperty("sheetName").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            worksheet.Delete();

                            if (_activeWorksheet == sheetName)
                                _activeWorksheet = string.Empty;

                            return $"成功删除工作表: {sheetName}";
                        }

                    case "get_worksheet_names":
                        {
                            var fileName = GetFileName();
                            var workbook = GetCurrentWorkbook(fileName);

                            var names = new List<string>();
                            foreach (Microsoft.Office.Interop.Excel.Worksheet ws in workbook.Worksheets)
                            {
                                names.Add(ws.Name);
                            }

                            return $"工作表列表: {string.Join(", ", names)}";
                        }

                    case "set_cell_value":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var row = arguments.GetProperty("row").GetInt32();
                            var column = arguments.GetProperty("column").GetInt32();
                            
                            // 支持value为字符串或数字类型
                            object value;
                            var valueProp = arguments.GetProperty("value");
                            if (valueProp.ValueKind == JsonValueKind.Number)
                            {
                                value = valueProp.GetDouble();
                            }
                            else if (valueProp.ValueKind == JsonValueKind.String)
                            {
                                value = valueProp.GetString();
                            }
                            else
                            {
                                value = valueProp.ToString();
                            }

                            var worksheet = GetWorksheet(fileName, sheetName);
                            Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[row, column];
                            cell.Value = value;

                            return $"成功设置单元格 ({row},{column}) 的值为: {value}";
                        }

                    case "get_cell_value":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var row = arguments.GetProperty("row").GetInt32();
                            var column = arguments.GetProperty("column").GetInt32();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            Microsoft.Office.Interop.Excel.Range cell = worksheet.Cells[row, column];
                            var value = cell.Value?.ToString() ?? "";

                            return $"单元格 ({row},{column}) 的值为: {value}";
                        }

                    case "set_range_values":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var dataProp = arguments.GetProperty("data");

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            // 支持data为字符串（JSON）或直接数组
                            List<List<object>> dataList;
                            if (dataProp.ValueKind == JsonValueKind.String)
                            {
                                // data是JSON字符串
                                var dataJson = dataProp.GetString();
                                dataList = JsonSerializer.Deserialize<List<List<object>>>(dataJson);
                            }
                            else if (dataProp.ValueKind == JsonValueKind.Array)
                            {
                                // data是直接的数组
                                dataList = new List<List<object>>();
                                foreach (var row in dataProp.EnumerateArray())
                                {
                                    var rowList = new List<object>();
                                    if (row.ValueKind == JsonValueKind.Array)
                                    {
                                        foreach (var cell in row.EnumerateArray())
                                        {
                                            if (cell.ValueKind == JsonValueKind.Number)
                                                rowList.Add(cell.GetDouble());
                                            else if (cell.ValueKind == JsonValueKind.String)
                                                rowList.Add(cell.GetString());
                                            else
                                                rowList.Add(cell.ToString());
                                        }
                                    }
                                    else
                                    {
                                        // 单个值，作为一列
                                        if (row.ValueKind == JsonValueKind.Number)
                                            rowList.Add(row.GetDouble());
                                        else if (row.ValueKind == JsonValueKind.String)
                                            rowList.Add(row.GetString());
                                        else
                                            rowList.Add(row.ToString());
                                    }
                                    dataList.Add(rowList);
                                }
                            }
                            else
                            {
                                throw new Exception("data参数格式不正确，应为JSON字符串或数组");
                            }

                            var rows = dataList.Count;
                            var cols = dataList[0].Count;
                            var data = new object[rows, cols];
                            for (int i = 0; i < rows; i++)
                            {
                                for (int j = 0; j < cols; j++)
                                {
                                    data[i, j] = dataList[i][j];
                                }
                            }

                            range.Value = data;
                            return $"成功设置区域 {rangeAddress} 的值";
                        }

                    case "get_range_values":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];
                            var values = range.Value as object[,];

                            if (values == null)
                            {
                                return $"区域 {rangeAddress} 为空";
                            }

                            // 转换为JSON字符串
                            var result = new List<List<object>>();
                            for (int i = values.GetLowerBound(0); i <= values.GetUpperBound(0); i++)
                            {
                                var row = new List<object>();
                                for (int j = values.GetLowerBound(1); j <= values.GetUpperBound(1); j++)
                                {
                                    row.Add(values[i, j]);
                                }
                                result.Add(row);
                            }
                            var jsonResult = JsonSerializer.Serialize(result);
                            return $"区域 {rangeAddress} 的值: {jsonResult}";
                        }

                    case "set_formula":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var cellAddress = arguments.GetProperty("cellAddress").GetString();
                            var formula = arguments.GetProperty("formula").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[cellAddress];
                            range.Formula = formula;

                            return $"成功设置单元格 {cellAddress} 的公式: {formula}";
                        }

                    case "get_formula":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var cellAddress = arguments.GetProperty("cellAddress").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[cellAddress];
                            var formula = range.Formula?.ToString() ?? "";

                            return $"单元格 {cellAddress} 的公式为: {formula}";
                        }

                    case "get_excel_files":
                        {
                            var files = _excelMcp.GetExcelFiles();
                            return $"Excel文件列表: {string.Join(", ", files)}";
                        }

                    case "delete_excel_file":
                        {
                            var fileName = arguments.GetProperty("fileName").GetString();
                            _excelMcp.DeleteExcelFile(fileName);

                            if (_activeWorkbook == fileName)
                            {
                                _activeWorkbook = string.Empty;
                                _activeWorksheet = string.Empty;
                            }

                            return $"成功删除文件: {fileName}";
                        }

                    case "get_current_excel_info":
                        {
                            try
                            {
                                if (ThisAddIn.app == null)
                                    return "Excel应用程序未初始化";

                                var info = new System.Text.StringBuilder();
                                info.AppendLine("当前Excel环境信息：");

                                if (ThisAddIn.app.ActiveWorkbook != null)
                                {
                                    var wb = ThisAddIn.app.ActiveWorkbook;
                                    info.AppendLine($"- 活跃工作簿: {wb.Name}");
                                    _activeWorkbook = wb.Name;

                                    if (ThisAddIn.app.ActiveSheet != null)
                                    {
                                        Microsoft.Office.Interop.Excel.Worksheet ws = ThisAddIn.app.ActiveSheet;
                                        info.AppendLine($"- 活跃工作表: {ws.Name}");
                                        _activeWorksheet = ws.Name;

                                        info.Append("- 所有工作表: ");
                                        var sheetNames = new List<string>();
                                        foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in wb.Worksheets)
                                        {
                                            sheetNames.Add(sheet.Name);
                                        }
                                        info.AppendLine(string.Join(", ", sheetNames));
                                    }
                                }
                                else
                                {
                                    info.AppendLine("- 当前没有打开的工作簿");
                                }

                                return info.ToString();
                            }
                            catch (Exception ex)
                            {
                                return $"获取Excel信息失败: {ex.Message}";
                            }
                        }

                    case "set_cell_format":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            
                            // 获取rangeAddress，如果未提供则使用当前选中的单元格
                            string rangeAddress;
                            if (arguments.TryGetProperty("rangeAddress", out var rangeAddressProp))
                            {
                                rangeAddress = rangeAddressProp.GetString();
                            }
                            else
                            {
                                // 未提供rangeAddress，使用当前选中的单元格
                                if (ThisAddIn.app?.Selection != null)
                                {
                                    Microsoft.Office.Interop.Excel.Range selection = ThisAddIn.app.Selection;
                                    rangeAddress = selection.Address.Replace("$", "");
                                }
                                else
                                {
                                    throw new Exception("未提供rangeAddress参数，且无法获取当前选中的单元格");
                                }
                            }
                            
                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            // 字体颜色
                            if (arguments.TryGetProperty("fontColor", out var fontColorProp))
                            {
                                var color = ParseColor(fontColorProp.GetString());
                                range.Font.Color = color;
                            }

                            // 背景色
                            if (arguments.TryGetProperty("backgroundColor", out var bgColorProp))
                            {
                                var color = ParseColor(bgColorProp.GetString());
                                range.Interior.Color = color;
                            }

                            // 字号
                            if (arguments.TryGetProperty("fontSize", out var fontSizeProp))
                            {
                                range.Font.Size = fontSizeProp.GetInt32();
                            }

                            // 加粗
                            if (arguments.TryGetProperty("bold", out var boldProp))
                            {
                                range.Font.Bold = boldProp.GetBoolean();
                            }

                            // 斜体
                            if (arguments.TryGetProperty("italic", out var italicProp))
                            {
                                range.Font.Italic = italicProp.GetBoolean();
                            }

                            // 水平对齐
                            if (arguments.TryGetProperty("horizontalAlignment", out var hAlignProp))
                            {
                                var align = hAlignProp.GetString().ToLower();
                                range.HorizontalAlignment = align switch
                                {
                                    "left" => Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft,
                                    "center" => Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter,
                                    "right" => Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight,
                                    _ => Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignGeneral
                                };
                            }

                            // 垂直对齐
                            if (arguments.TryGetProperty("verticalAlignment", out var vAlignProp))
                            {
                                var align = vAlignProp.GetString().ToLower();
                                range.VerticalAlignment = align switch
                                {
                                    "top" => Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop,
                                    "center" => Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter,
                                    "bottom" => Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom,
                                    _ => Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
                                };
                            }

                            return $"成功设置区域 {rangeAddress} 的格式";
                        }

                    case "set_border":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var borderType = arguments.GetProperty("borderType").GetString().ToLower();
                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            var lineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                            if (arguments.TryGetProperty("lineStyle", out var lineStyleProp))
                            {
                                lineStyle = lineStyleProp.GetString().ToLower() switch
                                {
                                    "dash" => Microsoft.Office.Interop.Excel.XlLineStyle.xlDash,
                                    "dot" => Microsoft.Office.Interop.Excel.XlLineStyle.xlDot,
                                    _ => Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                };
                            }

                            switch (borderType)
                            {
                                case "all":
                                    range.Borders.LineStyle = lineStyle;
                                    break;
                                case "outline":
                                    range.BorderAround(lineStyle);
                                    break;
                                case "horizontal":
                                    range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = lineStyle;
                                    break;
                                case "vertical":
                                    range.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = lineStyle;
                                    break;
                            }

                            return $"成功设置区域 {rangeAddress} 的边框";
                        }

                    case "merge_cells":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            range.Merge();

                            return $"成功合并单元格 {rangeAddress}";
                        }

                    case "unmerge_cells":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            range.UnMerge();

                            return $"成功取消合并单元格 {rangeAddress}";
                        }

                    case "set_row_height":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rowNumber = arguments.GetProperty("rowNumber").GetInt32();
                            var height = arguments.GetProperty("height").GetDouble();
                            var worksheet = GetWorksheet(fileName, sheetName);

                            Microsoft.Office.Interop.Excel.Range row = worksheet.Rows[rowNumber];
                            row.RowHeight = height;

                            return $"成功设置第 {rowNumber} 行的行高为 {height}";
                        }

                    case "set_column_width":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var columnNumber = arguments.GetProperty("columnNumber").GetInt32();
                            var width = arguments.GetProperty("width").GetDouble();
                            var worksheet = GetWorksheet(fileName, sheetName);

                            Microsoft.Office.Interop.Excel.Range column = worksheet.Columns[columnNumber];
                            column.ColumnWidth = width;

                            return $"成功设置第 {columnNumber} 列的列宽为 {width}";
                        }

                    case "insert_rows":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rowIndex = arguments.GetProperty("rowIndex").GetInt32();
                            var count = arguments.TryGetProperty("count", out var countProp) ? countProp.GetInt32() : 1;
                            var worksheet = GetWorksheet(fileName, sheetName);

                            Microsoft.Office.Interop.Excel.Range row = worksheet.Rows[rowIndex];
                            for (int i = 0; i < count; i++)
                            {
                                row.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, Type.Missing);
                            }

                            return $"成功在第 {rowIndex} 行插入了 {count} 行";
                        }

                    case "insert_columns":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var columnIndex = arguments.GetProperty("columnIndex").GetInt32();
                            var count = arguments.TryGetProperty("count", out var countProp) ? countProp.GetInt32() : 1;
                            var worksheet = GetWorksheet(fileName, sheetName);

                            Microsoft.Office.Interop.Excel.Range column = worksheet.Columns[columnIndex];
                            for (int i = 0; i < count; i++)
                            {
                                column.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight, Type.Missing);
                            }

                            return $"成功在第 {columnIndex} 列插入了 {count} 列";
                        }

                    case "delete_rows":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rowIndex = arguments.GetProperty("rowIndex").GetInt32();
                            var count = arguments.TryGetProperty("count", out var countProp) ? countProp.GetInt32() : 1;
                            var worksheet = GetWorksheet(fileName, sheetName);

                            Microsoft.Office.Interop.Excel.Range rows = worksheet.Range[worksheet.Rows[rowIndex], worksheet.Rows[rowIndex + count - 1]];
                            rows.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);

                            return $"成功删除从第 {rowIndex} 行开始的 {count} 行";
                        }

                    case "delete_columns":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var columnIndex = arguments.GetProperty("columnIndex").GetInt32();
                            var count = arguments.TryGetProperty("count", out var countProp) ? countProp.GetInt32() : 1;
                            var worksheet = GetWorksheet(fileName, sheetName);

                            Microsoft.Office.Interop.Excel.Range columns = worksheet.Range[worksheet.Columns[columnIndex], worksheet.Columns[columnIndex + count - 1]];
                            columns.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);

                            return $"成功删除从第 {columnIndex} 列开始的 {count} 列";
                        }

                    case "copy_worksheet":
                        {
                            var fileName = GetFileName();
                            var sourceSheetName = arguments.GetProperty("sourceSheetName").GetString();
                            var targetSheetName = arguments.GetProperty("targetSheetName").GetString();
                            var workbook = GetCurrentWorkbook(fileName);

                            Microsoft.Office.Interop.Excel.Worksheet sourceSheet = workbook.Worksheets[sourceSheetName];
                            sourceSheet.Copy(Type.Missing, workbook.Worksheets[workbook.Worksheets.Count]);
                            Microsoft.Office.Interop.Excel.Worksheet newSheet = workbook.Worksheets[workbook.Worksheets.Count];
                            newSheet.Name = targetSheetName;

                            _activeWorksheet = targetSheetName;
                            return $"成功将工作表 {sourceSheetName} 复制为 {targetSheetName}";
                        }

                    case "copy_range":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var sourceRange = arguments.GetProperty("sourceRange").GetString();
                            var targetRange = arguments.GetProperty("targetRange").GetString();
                            var worksheet = GetWorksheet(fileName, sheetName);

                            Microsoft.Office.Interop.Excel.Range source = worksheet.Range[sourceRange];
                            Microsoft.Office.Interop.Excel.Range target = worksheet.Range[targetRange];
                            source.Copy(target);

                            return $"成功将范围 {sourceRange} 复制到 {targetRange}";
                        }

                    case "clear_range":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var clearType = arguments.TryGetProperty("clearType", out var typeProp) ? typeProp.GetString() : "all";
                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            switch (clearType.ToLower())
                            {
                                case "contents":
                                    range.ClearContents();
                                    break;
                                case "formats":
                                    range.ClearFormats();
                                    break;
                                case "all":
                                default:
                                    range.Clear();
                                    break;
                            }

                            return $"成功清除范围 {rangeAddress} 的{clearType}";
                        }

                    case "get_workbook_metadata":
                        {
                            var fileName = GetFileName();
                            var includeRanges = arguments.TryGetProperty("includeRanges", out var includeProp) && includeProp.GetBoolean();
                            var workbook = GetCurrentWorkbook(fileName);

                            var metadata = new System.Text.StringBuilder();
                            metadata.AppendLine($"工作簿名称: {workbook.Name}");
                            metadata.AppendLine($"工作表数量: {workbook.Worksheets.Count}");
                            metadata.AppendLine($"完整路径: {workbook.FullName}");
                            metadata.AppendLine("工作表列表:");

                            foreach (Microsoft.Office.Interop.Excel.Worksheet ws in workbook.Worksheets)
                            {
                                metadata.AppendLine($"  - {ws.Name}");

                                if (includeRanges)
                                {
                                    Microsoft.Office.Interop.Excel.Range usedRange = ws.UsedRange;
                                    metadata.AppendLine($"    已使用范围: {usedRange.Address}");
                                    metadata.AppendLine($"    行数: {usedRange.Rows.Count}, 列数: {usedRange.Columns.Count}");
                                }
                            }

                            return metadata.ToString();
                        }

                    case "set_data_validation":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var validationType = arguments.GetProperty("validationType").GetString();
                            var operatorType = arguments.TryGetProperty("operatorType", out var opProp) ? opProp.GetString() : "between";
                            var formula1 = arguments.TryGetProperty("formula1", out var f1Prop) ? f1Prop.GetString() : null;
                            var formula2 = arguments.TryGetProperty("formula2", out var f2Prop) ? f2Prop.GetString() : null;
                            var inputMessage = arguments.TryGetProperty("inputMessage", out var imProp) ? imProp.GetString() : null;
                            var errorMessage = arguments.TryGetProperty("errorMessage", out var emProp) ? emProp.GetString() : null;

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            // 删除现有验证
                            range.Validation.Delete();

                            // 设置验证类型
                            Microsoft.Office.Interop.Excel.XlDVType xlType = validationType.ToLower() switch
                            {
                                "whole" => Microsoft.Office.Interop.Excel.XlDVType.xlValidateWholeNumber,
                                "decimal" => Microsoft.Office.Interop.Excel.XlDVType.xlValidateDecimal,
                                "list" => Microsoft.Office.Interop.Excel.XlDVType.xlValidateList,
                                "date" => Microsoft.Office.Interop.Excel.XlDVType.xlValidateDate,
                                "time" => Microsoft.Office.Interop.Excel.XlDVType.xlValidateTime,
                                "textlength" => Microsoft.Office.Interop.Excel.XlDVType.xlValidateTextLength,
                                "custom" => Microsoft.Office.Interop.Excel.XlDVType.xlValidateCustom,
                                _ => Microsoft.Office.Interop.Excel.XlDVType.xlValidateInputOnly
                            };

                            // 设置操作符类型
                            Microsoft.Office.Interop.Excel.XlFormatConditionOperator xlOperator = operatorType.ToLower() switch
                            {
                                "between" => Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween,
                                "notbetween" => Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlNotBetween,
                                "equal" => Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlEqual,
                                "notequal" => Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlNotEqual,
                                "greater" => Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlGreater,
                                "less" => Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlLess,
                                "greaterorequal" => Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlGreaterEqual,
                                "lessorequal" => Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlLessEqual,
                                _ => Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlBetween
                            };

                            // 添加验证
                            range.Validation.Add(xlType, Microsoft.Office.Interop.Excel.XlDVAlertStyle.xlValidAlertStop, xlOperator, formula1, formula2);

                            // 设置输入提示
                            if (!string.IsNullOrEmpty(inputMessage))
                            {
                                range.Validation.IgnoreBlank = true;
                                range.Validation.InCellDropdown = true;
                                range.Validation.ShowInput = true;
                                range.Validation.InputTitle = "输入提示";
                                range.Validation.InputMessage = inputMessage;
                            }

                            // 设置错误提示
                            if (!string.IsNullOrEmpty(errorMessage))
                            {
                                range.Validation.ShowError = true;
                                range.Validation.ErrorTitle = "输入错误";
                                range.Validation.ErrorMessage = errorMessage;
                            }

                            return $"成功为范围 {rangeAddress} 设置数据验证规则";
                        }

                    case "get_validation_rules":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.TryGetProperty("rangeAddress", out var raProp) ? raProp.GetString() : null;
                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = string.IsNullOrEmpty(rangeAddress) ? worksheet.UsedRange : worksheet.Range[rangeAddress];

                            var result = new System.Text.StringBuilder();
                            result.AppendLine($"范围 {range.Address} 的数据验证规则:");

                            try
                            {
                                if (range.Validation != null)
                                {
                                    result.AppendLine($"  类型: {range.Validation.Type}");
                                    result.AppendLine($"  公式1: {range.Validation.Formula1}");
                                    result.AppendLine($"  输入提示: {range.Validation.InputMessage}");
                                    result.AppendLine($"  错误提示: {range.Validation.ErrorMessage}");
                                }
                                else
                                {
                                    result.AppendLine("  无验证规则");
                                }
                            }
                            catch
                            {
                                result.AppendLine("  无验证规则");
                            }

                            return result.ToString();
                        }

                    case "set_number_format":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var formatCode = arguments.GetProperty("formatCode").GetString();
                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            range.NumberFormat = formatCode;

                            return $"成功设置范围 {rangeAddress} 的数字格式为 {formatCode}";
                        }

                    case "apply_conditional_formatting":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var ruleType = arguments.GetProperty("ruleType").GetString();
                            var formula1 = arguments.TryGetProperty("formula1", out var f1) ? f1.GetString() : null;
                            var formula2 = arguments.TryGetProperty("formula2", out var f2) ? f2.GetString() : null;
                            var color1 = arguments.TryGetProperty("color1", out var c1) ? c1.GetString() : null;
                            var color2 = arguments.TryGetProperty("color2", out var c2) ? c2.GetString() : null;
                            var color3 = arguments.TryGetProperty("color3", out var c3) ? c3.GetString() : null;

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            // 清除现有条件格式
                            range.FormatConditions.Delete();

                            switch (ruleType.ToLower())
                            {
                                case "cellvalue":
                                    var condition = range.FormatConditions.Add(
                                        Microsoft.Office.Interop.Excel.XlFormatConditionType.xlCellValue,
                                        Microsoft.Office.Interop.Excel.XlFormatConditionOperator.xlGreater,
                                        formula1);
                                    condition.Interior.Color = ParseColor(color1 ?? "yellow");
                                    break;

                                case "colorscale":
                                    var colorScale = range.FormatConditions.AddColorScale(3);
                                    colorScale.ColorScaleCriteria[1].Type = Microsoft.Office.Interop.Excel.XlConditionValueTypes.xlConditionValueLowestValue;
                                    colorScale.ColorScaleCriteria[1].FormatColor.Color = ParseColor(color1 ?? "red");
                                    colorScale.ColorScaleCriteria[2].Type = Microsoft.Office.Interop.Excel.XlConditionValueTypes.xlConditionValuePercentile;
                                    colorScale.ColorScaleCriteria[2].Value = 50;
                                    colorScale.ColorScaleCriteria[2].FormatColor.Color = ParseColor(color2 ?? "yellow");
                                    colorScale.ColorScaleCriteria[3].Type = Microsoft.Office.Interop.Excel.XlConditionValueTypes.xlConditionValueHighestValue;
                                    colorScale.ColorScaleCriteria[3].FormatColor.Color = ParseColor(color3 ?? "green");
                                    break;

                                case "databar":
                                    var databar = range.FormatConditions.AddDatabar();
                                    databar.BarColor.Color = ParseColor(color1 ?? "blue");
                                    break;

                                case "iconset":
                                    // 图标集 - AddIconSetCondition会自动应用默认图标集(3个交通灯)
                                    var iconSet = range.FormatConditions.AddIconSetCondition();
                                    // 默认已经是3个交通灯图标集，无需额外设置
                                    break;

                                case "expression":
                                    var exprCondition = range.FormatConditions.Add(
                                        Microsoft.Office.Interop.Excel.XlFormatConditionType.xlExpression,
                                        Type.Missing,
                                        formula1);
                                    exprCondition.Interior.Color = ParseColor(color1 ?? "yellow");
                                    break;
                            }

                            return $"成功为范围 {rangeAddress} 应用条件格式 ({ruleType})";
                        }

                    case "create_chart":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var chartType = arguments.TryGetProperty("chartType", out var chartTypeProp) ? chartTypeProp.GetString() : "column";
                            
                            // 支持 dataRange 或 rangeAddress 作为数据范围参数
                            string dataRange = null;
                            if (arguments.TryGetProperty("dataRange", out var dataRangeProp))
                                dataRange = dataRangeProp.GetString();
                            else if (arguments.TryGetProperty("rangeAddress", out var rangeAddressProp))
                                dataRange = rangeAddressProp.GetString();
                            else if (arguments.TryGetProperty("range", out var rangeProp))
                                dataRange = rangeProp.GetString();
                            
                            if (string.IsNullOrEmpty(dataRange))
                                return "错误: 缺少数据范围参数 (dataRange 或 rangeAddress)";
                            
                            var title = arguments.TryGetProperty("title", out var titleProp) ? titleProp.GetString() : null;
                            var width = arguments.TryGetProperty("width", out var widthProp) ? widthProp.GetInt32() : 400;
                            var height = arguments.TryGetProperty("height", out var heightProp) ? heightProp.GetInt32() : 300;

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var dataRangeObj = worksheet.Range[dataRange];
                            
                            // chartPosition 可选，默认放在数据区域右侧
                            Microsoft.Office.Interop.Excel.Range chartPositionObj;
                            if (arguments.TryGetProperty("chartPosition", out var chartPosProp) && !string.IsNullOrEmpty(chartPosProp.GetString()))
                            {
                                chartPositionObj = worksheet.Range[chartPosProp.GetString()];
                            }
                            else
                            {
                                // 默认位置：数据区域右侧偏移一列
                                chartPositionObj = dataRangeObj.Offset[0, dataRangeObj.Columns.Count + 1];
                            }

                            // 创建图表
                            var chartObjects = worksheet.ChartObjects(Type.Missing);
                            var chartObject = chartObjects.Add(
                                (double)chartPositionObj.Left,
                                (double)chartPositionObj.Top,
                                width,
                                height);

                            var chart = chartObject.Chart;

                            // 设置图表类型
                            Microsoft.Office.Interop.Excel.XlChartType xlChartType = chartType.ToLower() switch
                            {
                                "line" => Microsoft.Office.Interop.Excel.XlChartType.xlLine,
                                "bar" => Microsoft.Office.Interop.Excel.XlChartType.xlBarClustered,
                                "column" => Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered,
                                "pie" => Microsoft.Office.Interop.Excel.XlChartType.xlPie,
                                "scatter" => Microsoft.Office.Interop.Excel.XlChartType.xlXYScatter,
                                "area" => Microsoft.Office.Interop.Excel.XlChartType.xlArea,
                                "radar" => Microsoft.Office.Interop.Excel.XlChartType.xlRadar,
                                _ => Microsoft.Office.Interop.Excel.XlChartType.xlColumnClustered
                            };

                            chart.ChartType = xlChartType;
                            chart.SetSourceData(dataRangeObj);

                            // 设置标题
                            if (!string.IsNullOrEmpty(title))
                            {
                                chart.HasTitle = true;
                                chart.ChartTitle.Text = title;
                            }

                            return $"成功创建 {chartType} 图表";
                        }

                    case "create_table":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var tableName = arguments.GetProperty("tableName").GetString();
                            var hasHeaders = arguments.TryGetProperty("hasHeaders", out var headersProp) ? headersProp.GetBoolean() : true;
                            var tableStyle = arguments.TryGetProperty("tableStyle", out var styleProp) ? styleProp.GetString() : "TableStyleMedium2";

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            // 创建表格
                            var table = worksheet.ListObjects.Add(
                                Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange,
                                range,
                                Type.Missing,
                                hasHeaders ? Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes : Microsoft.Office.Interop.Excel.XlYesNoGuess.xlNo,
                                Type.Missing);

                            table.Name = tableName;

                            // 设置表格样式
                            try
                            {
                                table.TableStyle = tableStyle;
                            }
                            catch
                            {
                                // 如果样式不存在，使用默认样式
                                table.TableStyle = "TableStyleMedium2";
                            }

                            return $"成功创建表格 {tableName}";
                        }

                    case "get_table_names":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var worksheet = GetWorksheet(fileName, sheetName);

                            var tableNames = new List<string>();
                            foreach (Microsoft.Office.Interop.Excel.ListObject table in worksheet.ListObjects)
                            {
                                tableNames.Add(table.Name);
                            }

                            return $"工作表中的表格: {string.Join(", ", tableNames)}";
                        }

                    case "validate_formula":
                        {
                            var formula = arguments.GetProperty("formula").GetString();

                            try
                            {
                                // 创建临时工作簿进行公式验证
                                var tempWorkbook = ThisAddIn.app.Workbooks.Add();
                                var tempSheet = tempWorkbook.Worksheets[1];
                                var tempCell = tempSheet.Cells[1, 1];

                                try
                                {
                                    tempCell.Formula = formula;
                                    tempWorkbook.Close(false);
                                    return "公式语法正确";
                                }
                                catch (Exception ex)
                                {
                                    tempWorkbook.Close(false);
                                    return $"公式语法错误: {ex.Message}";
                                }
                            }
                            catch (Exception ex)
                            {
                                return $"验证失败: {ex.Message}";
                            }
                        }

                    case "find_value":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var searchValue = arguments.GetProperty("searchValue").GetString();
                            var matchCase = arguments.TryGetProperty("matchCase", out var mcProp) && mcProp.GetBoolean();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var usedRange = worksheet.UsedRange;
                            var results = new List<string>();

                            var foundCell = usedRange.Find(
                                What: searchValue,
                                LookIn: Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                                LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                                SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
                                MatchCase: matchCase);

                            if (foundCell != null)
                            {
                                string firstAddress = foundCell.Address;
                                do
                                {
                                    results.Add(foundCell.Address);
                                    foundCell = usedRange.FindNext(foundCell);
                                }
                                while (foundCell != null && foundCell.Address != firstAddress);
                            }

                            return results.Count > 0 
                                ? $"找到 {results.Count} 个匹配项: {string.Join(", ", results)}" 
                                : "未找到匹配项";
                        }

                    case "find_and_replace":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var findWhat = arguments.GetProperty("findWhat").GetString();
                            var replaceWith = arguments.GetProperty("replaceWith").GetString();
                            var matchCase = arguments.TryGetProperty("matchCase", out var mcProp) && mcProp.GetBoolean();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var usedRange = worksheet.UsedRange;
                            int count = 0;

                            var foundCell = usedRange.Find(
                                What: findWhat,
                                LookIn: Microsoft.Office.Interop.Excel.XlFindLookIn.xlValues,
                                LookAt: Microsoft.Office.Interop.Excel.XlLookAt.xlPart,
                                MatchCase: matchCase);

                            if (foundCell != null)
                            {
                                string firstAddress = foundCell.Address;
                                do
                                {
                                    foundCell.Value = replaceWith;
                                    count++;
                                    foundCell = usedRange.FindNext(foundCell);
                                }
                                while (foundCell != null && foundCell.Address != firstAddress);
                            }

                            return $"成功替换了 {count} 个单元格";
                        }

                    case "freeze_panes":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var row = arguments.GetProperty("row").GetInt32();
                            var column = arguments.GetProperty("column").GetInt32();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var cell = worksheet.Cells[row, column];
                            cell.Select();
                            ThisAddIn.app.ActiveWindow.FreezePanes = true;

                            return $"成功冻结窗格（在行 {row}，列 {column} 处）";
                        }

                    case "unfreeze_panes":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var worksheet = GetWorksheet(fileName, sheetName);

                            ThisAddIn.app.ActiveWindow.FreezePanes = false;
                            return "成功取消冻结窗格";
                        }

                    case "autofit_columns":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];
                            range.Columns.AutoFit();

                            return $"成功自动调整列宽: {rangeAddress}";
                        }

                    case "autofit_rows":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];
                            range.Rows.AutoFit();

                            return $"成功自动调整行高: {rangeAddress}";
                        }

                    case "set_column_visible":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var columnIndex = arguments.GetProperty("columnIndex").GetInt32();
                            var visible = arguments.GetProperty("visible").GetBoolean();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var column = worksheet.Columns[columnIndex];
                            column.Hidden = !visible;

                            return $"成功设置第 {columnIndex} 列的可见性为 {(visible ? "显示" : "隐藏")}";
                        }

                    case "set_row_visible":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rowIndex = arguments.GetProperty("rowIndex").GetInt32();
                            var visible = arguments.GetProperty("visible").GetBoolean();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var row = worksheet.Rows[rowIndex];
                            row.Hidden = !visible;

                            return $"成功设置第 {rowIndex} 行的可见性为 {(visible ? "显示" : "隐藏")}";
                        }

                    case "add_comment":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var cellAddress = arguments.GetProperty("cellAddress").GetString();
                            var commentText = arguments.GetProperty("commentText").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var cell = worksheet.Range[cellAddress];

                            // 如果已有批注，先删除
                            if (cell.Comment != null)
                            {
                                cell.Comment.Delete();
                            }

                            cell.AddComment(commentText);
                            return $"成功为单元格 {cellAddress} 添加批注";
                        }

                    case "delete_comment":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var cellAddress = arguments.GetProperty("cellAddress").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var cell = worksheet.Range[cellAddress];

                            if (cell.Comment != null)
                            {
                                cell.Comment.Delete();
                                return $"成功删除单元格 {cellAddress} 的批注";
                            }

                            return $"单元格 {cellAddress} 没有批注";
                        }

                    case "get_comment":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var cellAddress = arguments.GetProperty("cellAddress").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var cell = worksheet.Range[cellAddress];

                            var commentText = cell.Comment?.Text() ?? "";
                            return string.IsNullOrEmpty(commentText) 
                                ? $"单元格 {cellAddress} 没有批注" 
                                : $"单元格 {cellAddress} 的批注: {commentText}";
                        }

                    case "add_hyperlink":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var cellAddress = arguments.GetProperty("cellAddress").GetString();
                            var url = arguments.GetProperty("url").GetString();
                            var displayText = arguments.TryGetProperty("displayText", out var dtProp) ? dtProp.GetString() : null;

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var cell = worksheet.Range[cellAddress];

                            // 只处理外部链接（网址、文件路径等）
                            // 不处理文档内跳转（应使用 set_hyperlink_formula）
                            worksheet.Hyperlinks.Add(
                                Anchor: cell,
                                Address: url,
                                TextToDisplay: displayText ?? url);

                            return $"成功为单元格 {cellAddress} 添加超链接对象（外部链接）: {url}";
                        }

                    case "set_hyperlink_formula":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var cellAddress = arguments.GetProperty("cellAddress").GetString();
                            var targetLocation = arguments.GetProperty("targetLocation").GetString();
                            var displayText = arguments.GetProperty("displayText").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var cell = worksheet.Range[cellAddress];

                            // 使用 HYPERLINK 公式
                            // 格式：=HYPERLINK("#工作表名!单元格", "显示文本")
                            var formula = $"=HYPERLINK(\"#{targetLocation}\", \"{displayText}\")";
                            cell.Formula = formula;

                            return $"成功为单元格 {cellAddress} 设置HYPERLINK公式，跳转目标: {targetLocation}";
                        }

                    case "delete_hyperlink":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var cellAddress = arguments.GetProperty("cellAddress").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var cell = worksheet.Range[cellAddress];

                            if (cell.Hyperlinks.Count > 0)
                            {
                                cell.Hyperlinks.Delete();
                                return $"成功删除单元格 {cellAddress} 的超链接";
                            }

                            return $"单元格 {cellAddress} 没有超链接";
                        }

                    case "get_used_range":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var worksheet = GetWorksheet(fileName, sheetName);

                            var usedRange = worksheet.UsedRange;
                            var address = usedRange.Address;

                            return $"工作表 {sheetName} 的已使用范围: {address}";
                        }

                    case "get_range_statistics":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            var stats = new System.Text.StringBuilder();
                            stats.AppendLine($"范围 {rangeAddress} 的统计信息:");

                            try
                            {
                                stats.AppendLine($"  总和: {ThisAddIn.app.WorksheetFunction.Sum(range)}");
                            }
                            catch { stats.AppendLine("  总和: N/A"); }

                            try
                            {
                                stats.AppendLine($"  平均值: {ThisAddIn.app.WorksheetFunction.Average(range)}");
                            }
                            catch { stats.AppendLine("  平均值: N/A"); }

                            try
                            {
                                stats.AppendLine($"  计数: {ThisAddIn.app.WorksheetFunction.Count(range)}");
                            }
                            catch { stats.AppendLine("  计数: N/A"); }

                            try
                            {
                                stats.AppendLine($"  最大值: {ThisAddIn.app.WorksheetFunction.Max(range)}");
                            }
                            catch { stats.AppendLine("  最大值: N/A"); }

                            try
                            {
                                stats.AppendLine($"  最小值: {ThisAddIn.app.WorksheetFunction.Min(range)}");
                            }
                            catch { stats.AppendLine("  最小值: N/A"); }

                            return stats.ToString();
                        }

                    case "get_last_row":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var columnIndex = arguments.TryGetProperty("columnIndex", out var ciProp) ? ciProp.GetInt32() : 1;

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var column = worksheet.Columns[columnIndex];
                            var lastCell = column.Find("*", 
                                SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows, 
                                SearchDirection: Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious);

                            var lastRow = lastCell?.Row ?? 0;
                            return $"列 {columnIndex} 的最后一行: {lastRow}";
                        }

                    case "get_last_column":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rowIndex = arguments.TryGetProperty("rowIndex", out var riProp) ? riProp.GetInt32() : 1;

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var row = worksheet.Rows[rowIndex];
                            var lastCell = row.Find("*", 
                                SearchOrder: Microsoft.Office.Interop.Excel.XlSearchOrder.xlByColumns, 
                                SearchDirection: Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious);

                            var lastColumn = lastCell?.Column ?? 0;
                            return $"行 {rowIndex} 的最后一列: {lastColumn}";
                        }

                    case "sort_range":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var sortColumnIndex = arguments.GetProperty("sortColumnIndex").GetInt32();
                            var ascending = arguments.TryGetProperty("ascending", out var ascProp) ? ascProp.GetBoolean() : true;

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];
                            var sortKey = range.Columns[sortColumnIndex];

                            range.Sort(
                                Key1: sortKey,
                                Order1: ascending ? Microsoft.Office.Interop.Excel.XlSortOrder.xlAscending : Microsoft.Office.Interop.Excel.XlSortOrder.xlDescending,
                                Header: Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes);

                            return $"成功对范围 {rangeAddress} 按第 {sortColumnIndex} 列进行{(ascending ? "升序" : "降序")}排序";
                        }

                    case "set_auto_filter":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var columnIndex = arguments.TryGetProperty("columnIndex", out var ciProp) ? ciProp.GetInt32() : 0;
                            var criteria = arguments.TryGetProperty("criteria", out var cProp) ? cProp.GetString() : null;

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            // 如果已有筛选，先清除
                            if (worksheet.AutoFilterMode)
                            {
                                worksheet.AutoFilterMode = false;
                            }

                            if (columnIndex > 0 && !string.IsNullOrEmpty(criteria))
                            {
                                range.AutoFilter(Field: columnIndex, Criteria1: criteria);
                                return $"成功为范围 {rangeAddress} 的第 {columnIndex} 列设置筛选条件: {criteria}";
                            }
                            else
                            {
                                range.AutoFilter();
                                return $"成功为范围 {rangeAddress} 设置自动筛选";
                            }
                        }

                    case "remove_duplicates":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var columnIndicesJson = arguments.GetProperty("columnIndices").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            var columnIndices = JsonSerializer.Deserialize<int[]>(columnIndicesJson);
                            var columns = columnIndices.Cast<object>().ToArray();

                            range.RemoveDuplicates(Columns: columns, Header: Microsoft.Office.Interop.Excel.XlYesNoGuess.xlYes);
                            return $"成功删除范围 {rangeAddress} 中的重复项";
                        }

                    case "move_worksheet":
                        {
                            var fileName = GetFileName();
                            var sheetName = arguments.GetProperty("sheetName").GetString();
                            var position = arguments.GetProperty("position").GetInt32();

                            var workbook = GetCurrentWorkbook(fileName);
                            var worksheet = workbook.Worksheets[sheetName];

                            if (position == 1)
                            {
                                worksheet.Move(Before: workbook.Worksheets[1]);
                            }
                            else if (position > workbook.Worksheets.Count)
                            {
                                worksheet.Move(After: workbook.Worksheets[workbook.Worksheets.Count]);
                            }
                            else
                            {
                                worksheet.Move(Before: workbook.Worksheets[position]);
                            }

                            return $"成功将工作表 {sheetName} 移动到位置 {position}";
                        }

                    case "set_worksheet_visible":
                        {
                            var fileName = GetFileName();
                            var sheetName = arguments.GetProperty("sheetName").GetString();
                            var visible = arguments.GetProperty("visible").GetBoolean();

                            var workbook = GetCurrentWorkbook(fileName);
                            var worksheet = workbook.Worksheets[sheetName];
                            worksheet.Visible = visible ? Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetVisible : Microsoft.Office.Interop.Excel.XlSheetVisibility.xlSheetHidden;

                            return $"成功设置工作表 {sheetName} 的可见性为 {(visible ? "显示" : "隐藏")}";
                        }

                    case "get_worksheet_index":
                        {
                            var fileName = GetFileName();
                            var sheetName = arguments.GetProperty("sheetName").GetString();

                            var workbook = GetCurrentWorkbook(fileName);
                            var worksheet = workbook.Worksheets[sheetName];
                            var index = worksheet.Index;

                            return $"工作表 {sheetName} 的位置索引: {index}";
                        }

                    // 命名区域工具执行
                    case "create_named_range":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeName = arguments.GetProperty("rangeName").GetString();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();

                            _excelMcp.CreateNamedRange(fileName, sheetName, rangeName, rangeAddress);
                            return $"成功创建命名区域 '{rangeName}' 引用 {rangeAddress}";
                        }

                    case "delete_named_range":
                        {
                            var fileName = GetFileName();
                            var rangeName = arguments.GetProperty("rangeName").GetString();

                            _excelMcp.DeleteNamedRange(fileName, rangeName);
                            return $"成功删除命名区域 '{rangeName}'";
                        }

                    case "get_named_ranges":
                        {
                            var fileName = GetFileName();
                            var namedRanges = _excelMcp.GetNamedRanges(fileName);

                            if (namedRanges.Count == 0)
                                return "工作簿中没有命名区域";

                            return $"工作簿中的命名区域：\n{string.Join("\n", namedRanges)}";
                        }

                    case "get_named_range_address":
                        {
                            var fileName = GetFileName();
                            var rangeName = arguments.GetProperty("rangeName").GetString();

                            var address = _excelMcp.GetNamedRangeAddress(fileName, rangeName);
                            return $"命名区域 '{rangeName}' 的引用地址: {address}";
                        }

                    // 单元格格式增强工具执行
                    case "set_cell_text_wrap":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var wrap = arguments.GetProperty("wrap").GetBoolean();

                            _excelMcp.SetCellTextWrap(fileName, sheetName, rangeAddress, wrap);
                            return $"成功设置 {rangeAddress} 的文本换行为: {(wrap ? "启用" : "禁用")}";
                        }

                    case "set_cell_indent":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var indentLevel = arguments.GetProperty("indentLevel").GetInt32();

                            _excelMcp.SetCellIndent(fileName, sheetName, rangeAddress, indentLevel);
                            return $"成功设置 {rangeAddress} 的缩进级别为: {indentLevel}";
                        }

                    case "set_cell_orientation":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var degrees = arguments.GetProperty("degrees").GetInt32();

                            _excelMcp.SetCellOrientation(fileName, sheetName, rangeAddress, degrees);
                            return $"成功设置 {rangeAddress} 的文本旋转角度为: {degrees}度";
                        }

                    case "set_cell_shrink_to_fit":
                        {
                            var fileName = GetFileName();
                            var sheetName = GetSheetName();
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
                            var shrink = arguments.GetProperty("shrink").GetBoolean();

                            _excelMcp.SetCellShrinkToFit(fileName, sheetName, rangeAddress, shrink);
                            return $"成功设置 {rangeAddress} 的缩小字体填充为: {(shrink ? "启用" : "禁用")}";
                        }

                    case "get_current_selection":
                        {
                            try
                            {
                                if (ThisAddIn.app == null || ThisAddIn.app.Selection == null)
                                    return "无法获取当前选中的单元格";

                                var selection = ThisAddIn.app.Selection as Microsoft.Office.Interop.Excel.Range;
                                if (selection == null)
                                    return "当前没有选中单元格区域";

                                var result = new System.Text.StringBuilder();
                                result.AppendLine("当前选中的单元格信息:");
                                result.AppendLine($"- 地址: {selection.Address}");
                                result.AppendLine($"- 行号: {selection.Row}");
                                result.AppendLine($"- 列号: {selection.Column}");
                                result.AppendLine($"- 行数: {selection.Rows.Count}");
                                result.AppendLine($"- 列数: {selection.Columns.Count}");

                                // 如果是单个单元格，显示值
                                if (selection.Cells.Count == 1)
                                {
                                    result.AppendLine($"- 值: {selection.Value?.ToString() ?? "(空)"}");
                                    if (selection.HasFormula)
                                    {
                                        result.AppendLine($"- 公式: {selection.Formula}");
                                    }
                                }
                                else
                                {
                                    result.AppendLine($"- 单元格总数: {selection.Cells.Count}");
                                }

                                if (ThisAddIn.app.ActiveWorkbook != null)
                                {
                                    result.AppendLine($"- 所属工作簿: {ThisAddIn.app.ActiveWorkbook.Name}");
                                }

                                if (ThisAddIn.app.ActiveSheet != null)
                                {
                                    Microsoft.Office.Interop.Excel.Worksheet ws = ThisAddIn.app.ActiveSheet;
                                    result.AppendLine($"- 所属工作表: {ws.Name}");
                                }

                                return result.ToString();
                            }
                            catch (Exception ex)
                            {
                                return $"获取当前选中单元格信息失败: {ex.Message}";
                            }
                        }

                    default:
                        return $"未知的工具: {toolName}";
                }
            }
            catch (Exception ex)
            {
                return $"执行工具 {toolName} 时出错: {ex.Message}";
            }
        }

        // 辅助方法：解析颜色
        private int ParseColor(string colorStr)
        {
            // 支持颜色名称和十六进制颜色
            if (colorStr.StartsWith("#"))
            {
                // 十六进制颜色 #RRGGBB
                var hex = colorStr.Substring(1);
                var r = Convert.ToInt32(hex.Substring(0, 2), 16);
                var g = Convert.ToInt32(hex.Substring(2, 2), 16);
                var b = Convert.ToInt32(hex.Substring(4, 2), 16);
                return System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(r, g, b));
            }
            else
            {
                // 颜色名称
                var color = colorStr.ToLower() switch
                {
                    "红色" or "red" => System.Drawing.Color.Red,
                    "绿色" or "green" => System.Drawing.Color.Green,
                    "蓝色" or "blue" => System.Drawing.Color.Blue,
                    "黄色" or "yellow" => System.Drawing.Color.Yellow,
                    "橙色" or "orange" => System.Drawing.Color.Orange,
                    "紫色" or "purple" => System.Drawing.Color.Purple,
                    "黑色" or "black" => System.Drawing.Color.Black,
                    "白色" or "white" => System.Drawing.Color.White,
                    "灰色" or "gray" => System.Drawing.Color.Gray,
                    _ => System.Drawing.Color.Black
                };
                return System.Drawing.ColorTranslator.ToOle(color);
            }
        }

        //获取对话请求
        private async Task<string> GetDeepSeekResponse(string userInput)
        {
            string apiKey = _apiKey;
            string apiUrl = _apiUrl;
            bool useMcp = checkBoxUseMcp.Checked;

            // 根据用户勾选情况设置Prompt Engineering模式
            if (_isPromptEngineeringChecked)
            {
                // 用户勾选了"优先提示工程"，强制使用Prompt Engineering模式
                _usePromptEngineering = true;
            }
            else
            {
                // 用户没有勾选，重置为自动判断模式（后续会根据模型大小自动判断）
                _usePromptEngineering = false;
            }

            // 记录用户输入
            WriteLog("用户输入", userInput);

            // 将用户消息加入历史
            _chatHistory.Add(new ChatMessage
            {
                Role = "user",
                Content = userInput
            });

            using (var client = new HttpClient())
            {
                // 设置较长的超时时间，本地模型可能需要更长时间响应
                client.Timeout = TimeSpan.FromMinutes(_timeoutMinutes);

                // 只有云端连接时才添加Authorization头
                if (_isCloudConnection && !string.IsNullOrEmpty(apiKey))
                {
                    client.DefaultRequestHeaders.Authorization =
                        new AuthenticationHeaderValue("Bearer", apiKey);
                }

                // 构建请求体
                var requestBody = new Dictionary<string, object>
                {
                    { "model", _model },
                    { "messages", BuildMessages(useMcp, userInput) },
                    { "temperature", 0.7 },
                    { "max_tokens", 2000 }
                };

                // 仅对Ollama API添加特有参数（LM Studio等其他本地服务不支持这些参数）
                if (!_isCloudConnection && _isOllamaApi)
                {
                    // 禁用Qwen3的思考模式，大幅提升响应速度
                    requestBody["options"] = new Dictionary<string, object>
                    {
                        { "num_predict", 1000 },  // 限制生成token数
                        { "temperature", 0.7 }
                    };
                    // 对于支持的模型，尝试禁用思考模式
                    requestBody["think"] = false;
                }

                // 检测是否为小模型（参数量小于3B），小模型直接使用Prompt Engineering模式
                // 或者用户勾选了"优先提示工程"，强制使用Prompt Engineering模式
                bool isSmallModel = !_isCloudConnection && IsSmallModel(_model);
                if ((isSmallModel || _isPromptEngineeringChecked) && useMcp && !_usePromptEngineering)
                {
                    string reason = _isPromptEngineeringChecked ? "用户勾选了'优先提示工程'" : $"检测到小模型 {_model}";
                    WriteLog("模式切换", $"{reason}，切换到Prompt Engineering模式");
                    _usePromptEngineering = true;
                    // 重新构建消息（包含Prompt Engineering系统提示）
                    requestBody["messages"] = BuildMessages(useMcp, userInput);
                }

                // 如果启用MCP且ExcelMcp可用，且不是Prompt Engineering模式，添加工具定义
                if (useMcp && _excelMcp != null && !_usePromptEngineering)
                {
                    // 对于本地模型，使用智能工具选择减少token数量
                    if (!_isCloudConnection && _useToolGrouping)
                    {
                        // 根据用户输入预选相关工具组
                        var preSelectedGroups = PreSelectToolGroups(userInput);
                        var selectedTools = GetToolsByGroups(preSelectedGroups);
                        requestBody["tools"] = selectedTools;
                        WriteLog("智能工具选择", $"根据用户输入预选工具组: [{string.Join(", ", preSelectedGroups)}], 工具数量: {selectedTools.Count}");
                    }
                    else
                    {
                        // 云端模型或禁用分组时，发送全部工具
                        requestBody["tools"] = GetMcpTools();
                    }
                }

                // 记录请求信息（简化版，不包含完整工具定义）
                var requestJsonForLog = GetSimplifiedRequestBodyForLog(requestBody);
                WriteLog("API请求", $"URL: {apiUrl}\n模型: {_model}\nPrompt Engineering模式: {_usePromptEngineering}\n请求体:\n{requestJsonForLog}");

                var response = await client.PostAsJsonAsync(apiUrl, requestBody);
                var responseContent = await response.Content.ReadAsStringAsync();

                // 记录响应信息
                WriteLog("API响应", $"状态码: {response.StatusCode}\n响应内容:\n{responseContent}");

                System.Diagnostics.Debug.WriteLine($"API响应状态: {response.StatusCode}");
                System.Diagnostics.Debug.WriteLine($"API响应内容: {responseContent.Substring(0, Math.Min(500, responseContent.Length))}");

                if (!response.IsSuccessStatusCode)
                {
                    // 检查是否是因为不支持tools参数导致的错误（本地模型）
                    // 扩展检测条件：BadRequest通常表示请求格式不被支持
                    bool shouldSwitchToPromptEngineering = useMcp && !_isCloudConnection && !_usePromptEngineering &&
                        (response.StatusCode == System.Net.HttpStatusCode.BadRequest ||
                         responseContent.Contains("tools") || responseContent.Contains("tool") ||
                         responseContent.Contains("function") || responseContent.Contains("not supported") ||
                         responseContent.Contains("invalid") || responseContent.Contains("unknown"));

                    WriteLog("模式检测", $"请求失败，状态码: {response.StatusCode}\n是否应切换到Prompt Engineering: {shouldSwitchToPromptEngineering}\n原因: 本地模型={!_isCloudConnection}, 使用MCP={useMcp}, 当前非PE模式={!_usePromptEngineering}");

                    if (shouldSwitchToPromptEngineering)
                    {
                        // 本地模型不支持function calling，切换到Prompt Engineering模式
                        WriteLog("模式切换", "本地模型不支持function calling或请求格式不兼容，切换到Prompt Engineering模式");
                        System.Diagnostics.Debug.WriteLine("本地模型不支持function calling或请求格式不兼容，切换到Prompt Engineering模式");
                        _usePromptEngineering = true;

                        // 移除tools参数，重新构建消息（包含Prompt Engineering系统提示）
                        requestBody.Remove("tools");
                        requestBody["messages"] = BuildMessages(useMcp, userInput);

                        // 记录重试请求（简化版）
                        var retryRequestJsonForLog = GetSimplifiedRequestBodyForLog(requestBody);
                        WriteLog("重试请求(Prompt Engineering)", $"URL: {apiUrl}\n请求体:\n{retryRequestJsonForLog}");

                        response = await client.PostAsJsonAsync(apiUrl, requestBody);
                        responseContent = await response.Content.ReadAsStringAsync();

                        WriteLog("重试响应", $"状态码: {response.StatusCode}\n响应内容:\n{responseContent}");
                        System.Diagnostics.Debug.WriteLine($"重试后API响应状态: {response.StatusCode}");

                        if (!response.IsSuccessStatusCode)
                        {
                            throw new HttpRequestException($"HTTP Error: {response.StatusCode}, 响应: {responseContent.Substring(0, Math.Min(200, responseContent.Length))}");
                        }
                    }
                    else
                    {
                        throw new HttpRequestException($"HTTP Error: {response.StatusCode}");
                    }
                }

                var jsonResponse = JsonSerializer.Deserialize<DeepSeekResponse>(responseContent);
                var choice = jsonResponse?.choices[0];

                // 调试信息
                System.Diagnostics.Debug.WriteLine($"AI响应内容: {choice?.message?.content}");
                System.Diagnostics.Debug.WriteLine($"工具调用数量: {choice?.message?.tool_calls?.Length ?? 0}");
                System.Diagnostics.Debug.WriteLine($"Prompt Engineering模式: {_usePromptEngineering}");

                WriteLog("响应解析", $"AI响应内容: {choice?.message?.content}\n原生tool_calls数量: {choice?.message?.tool_calls?.Length ?? 0}\n当前Prompt Engineering模式: {_usePromptEngineering}");

                // 检查本地模型是否支持function calling
                // 如果是本地模型，发送了tools参数但没有返回tool_calls，需要判断是模型不支持还是模型主动选择不调用工具
                if (!_isCloudConnection && useMcp && _excelMcp != null && !_usePromptEngineering)
                {
                    bool hasToolCalls = choice?.message?.tool_calls != null && choice.message.tool_calls.Length > 0;
                    string responseText = choice?.message?.content?.Trim() ?? "";
                    bool hasMeaningfulContent = !string.IsNullOrEmpty(responseText) && responseText.Length > 10;
                    
                    WriteLog("Function Calling检测", $"本地模型是否返回tool_calls: {hasToolCalls}, 是否有有意义的文本内容: {hasMeaningfulContent}, 内容长度: {responseText.Length}");
                    
                    if (!hasToolCalls)
                    {
                        // 如果模型返回了有意义的文本内容（如澄清问题），直接返回给用户，不切换模式
                        if (hasMeaningfulContent)
                        {
                            WriteLog("响应处理", "本地模型未返回tool_calls但有有意义的文本内容，直接返回给用户（可能是澄清问题）");
                            System.Diagnostics.Debug.WriteLine($"本地模型返回文本响应（非工具调用）: {responseText}");
                            
                            // 将AI回复加入历史
                            _chatHistory.Add(new ChatMessage
                            {
                                Role = "assistant",
                                Content = responseText
                            });
                            
                            return responseText;
                        }
                        
                        // 本地模型不支持function calling，切换到Prompt Engineering模式
                        WriteLog("模式切换", "本地模型未返回tool_calls且无有意义内容，切换到Prompt Engineering模式");
                        System.Diagnostics.Debug.WriteLine("本地模型未返回tool_calls，切换到Prompt Engineering模式");
                        _usePromptEngineering = true;

                        // 清空历史记录中刚添加的用户消息，重新发送
                        if (_chatHistory.Count > 0 && _chatHistory[_chatHistory.Count - 1].Role == "user")
                        {
                            var lastUserMessage = _chatHistory[_chatHistory.Count - 1].Content;
                            _chatHistory.RemoveAt(_chatHistory.Count - 1);

                            // 重新添加用户消息
                            _chatHistory.Add(new ChatMessage
                            {
                                Role = "user",
                                Content = lastUserMessage
                            });
                        }

                        // 移除tools参数，重新构建消息（包含Prompt Engineering系统提示）
                        requestBody.Remove("tools");
                        requestBody["messages"] = BuildMessages(useMcp, userInput);

                        // 记录重试请求（简化版）
                        var retryRequestJsonForLog2 = GetSimplifiedRequestBodyForLog(requestBody);
                        WriteLog("重试请求(Prompt Engineering)", $"URL: {apiUrl}\n请求体:\n{retryRequestJsonForLog2}");

                        response = await client.PostAsJsonAsync(apiUrl, requestBody);
                        responseContent = await response.Content.ReadAsStringAsync();

                        WriteLog("重试响应", $"状态码: {response.StatusCode}\n响应内容:\n{responseContent}");

                        if (!response.IsSuccessStatusCode)
                        {
                            throw new HttpRequestException($"HTTP Error: {response.StatusCode}");
                        }

                        jsonResponse = JsonSerializer.Deserialize<DeepSeekResponse>(responseContent);
                        choice = jsonResponse?.choices[0];

                        System.Diagnostics.Debug.WriteLine($"Prompt Engineering模式响应: {choice?.message?.content}");
                    }
                }

                // 如果是Prompt Engineering模式，解析响应中的工具调用
                if (_usePromptEngineering && useMcp && _excelMcp != null)
                {
                    return await HandlePromptEngineeringResponse(client, apiUrl, choice?.message?.content ?? "", userInput);
                }

                // 原生Function Calling模式：检查是否有工具调用
                if (choice?.message?.tool_calls != null && choice.message.tool_calls.Length > 0)
                {
                    // 处理工具调用
                    var toolCalls = choice.message.tool_calls;

                    System.Diagnostics.Debug.WriteLine($"开始执行 {toolCalls.Length} 个工具调用");
                    SafeUpdatePromptLabel($"正在执行 {toolCalls.Length} 个工具操作...");

                    // 将AI的工具调用消息加入历史
                    _chatHistory.Add(new ChatMessage
                    {
                        Role = "assistant",
                        Content = choice.message.content,
                        ToolCalls = toolCalls.Select(tc => new ToolCall
                        {
                            Id = tc.id,
                            Type = tc.type,
                            Function = new FunctionCall
                            {
                                Name = tc.function.name,
                                Arguments = tc.function.arguments
                            }
                        }).ToList()
                    });

                    // 执行每个工具调用
                    foreach (var toolCall in toolCalls)
                    {
                        var functionName = toolCall.function.name;
                        
                        // 检查是否为一次性工具且已执行过
                        if (_oneTimeTools.Contains(functionName) && _executedOneTimeTools.Contains(functionName))
                        {
                            WriteLog("跳过重复工具", $"工具 {functionName} 已在本次请求中执行过，跳过重复执行");
                            // 将跳过信息作为工具结果加入历史
                            _chatHistory.Add(new ChatMessage
                            {
                                Role = "tool",
                                Content = $"工具 {functionName} 已执行过，跳过重复调用",
                                ToolCallId = toolCall.id
                            });
                            continue;
                        }
                        
                        var arguments = JsonSerializer.Deserialize<JsonElement>(toolCall.function.arguments);

                        System.Diagnostics.Debug.WriteLine($"执行工具: {functionName}");
                        System.Diagnostics.Debug.WriteLine($"参数: {toolCall.function.arguments}");
                        SafeUpdatePromptLabel($"正在执行工具: {functionName}...");

                        // 执行工具
                        var toolResult = ExecuteMcpTool(functionName, arguments);

                        System.Diagnostics.Debug.WriteLine($"工具执行结果: {toolResult}");
                        
                        // 记录一次性工具已执行
                        if (_oneTimeTools.Contains(functionName))
                        {
                            _executedOneTimeTools.Add(functionName);
                        }

                        // 将工具结果加入历史
                        _chatHistory.Add(new ChatMessage
                        {
                            Role = "tool",
                            Content = toolResult,
                            ToolCallId = toolCall.id
                        });
                    }

                    // 循环处理工具调用，直到AI不再请求工具
                    while (true)
                    {
                        // 再次调用API获取回复（可能是最终回复或更多工具调用）
                        var finalRequestBody = new Dictionary<string, object>
                        {
                            { "model", _model },
                            { "messages", BuildMessages(useMcp, userInput) },
                            { "temperature", 0.7 },
                            { "max_tokens", 2000 }
                        };

                        // 仅对Ollama API添加特有参数
                        if (!_isCloudConnection && _isOllamaApi)
                        {
                            finalRequestBody["options"] = new Dictionary<string, object>
                            {
                                { "num_predict", 1000 },
                                { "temperature", 0.7 }
                            };
                            finalRequestBody["think"] = false;
                        }

                        if (useMcp && _excelMcp != null)
                        {
                            // 对于本地模型，使用智能工具选择
                            if (!_isCloudConnection && _useToolGrouping)
                            {
                                var preSelectedGroups = PreSelectToolGroups(userInput);
                                finalRequestBody["tools"] = GetToolsByGroups(preSelectedGroups);
                            }
                            else
                            {
                                finalRequestBody["tools"] = GetMcpTools();
                            }
                        }

                        SafeUpdatePromptLabel("等待AI响应...");
                        var finalResponse = await client.PostAsJsonAsync(apiUrl, finalRequestBody);
                        var finalResponseContent = await finalResponse.Content.ReadAsStringAsync();

                        if (!finalResponse.IsSuccessStatusCode)
                        {
                            throw new HttpRequestException($"HTTP Error: {finalResponse.StatusCode}");
                        }

                        var finalJsonResponse = JsonSerializer.Deserialize<DeepSeekResponse>(finalResponseContent);
                        var finalChoice = finalJsonResponse?.choices[0];

                        System.Diagnostics.Debug.WriteLine($"第二轮AI响应内容: {finalChoice?.message?.content}");
                        System.Diagnostics.Debug.WriteLine($"第二轮工具调用数量: {finalChoice?.message?.tool_calls?.Length ?? 0}");

                        // 检查是否还有工具调用
                        if (finalChoice?.message?.tool_calls != null && finalChoice.message.tool_calls.Length > 0)
                        {
                            // 继续执行工具调用
                            var moreToolCalls = finalChoice.message.tool_calls;
                            System.Diagnostics.Debug.WriteLine($"继续执行 {moreToolCalls.Length} 个工具调用");
                            SafeUpdatePromptLabel($"正在执行 {moreToolCalls.Length} 个工具操作...");

                            // 将AI的工具调用消息加入历史
                            _chatHistory.Add(new ChatMessage
                            {
                                Role = "assistant",
                                Content = finalChoice.message.content,
                                ToolCalls = moreToolCalls.Select(tc => new ToolCall
                                {
                                    Id = tc.id,
                                    Type = tc.type,
                                    Function = new FunctionCall
                                    {
                                        Name = tc.function.name,
                                        Arguments = tc.function.arguments
                                    }
                                }).ToList()
                            });

                            // 执行每个工具调用
                            foreach (var toolCall in moreToolCalls)
                            {
                                var functionName = toolCall.function.name;
                                
                                // 检查是否为一次性工具且已执行过
                                if (_oneTimeTools.Contains(functionName) && _executedOneTimeTools.Contains(functionName))
                                {
                                    WriteLog("跳过重复工具", $"工具 {functionName} 已在本次请求中执行过，跳过重复执行");
                                    // 将跳过信息作为工具结果加入历史
                                    _chatHistory.Add(new ChatMessage
                                    {
                                        Role = "tool",
                                        Content = $"工具 {functionName} 已执行过，跳过重复调用",
                                        ToolCallId = toolCall.id
                                    });
                                    continue;
                                }
                                
                                var arguments = JsonSerializer.Deserialize<JsonElement>(toolCall.function.arguments);

                                System.Diagnostics.Debug.WriteLine($"执行工具: {functionName}");
                                System.Diagnostics.Debug.WriteLine($"参数: {toolCall.function.arguments}");
                                SafeUpdatePromptLabel($"正在执行工具: {functionName}...");

                                // 执行工具
                                var toolResult = ExecuteMcpTool(functionName, arguments);

                                System.Diagnostics.Debug.WriteLine($"工具执行结果: {toolResult}");
                                
                                // 记录一次性工具已执行
                                if (_oneTimeTools.Contains(functionName))
                                {
                                    _executedOneTimeTools.Add(functionName);
                                }

                                // 将工具结果加入历史
                                _chatHistory.Add(new ChatMessage
                                {
                                    Role = "tool",
                                    Content = toolResult,
                                    ToolCallId = toolCall.id
                                });
                            }

                            // 继续循环，再次调用API
                        }
                        else
                        {
                            // 没有更多工具调用，这是最终回复
                            var aiResponse = finalChoice?.message?.content?.Trim();

                            // 将最终AI回复加入历史
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
                }
                else
                {
                    // 没有工具调用，直接返回回复
                    var aiResponse = choice?.message?.content?.Trim();

                    // 将AI回复加入历史
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
        }

        // 处理Prompt Engineering模式的响应（用于不支持原生Function Calling的本地模型）
        private async Task<string> HandlePromptEngineeringResponse(HttpClient client, string apiUrl, string aiResponse, string userInput = null, int depth = 0, bool hasExecutedTools = false)
        {
            // 记录AI响应
            WriteLog("AI响应(Prompt Engineering)", $"递归深度: {depth}\n响应内容:\n{aiResponse}");

            // 限制递归深度，防止无限循环
            const int maxDepth = 3;
            if (depth >= maxDepth)
            {
                WriteLog("调试", "已达到最大递归深度，停止处理");
                _chatHistory.Add(new ChatMessage
                {
                    Role = "assistant",
                    Content = aiResponse
                });
                // 如果已经成功执行过工具，不显示警告
                if (hasExecutedTools)
                {
                    return aiResponse;
                }
                return aiResponse + "\n\n⚠️ [系统提示：已达到最大处理深度，停止继续处理。]";
            }

            // 解析响应中的工具调用
            var toolCalls = ParsePromptToolCalls(aiResponse);
            WriteLog("工具调用解析", $"解析到 {toolCalls.Count} 个工具调用");

            // 如果没有工具调用，检查是否模型错误地用文字描述了操作
            if (toolCalls.Count == 0)
            {
                WriteLog("调试", "未检测到工具调用");
                // 检测模型是否错误地用文字描述操作而没有输出工具调用
                bool seemsLikeFailedToolCall = aiResponse.Contains("已") && 
                    (aiResponse.Contains("写入") || aiResponse.Contains("设置") || aiResponse.Contains("创建") || 
                     aiResponse.Contains("删除") || aiResponse.Contains("保存") || aiResponse.Contains("完成"));
                
                if (seemsLikeFailedToolCall)
                {
                    WriteLog("警告", "模型似乎在描述操作但未输出工具调用格式");
                    // 模型似乎在描述操作但没有实际调用工具，添加提示
                    var warningResponse = aiResponse + "\n\n⚠️ [系统提示：当前本地模型未能正确输出工具调用格式，操作可能未实际执行。建议使用支持Function Calling的模型，或尝试更大参数的本地模型。]";
                    
                    _chatHistory.Add(new ChatMessage
                    {
                        Role = "assistant",
                        Content = aiResponse
                    });
                    return warningResponse;
                }
                
                // 将AI回复加入历史
                _chatHistory.Add(new ChatMessage
                {
                    Role = "assistant",
                    Content = aiResponse
                });
                return aiResponse;
            }

            // 记录解析到的工具调用详情
            var toolCallsDetail = new StringBuilder();
            foreach (var tc in toolCalls)
            {
                toolCallsDetail.AppendLine($"  - {tc.Name}: {tc.ArgumentsJson}");
            }
            WriteLog("工具调用详情", toolCallsDetail.ToString());

            // 获取纯文本内容（移除工具调用标签）
            string textContent = RemoveToolCallTags(aiResponse);

            System.Diagnostics.Debug.WriteLine($"Prompt Engineering模式：检测到 {toolCalls.Count} 个工具调用");
            SafeUpdatePromptLabel($"正在执行 {toolCalls.Count} 个工具操作...");

            // 将AI的响应（包含工具调用意图）加入历史
            _chatHistory.Add(new ChatMessage
            {
                Role = "assistant",
                Content = aiResponse
            });

            // 执行每个工具调用并收集结果
            var toolResults = new StringBuilder();
            foreach (var toolCall in toolCalls)
            {
                // 检查是否为一次性工具且已执行过
                if (_oneTimeTools.Contains(toolCall.Name) && _executedOneTimeTools.Contains(toolCall.Name))
                {
                    WriteLog("跳过重复工具", $"工具 {toolCall.Name} 已在本次请求中执行过，跳过重复执行");
                    toolResults.AppendLine($"工具 {toolCall.Name}: 已执行过，跳过重复调用");
                    continue;
                }

                System.Diagnostics.Debug.WriteLine($"执行工具: {toolCall.Name}");
                SafeUpdatePromptLabel($"正在执行工具: {toolCall.Name}...");

                try
                {
                    // 将JSON字符串解析为JsonElement
                    using (var argDoc = JsonDocument.Parse(toolCall.ArgumentsJson))
                    {
                        // 执行工具
                        var toolResult = ExecuteMcpTool(toolCall.Name, argDoc.RootElement);
                        System.Diagnostics.Debug.WriteLine($"工具执行结果: {toolResult}");
                        WriteLog("工具执行", $"工具: {toolCall.Name}\n参数: {toolCall.ArgumentsJson}\n结果: {toolResult}");

                        toolResults.AppendLine($"工具 {toolCall.Name} 执行结果: {toolResult}");
                        
                        // 记录一次性工具已执行
                        if (_oneTimeTools.Contains(toolCall.Name))
                        {
                            _executedOneTimeTools.Add(toolCall.Name);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"工具执行失败: {ex.Message}");
                    WriteLog("工具执行失败", $"工具: {toolCall.Name}\n参数: {toolCall.ArgumentsJson}\n错误: {ex.Message}");
                    toolResults.AppendLine($"工具 {toolCall.Name} 执行失败: {ex.Message}");
                }
            }

            // 将工具执行结果作为用户消息加入历史，让AI继续处理
            // 构建更清晰的结果消息，明确告知已完成的操作
            var toolResultMessage = new StringBuilder();
            toolResultMessage.AppendLine("工具执行完成，结果如下：");
            toolResultMessage.AppendLine(toolResults.ToString());
            
            // 如果执行了一次性工具，明确告知不要重复
            if (_executedOneTimeTools.Count > 0)
            {
                toolResultMessage.AppendLine($"⚠️ 以下工具已执行完成，请勿重复调用：{string.Join(", ", _executedOneTimeTools)}");
            }
            toolResultMessage.AppendLine("请根据执行结果用文字回复用户，不要再调用已执行的工具。");
            
            _chatHistory.Add(new ChatMessage
            {
                Role = "user",
                Content = toolResultMessage.ToString()
            });

            // 再次调用API获取最终回复
            SafeUpdatePromptLabel("等待AI响应...");

            var requestBody = new Dictionary<string, object>
            {
                { "model", _model },
                { "messages", BuildMessages(true, userInput) },
                { "temperature", 0.7 },
                { "max_tokens", 2000 }
            };

            // 仅对Ollama API添加特有参数
            if (!_isCloudConnection && _isOllamaApi)
            {
                requestBody["options"] = new Dictionary<string, object>
                {
                    { "num_predict", 1000 },
                    { "temperature", 0.7 }
                };
                requestBody["think"] = false;
            }

            var response = await client.PostAsJsonAsync(apiUrl, requestBody);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new HttpRequestException($"HTTP Error: {response.StatusCode}");
            }

            var jsonResponse = JsonSerializer.Deserialize<DeepSeekResponse>(responseContent);
            var choice = jsonResponse?.choices[0];
            var finalResponse = choice?.message?.content ?? "";

            // 检查是否还有更多工具调用
            var moreToolCalls = ParsePromptToolCalls(finalResponse);
            if (moreToolCalls.Count > 0)
            {
                // 递归处理更多工具调用，增加深度，标记已执行过工具
                return await HandlePromptEngineeringResponse(client, apiUrl, finalResponse, userInput, depth + 1, true);
            }

            // 将最终AI回复加入历史
            _chatHistory.Add(new ChatMessage
            {
                Role = "assistant",
                Content = finalResponse
            });

            return finalResponse;
        }

        // 构建消息列表（用于API请求）
        private List<object> BuildMessages(bool useMcp, string userInput = null)
        {
            var messages = new List<object>();

            // 添加系统提示词（仅在使用MCP时）
            if (useMcp && _excelMcp != null)
            {
                string systemPrompt;
                
                // 根据模式选择不同的系统提示词
                if (_usePromptEngineering)
                {
                    // Prompt Engineering模式：使用特殊格式的系统提示词，根据用户输入智能选择工具组
                    systemPrompt = GetPromptEngineeringSystemPrompt(userInput);
                }
                else
                {
                    // 获取当前选中单元格信息
                    string currentCell = "A1";
                    int currentRow = 1;
                    int currentCol = 1;
                    try
                    {
                        if (ThisAddIn.app?.Selection != null)
                        {
                            Microsoft.Office.Interop.Excel.Range selection = ThisAddIn.app.Selection;
                            currentCell = selection.Address.Replace("$", "");
                            currentRow = selection.Row;
                            currentCol = selection.Column;
                        }
                    }
                    catch { }
                    
                    string colLetter = GetColumnLetter(currentCol);
                    
                    // 原生Function Calling模式
                    systemPrompt = @"你是一个Excel操作助手。你必须通过调用工具来操作Excel文件。

**核心原则**：
🚫 **禁止仅用文字描述操作** - 例如：""我将在A1写入数据""、""现在我把名称写入A列""
✅ **必须实际调用工具函数** - 直接使用 set_cell_value、get_worksheet_names 等工具

**重要规则**：
1. **必须直接调用工具，不要只是描述要做什么**：
   - 错误示例：""我将在A1单元格写入xxx"" ❌
   - 错误示例：""现在我将这些工作表名称写入当前表的A列"" ❌
   - 正确示例：直接调用 set_cell_value 工具，参数为 row=行号, column=列号, value=用户指定的内容 ✅
   - 正确示例：循环调用 set_cell_value 工具，将每个工作表名称写入 A1、A2、A3... ✅

2. **对于需要多步操作的任务，必须调用多次工具**：
   - 例如：要将5个工作表名称写入A1-A5，必须调用5次 set_cell_value 工具
   - 第一次：set_cell_value(row=1, column=1, value=第一个表名)
   - 第二次：set_cell_value(row=2, column=1, value=第二个表名)
   - ...以此类推

3. **""表""默认指工作表（worksheet）**：
   - 当用户说""新建一个表""、""创建表""时，指的是在当前工作簿中创建新的工作表（sheet），而不是创建新的工作簿
   - 当用户说""新建工作簿""、""创建Excel文件""时，才是创建工作簿
   - 例如：""新建一个销售表"" → 使用 create_worksheet，而不是 create_workbook
   - **重要**：create_worksheet 默认会在工作簿的最前面（第一个位置）创建新工作表
   - 除非用户明确说明""在某张表后面/前面新建""，否则默认就是在最前面新建

4. **创建目录表的正确方式**：
   - 当用户要求创建目录表并写入表名时，注意行号分配：
   - 如果需要添加标题，标题应在A1，表名从A2开始
   - 例如：创建目录表 → 先在A1写入标题 → 表名从A2、A3、A4...开始写入
   - **错误做法**：标题在A1，第一个表名也在A1 ❌
   - **正确做法**：标题在A1（row=1），第一个表名在A2（row=2），第二个在A3（row=3）✅

5. **理解""当前单元格""的含义**：
   - 当用户说""当前单元格""、""选中的单元格""、""这个单元格""时，指的是用户在Excel中当前选中的单元格或区域
   - 当前选中单元格：" + currentCell + @"（行=" + currentRow + @", 列=" + currentCol + @"即" + colLetter + @"列）
   - 操作当前单元格时，直接使用 row=" + currentRow + @", column=" + currentCol + @"
   - 例如：""在当前单元格输入xxx"" → 调用 set_cell_value(row=" + currentRow + @", column=" + currentCol + @", value=用户指定的内容)

6. **区分两种超链接方式及其应用场景**：
   
   **A. HYPERLINK公式方式（set_hyperlink_formula）**：
   - 适用场景：工作簿内部跳转
   - 典型用途：
     * 跳转到同一工作簿的其他工作表
     * 创建目录页，链接到各个数据表
     * 在数据表中创建""返回目录""链接
   - 优点：在Excel内部打开，不会启动浏览器
   - 公式格式：=HYPERLINK(""#工作表名!单元格"", ""显示文本"")
   - 示例用法：
     * 用户说""在A1创建跳转到Sheet2的链接"" → 使用 set_hyperlink_formula
     * 用户说""创建目录，链接到各个工作表"" → 使用 set_hyperlink_formula
     * 用户说""在当前单元格添加返回首页的链接"" → 使用 set_hyperlink_formula
   
   **B. 超链接对象方式（add_hyperlink）**：
   - 适用场景：外部资源访问
   - 典型用途：
     * 打开网址（会启动默认浏览器）
     * 打开本地文件（Excel、Word、PDF等）
     * 打开网络共享文件
   - 优点：可以链接到任何外部资源
   - 示例用法：
     * 用户说""在A1添加某网站的链接"" → 使用 add_hyperlink
     * 用户说""链接到本地的报告文档"" → 使用 add_hyperlink
     * 用户说""添加公司网站链接"" → 使用 add_hyperlink
   
   **重要：如何选择**：
   - 如果目标是同一工作簿内的其他位置 → 使用 set_hyperlink_formula ✅
   - 如果目标是网址、本地文件、网络文件 → 使用 add_hyperlink ✅
   - 错误示例：用户说""跳转到Sheet2""却使用 add_hyperlink ❌
   - 正确示例：用户说""跳转到Sheet2""使用 set_hyperlink_formula ✅

7. 当用户说""当前工作簿""、""这个工作簿""、""当前表""、""这个表""时，指的是最近操作的工作簿和工作表

8. 当用户未明确指定工作簿名称时，使用当前活跃的工作簿

9. 当用户未明确指定工作表名称时，使用当前活跃的工作表

10. 通过上下文分析推断用户想要操作的对象

**当前环境**：
- 这是Excel插件环境，用户在Excel中打开了工作簿并启动了对话框
- 当前活跃工作簿（文件名）：" + (string.IsNullOrEmpty(_activeWorkbook) ? "无" : _activeWorkbook) + @"
- 当前活跃工作表（表名）：" + (string.IsNullOrEmpty(_activeWorksheet) ? "无" : _activeWorksheet) + @"
- 当前选中单元格：" + currentCell + @"（行=" + currentRow + @", 列=" + currentCol + @"）
- 注意：工作表名≠工作簿名！sheetName参数应填写工作表名（如""" + _activeWorksheet + @"""）

**重要提示**：
- 如果当前活跃工作簿为""无""，请先使用 get_current_excel_info 工具获取最新的Excel环境信息
- 获取信息后，你就能知道用户当前打开的工作簿和工作表，然后可以直接对其进行操作
- 不要只是告诉用户你将要做什么，必须实际调用工具来执行操作
- 每个操作都必须对应一个工具调用，不能省略
- value参数必须填写用户实际指定的内容，不要使用示例中的占位符

**操作流程示例**：
用户：""请将当前工作簿中所有表的名称写入当前表的A列""
正确做法：
1. 调用 get_worksheet_names 获取所有工作表名称
2. 对每个工作表名称，调用 set_cell_value(row=行号, column=1, value=实际表名)
3. 完成后告诉用户操作完成

错误做法：
只回复""现在我将这些工作表名称写入当前表的A列""但不调用任何工具 ❌

用户：""在所有表前新建一个目录表，写入所有表名，并加上超链接""
正确做法：
1. 调用 create_worksheet(sheetName=用户指定的表名) → 自动在最前面创建目录表
2. 调用 get_worksheet_names() → 获取所有表名
3. 调用 set_cell_value(row=1, column=1, value=标题内容) → 在A1写入标题
4. 对每个表名，调用 set_hyperlink_formula(cellAddress=对应单元格, targetLocation=表名!A1, displayText=表名) → 从A2开始
5. 告诉用户完成

**重要**：注意行号从2开始（跳过标题行A1），避免标题被覆盖

用户：""在当前单元格输入xxx""
正确做法：
1. 直接调用 set_cell_value(row=" + currentRow + @", column=" + currentCol + @", value=用户指定的内容)
2. 告诉用户操作完成

用户：""在A1创建跳转到某工作表的链接""
正确做法：
1. 调用 set_hyperlink_formula(cellAddress=""A1"", targetLocation=""目标表名!A1"", displayText=用户指定的显示文本)
2. 告诉用户已创建工作簿内部跳转链接

错误做法：
使用 add_hyperlink 添加外部链接 ❌（这会导致无法正确跳转）

用户：""在B2添加某网站的链接""
正确做法：
1. 调用 add_hyperlink(cellAddress=""B2"", url=用户指定的网址, displayText=用户指定的显示文本)
2. 告诉用户已添加外部网址链接

错误做法：
使用 set_hyperlink_formula ❌（这只适用于工作簿内部跳转）

用户：""根据河南省数据生成图表"" 或 ""将某某数据生成折线图/柱状图""
正确做法（必须按顺序执行）：
1. **先查找数据位置**：调用 find_value(searchValue=""河南省"") → 找到数据所在行/列
2. **再获取数据内容**：调用 get_range_values(rangeAddress=根据find_value结果确定的范围) → 获取完整数据
3. **最后创建图表**：调用 create_chart(dataRange=数据范围, chartType=图表类型, title=标题)
4. 告诉用户图表已创建，并简要分析数据

错误做法：
- 直接调用 create_chart 而不先查找和确认数据位置 ❌
- 假设数据在某个固定位置而不验证 ❌
- 只描述要创建图表但不调用工具 ❌

用户：""分析当前选中区域的数据并生成图表""
正确做法：
1. 调用 get_range_values(rangeAddress=""" + currentCell + @""") → 获取选中区域数据
2. 根据数据内容决定合适的图表类型
3. 调用 create_chart(dataRange=选中区域, chartType=合适的类型, title=描述性标题)
4. 分析数据并生成报告

**⚠️ 数据操作的核心原则：先查找，再操作**

当用户提到特定数据（如""河南省""、""销售额""、""2024年""等）时，必须遵循以下流程：

**11. 数据查找与定位规则**：
- **永远不要假设数据位置**：即使用户说""A列的数据""，也应先验证
- **使用 find_value 定位**：根据关键词找到数据的确切位置
- **使用 get_range_values 获取数据**：确认数据内容后再进行后续操作

**数据操作标准流程**：
| 操作类型 | 第一步 | 第二步 | 第三步 |
|---------|--------|--------|--------|
| 读取特定数据 | find_value 查找位置 | get_range_values 获取数据 | 返回结果给用户 |
| 分析数据 | find_value 查找位置 | get_range_values 获取数据 | 分析并生成报告 |
| 创建图表 | find_value 查找位置 | get_range_values 确认数据 | create_chart 创建图表 |
| 修改特定数据 | find_value 查找位置 | 确认目标单元格 | set_cell_value 修改 |
| 格式化特定区域 | find_value 查找位置 | 确定范围 | set_cell_format 等格式工具 |
| 排序/筛选 | find_value 查找表头 | 确定数据范围 | sort_range/set_auto_filter |
| 删除特定行/列 | find_value 查找位置 | 确认行号/列号 | delete_rows/delete_columns |

用户：""读取北京市的GDP数据"" 或 ""获取某某的销售额""
正确做法：
1. 调用 find_value(searchValue=""北京市"") → 找到数据位置
2. 根据返回的行列信息，调用 get_range_values 获取相关数据
3. 返回数据给用户

错误做法：
- 直接调用 get_range_values(""A1:D10"") 假设数据位置 ❌
- 不查找就直接读取 ❌

用户：""分析2020-2024年的收入变化""
正确做法：
1. 调用 find_value(searchValue=""2020"") → 找到年份数据起始位置
2. 调用 find_value(searchValue=""收入"") → 找到收入数据位置
3. 调用 get_range_values 获取完整数据范围
4. 分析数据趋势并生成报告

用户：""将河南省的数据标红"" 或 ""给某某数据加粗""
正确做法：
1. 调用 find_value(searchValue=""河南省"") → 找到数据位置
2. 根据返回的位置，调用 set_cell_format 设置格式

错误做法：
- 假设河南省在某行直接设置格式 ❌

用户：""删除空白行"" 或 ""删除包含某某的行""
正确做法：
1. 如果是删除特定内容的行，先调用 find_value 查找位置
2. 确认行号后调用 delete_rows

用户：""对销售数据进行排序""
正确做法：
1. 调用 find_value(searchValue=""销售"") → 找到销售数据列
2. 调用 get_range_values 确定数据范围
3. 调用 sort_range 进行排序

**特殊情况**：
- 如果用户明确指定了单元格地址（如""读取A1:D10的数据""），可以直接操作
- 如果用户说""当前选中区域""，使用当前选中单元格：" + currentCell + @"
- 如果 find_value 返回""未找到""，应告知用户并询问正确的关键词

请根据用户的自然语言指令，**立即调用**相应的工具完成任务，而不是仅仅描述你要做什么。";
                }

                messages.Add(new
                {
                    role = "system",
                    content = systemPrompt
                });
            }

            foreach (var msg in _chatHistory)
            {
                if (msg.Role == "tool")
                {
                    // 工具响应消息
                    messages.Add(new
                    {
                        role = "tool",
                        content = msg.Content,
                        tool_call_id = msg.ToolCallId
                    });
                }
                else if (msg.ToolCalls != null && msg.ToolCalls.Count > 0)
                {
                    // 带工具调用的助手消息
                    messages.Add(new
                    {
                        role = msg.Role,
                        content = msg.Content ?? "",
                        tool_calls = msg.ToolCalls.Select(tc => new
                        {
                            id = tc.Id,
                            type = tc.Type,
                            function = new
                            {
                                name = tc.Function.Name,
                                arguments = tc.Function.Arguments
                            }
                        }).ToArray()
                    });
                }
                else
                {
                    // 普通消息
                    // 在Prompt Engineering模式下，移除assistant消息中的<tool_calls>标签
                    string content = msg.Content;
                    if (_usePromptEngineering && msg.Role == "assistant" && content != null && content.Contains("<tool_calls>"))
                    {
                        content = RemoveToolCallTags(content);
                    }
                    
                    messages.Add(new
                    {
                        role = msg.Role,
                        content = content
                    });
                }
            }

            return messages;
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
                public ToolCallResponse[] tool_calls { get; set; }
            }

            public class ToolCallResponse
            {
                public string id { get; set; }
                public string type { get; set; }
                public FunctionResponse function { get; set; }
            }

            public class FunctionResponse
            {
                public string name { get; set; }
                public string arguments { get; set; }
            }
        }

        private void AddChatItem(string text, bool isUser)
        {
            // 暂停布局更新，避免闪烁
            flowLayoutPanelChat.SuspendLayout();
            
            try
            {
                int scrollBarWidth = SystemInformation.VerticalScrollBarWidth;
                int availableWidth = flowLayoutPanelChat.ClientSize.Width - scrollBarWidth - 20;
                // 最大宽度为容器宽度的75%
                int maxWidth = (int)(availableWidth * 0.75);
                int minWidth = 80; // 最小宽度
                int maxHeight = 300; // 最大高度，超过则显示滚动条
                int cornerRadius = 12; // 圆角半径
                int buttonPanelWidth = isUser ? 68 : 46; // 用户消息3个按钮，模型消息2个按钮
                int buttonHeight = 20; // 按钮高度

                // 先创建RichTextBox但不设置Text，避免触发布局
                RichTextBox richTextBox = new RichTextBox
                {
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    WordWrap = true,
                    Padding = new Padding(8),
                    ContextMenuStrip = CreateMessageContextMenu(isUser),
                    ScrollBars = RichTextBoxScrollBars.None
                };

                int finalWidth, finalHeight;
                bool needScroll = false;

                // 使用临时字体计算文本尺寸
                using (Graphics g = flowLayoutPanelChat.CreateGraphics())
                {
                    // 先计算单行文本的宽度
                    SizeF singleLineSize = g.MeasureString(text, richTextBox.Font);
                    int textWidth = (int)Math.Ceiling(singleLineSize.Width) + richTextBox.Padding.Horizontal + 10;

                    // 限制宽度在最小和最大之间
                    finalWidth = Math.Max(minWidth, Math.Min(textWidth, maxWidth));

                    // 根据最终宽度计算高度
                    SizeF textSize = g.MeasureString(text, richTextBox.Font, finalWidth - richTextBox.Padding.Horizontal);
                    int calculatedHeight = (int)Math.Ceiling(textSize.Height) + richTextBox.Padding.Vertical + 6;
                    
                    // 如果高度超过最大高度，启用滚动条
                    if (calculatedHeight > maxHeight)
                    {
                        finalHeight = maxHeight;
                        needScroll = true;
                    }
                    else
                    {
                        finalHeight = Math.Max(calculatedHeight, 30);
                    }
                }

                // 创建圆角对话框容器Panel
                Panel chatBubble = new Panel
                {
                    Size = new Size(finalWidth, finalHeight),
                    BackColor = isUser ? Color.LightBlue : Color.LightGreen,
                    Tag = isUser ? "user_container" : "model_container"
                };

                // 设置圆角
                System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                path.AddArc(0, 0, cornerRadius, cornerRadius, 180, 90);
                path.AddArc(chatBubble.Width - cornerRadius, 0, cornerRadius, cornerRadius, 270, 90);
                path.AddArc(chatBubble.Width - cornerRadius, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 0, 90);
                path.AddArc(0, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 90, 90);
                path.CloseAllFigures();
                chatBubble.Region = new Region(path);

                // 配置RichTextBox - 实现文本垂直居中
                int rtbWidth = finalWidth - 4;
                int rtbHeight = finalHeight - 4;
                
                // 计算实际文本高度，用于垂直居中
                int textTopPadding = 0;
                if (!needScroll)
                {
                    using (Graphics g = flowLayoutPanelChat.CreateGraphics())
                    {
                        SizeF textSize = g.MeasureString(text, richTextBox.Font, rtbWidth - richTextBox.Padding.Horizontal);
                        int actualTextHeight = (int)Math.Ceiling(textSize.Height) + richTextBox.Padding.Vertical;
                        if (actualTextHeight < rtbHeight)
                        {
                            textTopPadding = (rtbHeight - actualTextHeight) / 2;
                        }
                    }
                }
                
                richTextBox.Size = new Size(rtbWidth, rtbHeight);
                richTextBox.Location = new Point(2, 2);
                richTextBox.BackColor = chatBubble.BackColor;
                richTextBox.ScrollBars = needScroll ? RichTextBoxScrollBars.Vertical : RichTextBoxScrollBars.None;
                richTextBox.SelectionAlignment = HorizontalAlignment.Left;
                richTextBox.Tag = isUser ? "user_message" : "model_message";
                
                // 通过设置上边距实现垂直居中效果
                if (textTopPadding > 0)
                {
                    richTextBox.Padding = new Padding(8, 8 + textTopPadding, 8, 8);
                }
                
                // 最后设置文本，避免提前触发布局
                richTextBox.Text = text;

                chatBubble.Controls.Add(richTextBox);

                // 创建按钮面板
                Panel buttonPanel = new Panel
                {
                    Size = new Size(buttonPanelWidth, buttonHeight),
                    BackColor = Color.Transparent,
                    Tag = isUser ? "user_button_panel" : "model_button_panel"
                };

                // 创建按钮
                Button btn1, btn2, btn3 = null;
                ToolTip toolTip = new ToolTip();
                
                if (isUser)
                {
                    // 用户消息：编辑、重发、删除
                    btn1 = new Button
                    {
                        Text = "✎",
                        Size = new Size(20, 20),
                        Location = new Point(0, 0),
                        FlatStyle = FlatStyle.Flat,
                        Font = new Font("Segoe UI Symbol", 7),
                        Cursor = Cursors.Hand
                    };
                    btn1.FlatAppearance.BorderSize = 1;
                    btn1.Click += (s, e) => { richTextBoxInput.Text = text; richTextBoxInput.Focus(); richTextBoxInput.SelectAll(); };
                    toolTip.SetToolTip(btn1, "编辑");

                    btn2 = new Button
                    {
                        Text = "↻",
                        Size = new Size(20, 20),
                        Location = new Point(22, 0),
                        FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI Symbol", 7),
                    Cursor = Cursors.Hand
                };
                btn2.FlatAppearance.BorderSize = 1;
                btn2.Click += (s, e) => { richTextBoxInput.Text = text; send_button_Click(null, EventArgs.Empty); };
                toolTip.SetToolTip(btn2, "重发");

                btn3 = new Button
                {
                    Text = "🗑",
                    Size = new Size(20, 20),
                    Location = new Point(44, 0),
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI Symbol", 7),
                    Cursor = Cursors.Hand
                };
                btn3.FlatAppearance.BorderSize = 1;
                toolTip.SetToolTip(btn3, "删除");
            }
            else
            {
                btn1 = new Button
                {
                    Text = "📋",
                    Size = new Size(20, 20),
                    Location = new Point(0, 0),
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI Symbol", 7),
                    Cursor = Cursors.Hand
                };
                btn1.FlatAppearance.BorderSize = 1;
                btn1.Click += (s, e) => { Clipboard.SetText(text); prompt_label.Text = "已复制到剪贴板"; };
                toolTip.SetToolTip(btn1, "复制");

                btn2 = new Button
                {
                    Text = "🗑",
                    Size = new Size(20, 20),
                    Location = new Point(22, 0),
                    FlatStyle = FlatStyle.Flat,
                    Font = new Font("Segoe UI Symbol", 7),
                    Cursor = Cursors.Hand
                };
                btn2.FlatAppearance.BorderSize = 1;
                toolTip.SetToolTip(btn2, "删除");
            }
            buttonPanel.Controls.Add(btn1);
            buttonPanel.Controls.Add(btn2);
            if (btn3 != null) buttonPanel.Controls.Add(btn3);

            // 创建外层容器，包含按钮和对话框
            int rowHeight = Math.Max(finalHeight, buttonHeight);
            Panel rowPanel = new Panel
            {
                Size = new Size(availableWidth, rowHeight),
                BackColor = Color.Transparent,
                Tag = isUser ? "user_row" : "model_row"
            };

            // 按钮底部与对话框底部对齐
            int buttonTop = finalHeight - buttonHeight;
            if (buttonTop < 0) buttonTop = 0;

            if (isUser)
            {
                // 用户消息：对话框靠右，按钮在对话框左侧
                int chatBubbleLeft = availableWidth - finalWidth;
                chatBubble.Location = new Point(chatBubbleLeft, 0);
                buttonPanel.Location = new Point(chatBubbleLeft - buttonPanelWidth - 5, buttonTop);

                // 用户消息删除按钮事件
                btn3.Click += (s, e) =>
                {
                    flowLayoutPanelChat.Controls.Remove(rowPanel);
                    rowPanel.Dispose();
                };
            }
            else
            {
                // 模型消息：对话框靠左（X=0），按钮在对话框右侧
                chatBubble.Location = new Point(0, 0);
                buttonPanel.Location = new Point(finalWidth + 5, buttonTop);

                // 模型消息删除按钮事件
                btn2.Click += (s, e) =>
                {
                    flowLayoutPanelChat.Controls.Remove(rowPanel);
                    rowPanel.Dispose();
                };
            }

            rowPanel.Controls.Add(chatBubble);
            rowPanel.Controls.Add(buttonPanel);

            // 设置外层容器的边距 - 左边距固定为10，确保靠左显示
            rowPanel.Margin = new Padding(10, 5, 10, 10);
            flowLayoutPanelChat.Controls.Add(rowPanel);
            flowLayoutPanelChat.ScrollControlIntoView(rowPanel);
            }
            finally
            {
                // 恢复布局更新
                flowLayoutPanelChat.ResumeLayout(true);
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

        // 保存打开设置窗口前的配置，用于比较是否有变化
        private string _configBeforeSettings = string.Empty;

        private void settingsMenuItem_Click(object sender, EventArgs e)
        {
            // 保存当前配置状态
            _configBeforeSettings = $"{_apiKey}|{_model}|{_apiUrl}|{_enterMode}|{_isCloudConnection}";

            Form8 form8 = new Form8();
            form8.FormClosed += Form8_FormClosed;
            form8.ShowDialog();
        }

        private void Form8_FormClosed(object sender, FormClosedEventArgs e)
        {
            // 重新读取配置
            string oldApiKey = _apiKey;
            string oldModel = _model;
            string oldApiUrl = _apiUrl;
            string oldEnterMode = _enterMode;
            bool oldIsCloudConnection = _isCloudConnection;

            DecodeConfig();

            // 比较配置是否有变化
            string newConfig = $"{_apiKey}|{_model}|{_apiUrl}|{_enterMode}|{_isCloudConnection}";
            if (_configBeforeSettings != newConfig)
            {
                // 配置有变化，重新初始化
                // 根据用户勾选状态设置Prompt Engineering模式标志
                _usePromptEngineering = _isPromptEngineeringChecked;

                // 记录配置变化
                WriteLog("配置更新", $"模型: {_model}\nAPI地址: {_apiUrl}\n是否云端: {_isCloudConnection}\n是否Ollama: {_isOllamaApi}\n用户勾选'优先提示工程': {_isPromptEngineeringChecked}\nPrompt Engineering模式设置为: {_usePromptEngineering}");

                // 更新提示信息
                if (string.IsNullOrEmpty(_apiKey) && _isCloudConnection)
                {
                    prompt_label.Text = "请先进入设置配置API KEY";
                }
                else if (string.IsNullOrEmpty(_apiUrl))
                {
                    prompt_label.Text = "请先进入设置配置API地址";
                }
                else
                {
                    prompt_label.Text = "配置已更新，可以开始对话了！";
                }

                // 更新模型信息标签
                UpdateModelInfoLabel();

                // 根据连接类型设置"优先提示工程"复选框的状态
                UpdatePromptEngineeringCheckBoxState();
            }
            // 配置没有变化，不做任何操作
        }

        private void clearHistoryMenuItem_Click(object sender, EventArgs e)
        {
            // 清除对话历史
            _chatHistory.Clear();
            
            // 清除界面上的对话记录
            flowLayoutPanelChat.Controls.Clear();
            
            prompt_label.Text = "对话历史已清除";
        }

        private void exitMenuItem_Click(object sender, EventArgs e)
        {
            // 释放ExcelMcp资源
            if (_excelMcp != null)
            {
                _excelMcp.Dispose();
                _excelMcp = null;
            }
            this.Dispose();
        }

        private const string KeyFilePath = "encryption.key"; // 保存密钥和IV的文件路径
        private const string ConfigFilePath = "config.encrypted"; // 保存加密配置信息的文件路径

        //读取配置信息
        private void DecodeConfig()
        {
            if (!File.Exists(ConfigFilePath))
            {
                // 不在这里更新UI，只设置变量
                _apiKey = string.Empty;
                _model = string.Empty;
                _apiUrl = string.Empty;
                _enterMode = string.Empty;
                _isCloudConnection = true;
                _isOllamaApi = false;
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
                _apiKey = parts[0].Split('^')[1];
                _model = parts[1].Split('^')[1];
                _apiUrl = parts[2].Split('^')[1];
                _enterMode = parts[3].Split('^')[1];

                // 读取连接类型(如果配置文件中有的话)
                if (parts.Length >= 5)
                {
                    string connectionType = parts[4].Split('^')[1];
                    _isCloudConnection = (connectionType == "cloud");
                }
                else
                {
                    // 兼容旧配置，根据URL判断
                    _isCloudConnection = !IsLocalApiUrl(_apiUrl);
                }

                // 检测是否为Ollama API（通过端口或URL特征判断）
                _isOllamaApi = IsOllamaApi(_apiUrl);

                if (parts.Length >= 7)
                {
                    string timeoutMinutes = parts[6].Split('^')[1];
                    if (!string.IsNullOrEmpty(timeoutMinutes))
                    {
                        _timeoutMinutes = int.Parse(timeoutMinutes);
                    }
                    else
                    {
                        _timeoutMinutes = 5;
                    }
                }
                else
                {
                    _timeoutMinutes = 5;
                }

                // 不在这里更新UI
            }
            catch (Exception ex)
            {
                // 记录错误日志，不更新UI
                System.Diagnostics.Debug.WriteLine($"解密配置失败：{ex.Message}");
                _apiKey = string.Empty;
                _model = string.Empty;
                _apiUrl = string.Empty;
                _isCloudConnection = true;
                _isOllamaApi = false;
            }
        }

        // 检测是否为Ollama API
        private bool IsOllamaApi(string url)
        {
            try
            {
                Uri uri = new Uri(url);
                // Ollama默认端口是11434
                if (uri.Port == 11434)
                    return true;
                // 检查URL中是否包含ollama特征
                if (url.ToLower().Contains("ollama"))
                    return true;
                return false;
            }
            catch
            {
                return false;
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
                string[] hostParts = host.Split('.');
                if (hostParts.Length == 4 && hostParts[0] == "172")
                {
                    if (int.TryParse(hostParts[1], out int second))
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
            // 如果_enterMode为空，默认使用模式0（回车发送）
            string enterMode = string.IsNullOrEmpty(_enterMode) ? "0" : _enterMode;
            
            switch (enterMode)
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

        private void checkBoxPromptEngineering_CheckedChanged(object sender, EventArgs e)
        {
            _isPromptEngineeringChecked= checkBoxPromptEngineering.Checked;
        }

        // 根据连接类型设置checkboxPromptEngineering的状态
        private void UpdatePromptEngineeringCheckBoxState()
        {
            if (_isCloudConnection)
            {
                // 云端模型时，禁用"优先提示工程"复选框
                checkBoxPromptEngineering.Enabled = false;
                checkBoxPromptEngineering.Checked = false;
                _isPromptEngineeringChecked = false;
            }
            else
            {
                // 本地模型时，启用"优先提示工程"复选框
                checkBoxPromptEngineering.Enabled = true;
            }
        }
    }
}

