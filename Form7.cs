using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAddIn.Skills;



namespace ExcelAddIn
{
    public partial class Form7 : Form
    {

        private string _apiKey = string.Empty;           //api key变量
        private string _model = string.Empty;           //模型变量
        private string _apiUrl = string.Empty;         //api接口地址变量
        private string _enterMode = string.Empty;     //回车模式变量
        private bool _isCloudConnection = true;       //是否为云端连接（true=云端，false=本地）
        private bool _usePromptEngineering = false;   //是否使用Prompt Engineering模式（本地模型不支持function calling时自动启用）
        private bool _isOllamaApi = false;            //是否为Ollama API（用于添加Ollama特有参数）
        private int _timeoutMinutes = 5;            //请求超时时间，默认5分钟

        private bool _isStreamingChatItemCreated = false; // 标记流式聊天项是否已创建

        private ExcelMcp _excelMcp = null;  // Excel MCP实例
        private string _activeWorkbook = string.Empty;  // 当前活跃的工作簿
        private string _activeWorksheet = string.Empty;  // 当前活跃的工作表

        // 缓存Excel工具定义，避免重复创建
        private List<object> _cachedExcelTools = null;
        
        // 跟踪当前会话中已执行的一次性工具（如create_chart），防止递归时重复执行
        private HashSet<string> _executedOneTimeTools = new HashSet<string>();
        // 一次性工具列表（这些工具在一次用户请求中只应执行一次，但参数不同时可以重复执行）
        private static readonly HashSet<string> _oneTimeTools = new HashSet<string> 
        { 
            "create_chart", "create_table", "create_workbook", "create_worksheet", 
            "create_named_range", "save_workbook", "save_workbook_as" 
        };
        
        // 两阶段工具调用：是否启用工具分组模式（用于减少小模型的处理负担）
        private bool _useToolGrouping = true;
        
        // 性能优化相关变量
        private bool _hasSentDetailedSystemPrompt = false;  // 是否已发送详细系统提示词
        private bool _enableVerboseLogging = false;  // 是否启用详细日志（默认关闭以提升性能）
        
        // Skills系统相关
        private SkillManager _skillManager = null;  // 技能管理器
        
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
        // 性能优化：可通过 _enableVerboseLogging 控制日志详细程度
        private void WriteLog(string category, string message)
        {
            try
            {
                // 性能优化：如果关闭详细日志，则跳过某些日志记录
                if (!_enableVerboseLogging)
                {
                    // 跳过详细的请求体和响应体日志，只记录摘要
                    if (category.Contains("请求体") || category.Contains("响应体") || 
                        category.Contains("API请求") || category.Contains("API响应"))
                    {
                        // 只在调试模式下输出到调试窗口
                        System.Diagnostics.Debug.WriteLine($"[{category}] 已跳过详细日志记录（详细日志已关闭）");
                        return;
                    }
                }
                
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

            // 初始化Skills系统
            InitializeSkills();

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
            WriteLog("用户输入", $"内容: {userInput}\n当前模型: {_model}\nAPI地址: {_apiUrl}\n连接类型: {(_isCloudConnection ? "云端" : "本地")}\nPrompt Engineering模式: {_usePromptEngineering}");

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

                // 调用 AI API
                var response = await GetAIResponse(userInput);

                // 不再移除思考中占位符，而是直接在同一个容器中更新内容
                // 注意：如果流式响应创建了聊天项，则不需要处理思考占位符
                if (!_isStreamingChatItemCreated)
                {
                    // 将思考占位符转换为AI回复容器，并填入内容
                    ConvertThinkingPlaceholderToResponse(response);
                }
                else
                {
                    // 流式响应创建了聊天项，需要移除思考占位符
                    RemoveThinkingPlaceholder();
                }

                // 重置标记
                _isStreamingChatItemCreated = false;
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

        // 添加思考中占位符（创建一个可复用的对话气泡容器）
        private void AddThinkingPlaceholder()
        {
            flowLayoutPanelChat.SuspendLayout();
            try
            {
                int scrollBarWidth = SystemInformation.VerticalScrollBarWidth;
                int availableWidth = flowLayoutPanelChat.ClientSize.Width - scrollBarWidth - 20;
                int minWidth = 80;
                int cornerRadius = 12;

                // 创建RichTextBox（用于显示内容，后续可复用）
                RichTextBox richTextBox = new RichTextBox
                {
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    WordWrap = true,
                    Padding = new Padding(8),
                    ScrollBars = RichTextBoxScrollBars.None,
                    Text = "......",
                    Size = new Size(minWidth - 4, 26),
                    Location = new Point(2, 2),
                    BackColor = Color.LightGreen,
                    Font = new Font("微软雅黑", 12, FontStyle.Bold),
                    Tag = "thinking_content"
                };
                richTextBox.SelectAll();
                richTextBox.SelectionAlignment = HorizontalAlignment.Center;

                // 创建圆角对话框容器Panel
                Panel chatBubble = new Panel
                {
                    Size = new Size(minWidth, 30),
                    BackColor = Color.LightGreen,
                    Tag = "thinking_container",
                    Visible = false
                };

                // 设置圆角
                System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                path.AddArc(0, 0, cornerRadius, cornerRadius, 180, 90);
                path.AddArc(chatBubble.Width - cornerRadius, 0, cornerRadius, cornerRadius, 270, 90);
                path.AddArc(chatBubble.Width - cornerRadius, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 0, 90);
                path.AddArc(0, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 90, 90);
                path.CloseAllFigures();
                chatBubble.Region = new Region(path);

                chatBubble.Controls.Add(richTextBox);

                // 创建行容器
                Panel rowPanel = new Panel
                {
                    Size = new Size(availableWidth, 30),
                    BackColor = Color.Transparent,
                    Tag = "thinking_row"
                };

                chatBubble.Location = new Point(0, 0);
                rowPanel.Controls.Add(chatBubble);
                rowPanel.Margin = new Padding(10, 5, 10, 10);

                flowLayoutPanelChat.Controls.Add(rowPanel);
                flowLayoutPanelChat.ScrollControlIntoView(rowPanel);
                
                // 显示对话气泡
                chatBubble.Visible = true;

                _thinkingPlaceholder = rowPanel;
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

        // 流式响应相关类
        public class StreamResponse
        {
            public string id { get; set; }
            public string @object { get; set; }
            public int created { get; set; }
            public string model { get; set; }
            public StreamChoice[] choices { get; set; }
        }

        public class StreamChoice
        {
            public int index { get; set; }
            public StreamDelta delta { get; set; }
            public string finish_reason { get; set; }
        }

        public class StreamDelta
        {
            public string role { get; set; }
            public string content { get; set; }
            public StreamToolCall[] tool_calls { get; set; }
        }
        
        // 流式响应中的工具调用
        public class StreamToolCall
        {
            public int index { get; set; }
            public string id { get; set; }
            public string type { get; set; }
            public StreamFunctionCall function { get; set; }
        }
        
        public class StreamFunctionCall
        {
            public string name { get; set; }
            public string arguments { get; set; }
        }

        // 非流式响应的 Choice 和 Message 类（用于 Prompt Engineering 模式）
        public class AIChoice
        {
            public AIMessage message { get; set; }
        }

        public class AIMessage
        {
            public string role { get; set; }
            public string content { get; set; }
        }

        // 工具分组定义：组名 -> (关键词列表，工具列表)
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
                new[] { "工作表", "表名", "创建表", "新建表", "新表", "创建", "重命名", "删除表", "复制表", "移动表", "隐藏表", "显示表", "冻结", "取消冻结", "sheet", "切换", "激活", "跳转到", "添加表", "添加工作表" },
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
            sb.AppendLine("注意：你是Excel操作助手，不是Claude或其他AI模型。");
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
            bool wantsCreateSheet = inputLower.Contains("新建") && (inputLower.Contains("表") || inputLower.Contains("sheet") || inputLower.Contains("工作表")) ||
                                   inputLower.Contains("创建") && (inputLower.Contains("表") || inputLower.Contains("sheet") || inputLower.Contains("工作表")) ||
                                   inputLower.Contains("添加") && (inputLower.Contains("表") || inputLower.Contains("sheet") || inputLower.Contains("工作表"));
            
            if (wantsCreateSheet)
            {
                // 用户要创建新工作表
                sb.AppendLine("📋 创建工作表任务：");
                sb.AppendLine("使用 create_worksheet 工具创建新工作表：");
                sb.AppendLine("<tool_calls>");
                sb.AppendLine("[{\"name\": \"create_worksheet\", \"arguments\": {\"sheetName\": \"新工作表名称\"}}]");
                sb.AppendLine("</tool_calls>");
                sb.AppendLine();
                sb.AppendLine("注意：每次只创建一个工作表，等待结果后再创建下一个。");
            }
            else if (wantsChart && hasSelectedRange)
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
            var allTools = GetExcelTools();
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

            // 关键词映射（扩展版，覆盖更多自然语言表达）
            var keywordMap = new Dictionary<string, string[]>
            {
                ["cell_rw"] = new[] { 
                    "写入", "输入", "设置值", "读取", "获取", "单元格", "公式", "清除", "复制", "范围", "查找", "替换", "统计", "最后", "区域", "填写", "填充", "修改值", "更改值",
                    "写入数据", "填写数据", "输入数据", "修改数据", "更改数据", "设置内容", "修改内容", "更改内容",
                    "查看", "显示", "读取数据", "获取数据", "获取值", "读取值", "查看值",
                    "函数", "计算", "求和", "平均值", "最大值", "最小值", "计数", "sum", "average", "max", "min", "count",
                    "批量", "区域数据", "多行多列", "批量写入", "批量读取",
                    "填充区域", "写入区域", "读取区域", "获取区域",
                    "复制区域", "复制数据", "复制粘贴", "粘贴",
                    "清空", "清空数据", "清空区域", "清空单元格", "清空内容",
                    "删除内容", "清除内容", "删除数据", "清除数据"
                },
                ["format"] = new[] { 
                    "格式", "颜色", "字体", "背景", "加粗", "斜体", "边框", "合并", "对齐", "居中", "换行", "条件格式", "样式", "美化", 
                    "底纹", "填充色", "背景色", "底色", "字体颜色", "字体大小", "字号",
                    "画边框", "加边框", "设置边框", "边框线",
                    "合并单元格", "合并区域", "取消合并", "拆分单元格",
                    "自动换行", "设置换行",
                    "水平对齐", "垂直对齐", "左对齐", "右对齐", "居中对齐",
                    "下划线", "删除线", "粗体", "斜体字"
                },
                ["row_col"] = new[] { 
                    "行高", "列宽", "插入行", "插入列", "删除行", "删除列", "隐藏", "显示", "行", "列",
                    "调整行高", "调整列宽", "设置行高", "设置列宽", "改变行高", "改变列宽",
                    "插入一行", "插入一列", "删除一行", "删除一列", "添加行", "添加列",
                    "隐藏行", "隐藏列", "显示行", "显示列", "取消隐藏",
                    "行数", "列数", "多少行", "多少列"
                },
                ["sheet"] = new[] { 
                    "工作表", "表名", "创建表", "新建表", "新表", "重命名", "删除表", "复制表", "冻结", "sheet", 
                    "添加表", "增加表", "建表", "建一个表", "建几个表", "建三个表", "建多个表", "建2个表", "建两个表", "新建2个表", "新建两个表",
                    "创建新表", "创建工作表", "新建工作表", "添加工作表", "增加工作表",
                    "切换表", "激活表", "跳转表", "转到表", "选择表",
                    "重命名表", "改名", "修改表名", "更改表名",
                    "删除工作表", "移除表", "移除工作表",
                    "复制工作表", "拷贝表",
                    "移动表", "移动工作表", "调整表顺序",
                    "隐藏表", "显示表", "隐藏工作表", "显示工作表",
                    "冻结窗格", "冻结行", "冻结列", "取消冻结",
                    "标签", "工作表标签", "底部标签",
                    "表", "个工作表", "个表"
                },
                ["workbook"] = new[] { 
                    "工作簿", "文件", "打开", "保存", "关闭", "工作本",
                    "新建文件", "创建文件", "新建工作簿", "创建工作簿", "新建excel", "创建excel",
                    "打开文件", "打开工作簿", "打开excel",
                    "关闭文件", "关闭工作簿", "关闭excel",
                    "保存文件", "保存工作簿", "保存excel",
                    "另存为", "另存", "保存副本", "保存为",
                    "excel文件", "表格文件",
                    "新工作簿", "新文件", "创建新工作簿"
                },
                ["data"] = new[] { 
                    "排序", "筛选", "去重", "验证", "表格", "图表", "chart", 
                    "折线", "柱形", "饼图", "曲线", "柱状", "散点", "面积", "雷达", 
                    "生成图", "创建图", "画图", "可视化", "分析",
                    "升序", "降序", "从小到大", "从大到小", "排序方式",
                    "过滤", "筛选数据", "筛选条件",
                    "去重复", "删除重复", "重复值",
                    "数据验证", "数据有效性",
                    "创建图表", "生成图表", "画图表", "插入图表",
                    "柱状图", "折线图", "饼状图", "条形图", "面积图", "散点图", "雷达图",
                    "数据透视", "透视表", "pivot", "数据透视表", "创建透视表", "生成透视表", "插入透视表"
                },
                ["named"] = new[] { 
                    "命名区域", "命名范围", "命名单元格", "定义名称", "名称管理",
                    "命名", "定义名称", "创建名称", "设置名称"
                },
                ["link"] = new[] { 
                    "批注", "注释", "超链接", "链接", "跳转",
                    "添加批注", "添加注释", "插入批注", "插入注释",
                    "添加链接", "插入链接", "设置链接", "添加超链接"
                }
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
                            WriteLog("关键词匹配调试", $"关键词 '{keyword}' 匹配到工具组: {kv.Key}");
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

        // 获取Excel工具定义（带缓存优化）
        private List<object> GetExcelTools()
        {
            // 如果已缓存，直接返回
            if (_cachedExcelTools != null)
            {
                return _cachedExcelTools;
            }

            // 首次调用时从技能系统动态获取并缓存
            var toolsFromSkills = _skillManager.GetAllTools();
            _cachedExcelTools = new List<object>();

            foreach (var tool in toolsFromSkills)
            {
                // 构建正确的参数格式
                var parameters = tool.Parameters as Dictionary<string, object>;
                if (parameters != null && tool.RequiredParameters != null && tool.RequiredParameters.Count > 0)
                {
                    // 添加 required 字段到参数定义中
                    parameters["required"] = tool.RequiredParameters;
                }
                
                _cachedExcelTools.Add(new
                {
                    type = "function",
                    function = new
                    {
                        name = tool.Name,
                        description = tool.Description,
                        parameters = parameters
                    }
                });
            }

            return _cachedExcelTools;
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

        // 辅助方法：将 JsonElement 转换为 Dictionary<string, object>
        private Dictionary<string, object> JsonElementToDictionary(JsonElement element)
        {
            var dict = new Dictionary<string, object>();
            
            foreach (var prop in element.EnumerateObject())
            {
                dict[prop.Name] = JsonValueToObject(prop.Value);
            }
            
            return dict;
        }

        // 辅助方法：将 JsonValue 转换为 object
        private object JsonValueToObject(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.String:
                    return element.GetString();
                case JsonValueKind.Number:
                    if (element.TryGetInt32(out var intVal))
                        return intVal;
                    if (element.TryGetInt64(out var longVal))
                        return longVal;
                    return element.GetDouble();
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                case JsonValueKind.Null:
                    return null;
                case JsonValueKind.Array:
                    var list = new List<object>();
                    foreach (var item in element.EnumerateArray())
                    {
                        list.Add(JsonValueToObject(item));
                    }
                    return list;
                case JsonValueKind.Object:
                    var objDict = new Dictionary<string, object>();
                    foreach (var prop in element.EnumerateObject())
                    {
                        objDict[prop.Name] = JsonValueToObject(prop.Value);
                    }
                    return objDict;
                default:
                    return element.ToString();
            }
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

                // 直接使用技能系统执行所有工具
                var argumentsDict = JsonElementToDictionary(arguments);
                var result = _skillManager.ExecuteToolAsync(toolName, argumentsDict).Result;
                
                if (result.Success)
                {
                    return result.Content;
                }
                else
                {
                    return $"错误：{result.Error}";
                }
            }
            catch (AggregateException aex) when (aex.InnerException != null)
            {
                return $"执行工具 {toolName} 时出错：{aex.InnerException.Message}";
            }
            catch (Exception ex)
            {
                return $"执行工具 {toolName} 时出错：{ex.Message}";
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
        private async Task<string> GetAIResponse(string userInput)
        {
            string apiKey = _apiKey;
            string apiUrl = _apiUrl;
            bool useMcp = checkBoxUseMcp.Checked;

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
                
                // 添加流式输出参数
                requestBody["stream"] = true;

                // 如果启用MCP且ExcelMcp可用，且不是Prompt Engineering模式，添加工具定义
                if (useMcp && _excelMcp != null && !_usePromptEngineering)
                {
                    // 使用智能工具选择减少token数量
                    if (_useToolGrouping)
                    {
                        // 根据用户输入预选相关工具组
                        var preSelectedGroups = PreSelectToolGroups(userInput);
                        var selectedTools = GetToolsByGroups(preSelectedGroups);
                        requestBody["tools"] = selectedTools;
                        WriteLog("智能工具选择", $"根据用户输入预选工具组：[{string.Join(", ", preSelectedGroups)}], 工具数量：{selectedTools.Count}");
                    }
                    else
                    {
                        // 禁用分组时，发送全部工具
                        requestBody["tools"] = GetExcelTools();
                    }
                }

                // 记录请求信息（简化版，不包含完整工具定义）
                var requestJsonForLog = GetSimplifiedRequestBodyForLog(requestBody);
                WriteLog("API请求", $"URL: {apiUrl}\n模型: {_model}\nPrompt Engineering模式: {_usePromptEngineering}\n请求体:\n{requestJsonForLog}");

                // Kimi K2.5 模型要求 temperature=1，其他模型使用默认的 0.7
                if (_model.StartsWith("kimi-k2.5") || _model.StartsWith("kimi-k2"))
                {
                    requestBody["temperature"] = 1.0;
                }
                else
                {
                    requestBody["temperature"] = 0.7;
                }

                // 发送流式请求
                using (var request = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                {
                    request.Content = new StringContent(
                        JsonSerializer.Serialize(requestBody),
                        Encoding.UTF8,
                        "application/json"
                    );

                    // 只有云端连接时才添加Authorization头
                    if (_isCloudConnection && !string.IsNullOrEmpty(apiKey))
                    {
                        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                    }

                    using (var response = await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            var errorContent = await response.Content.ReadAsStringAsync();
                            WriteLog("API响应", $"状态码: {response.StatusCode}\n错误内容:\n{errorContent}");
                            
                            // 检查是否是因为不支持tools参数导致的错误
                            // 只有在真正不支持function calling时才切换到Prompt Engineering模式
                            // 需要更严格的判断：错误内容必须明确提到 tools/function 不支持
                            bool isToolsNotSupportedError = 
                                errorContent.Contains("tools is not supported") ||
                                errorContent.Contains("tool calls are not supported") ||
                                errorContent.Contains("function calling is not supported") ||
                                errorContent.Contains("does not support tools") ||
                                errorContent.Contains("unsupported parameter: tools") ||
                                errorContent.Contains("unknown parameter: tools") ||
                                errorContent.Contains("invalid parameter: tools");
                            
                            bool shouldSwitchToPromptEngineering = useMcp && !_usePromptEngineering && isToolsNotSupportedError;

                            WriteLog("模式检测", $"请求失败，状态码: {response.StatusCode}\n是否应切换到Prompt Engineering: {shouldSwitchToPromptEngineering}\n使用MCP={useMcp}, 当前非PE模式={!_usePromptEngineering}\n错误内容是否明确表示不支持tools: {isToolsNotSupportedError}");

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

                                // 重新发送请求
                                using (var retryRequest = new HttpRequestMessage(HttpMethod.Post, apiUrl))
                                {
                                    retryRequest.Content = new StringContent(
                                        JsonSerializer.Serialize(requestBody),
                                        Encoding.UTF8,
                                        "application/json"
                                    );

                                    if (_isCloudConnection && !string.IsNullOrEmpty(apiKey))
                                    {
                                        retryRequest.Headers.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                                    }

                                    using (var retryResponse = await client.SendAsync(retryRequest, HttpCompletionOption.ResponseHeadersRead))
                                    {
                                        if (!retryResponse.IsSuccessStatusCode)
                                        {
                                            var retryErrorContent = await retryResponse.Content.ReadAsStringAsync();
                                            throw new HttpRequestException($"HTTP Error: {retryResponse.StatusCode}, 响应: {retryErrorContent.Substring(0, Math.Min(200, retryErrorContent.Length))}");
                                        }

                                        return await HandleStreamingResponse(retryResponse);
                                    }
                                }
                            }
                            else
                            {
                                // 不是 tools 不支持的错误，抛出包含详细信息的异常
                                throw new HttpRequestException($"HTTP Error: {response.StatusCode}, 响应: {errorContent.Substring(0, Math.Min(500, errorContent.Length))}");
                            }
                        }

                        // 先获取流式响应内容
                        var streamingContent = await HandleStreamingResponse(response);
                        
                        // 如果有工具调用（无论是原生 Function Calling 还是 Prompt Engineering 模式），处理它们
                        if (useMcp && streamingContent.Contains("<tool_calls>"))
                        {
                            // 处理工具调用（使用外层的 client）
                            return await HandlePromptEngineeringResponse(client, apiUrl, streamingContent, userInput, 0, false);
                        }
                        
                        return streamingContent;
                    }
                }
            }
        }

        // 处理流式响应
        private async Task<string> HandleStreamingResponse(HttpResponseMessage response)
        {
            using (var stream = await response.Content.ReadAsStreamAsync())
            using (var reader = new StreamReader(stream))
            {
                var fullResponse = new StringBuilder();
                var toolCallsBuilder = new StringBuilder();
                string line;
                
                // 创建流式输出的聊天项
                Panel streamingChatItem = null;
                RichTextBox streamingRichTextBox = null;
                
                // 收集工具调用
                var collectedToolCalls = new Dictionary<int, ToolCallBuilder>();
                
                try
                {
                    while ((line = await reader.ReadLineAsync()) != null)
                    {
                        if (string.IsNullOrWhiteSpace(line))
                            continue;

                        // 移除SSE前缀
                        if (line.StartsWith("data: "))
                        {
                            line = line.Substring(6);
                        }

                        if (line == "[DONE]")
                            break;

                        try
                        {
                            var chunk = JsonSerializer.Deserialize<StreamResponse>(line);
                            
                            // 调试日志：记录原始响应
                            if (chunk == null)
                            {
                                WriteLog("流式响应调试", $"反序列化结果为null，原始行: {line}");
                                continue;
                            }
                            
                            if (chunk.choices == null || chunk.choices.Length == 0)
                            {
                                WriteLog("流式响应调试", $"choices为空，原始行: {line}");
                                continue;
                            }
                            
                            var choice = chunk.choices[0];
                            
                            // 处理文本内容
                            if (choice?.delta?.content != null)
                            {
                                string content = choice.delta.content;
                                fullResponse.Append(content);
                                
                                // 第一次收到内容时创建聊天项
                                if (streamingChatItem == null)
                                {
                                    // 创建流式聊天项
                                    streamingChatItem = CreateStreamingChatItem();
                                    streamingRichTextBox = (RichTextBox)streamingChatItem.Controls[0];
                                }
                                
                                // 更新聊天内容
                                UpdateStreamingChatItem(streamingRichTextBox, fullResponse.ToString());
                            }
                            
                            // 处理工具调用（原生 Function Calling 模式）
                            if (choice?.delta?.tool_calls != null)
                            {
                                foreach (var toolCall in choice.delta.tool_calls)
                                {
                                    int index = toolCall.index;
                                    
                                    if (!collectedToolCalls.ContainsKey(index))
                                    {
                                        collectedToolCalls[index] = new ToolCallBuilder
                                        {
                                            id = toolCall.id,
                                            type = toolCall.type,
                                            function = new FunctionCallBuilder()
                                        };
                                    }
                                    
                                    var builder = collectedToolCalls[index];
                                    
                                    if (!string.IsNullOrEmpty(toolCall.id))
                                    {
                                        builder.id = toolCall.id;
                                    }
                                    
                                    if (toolCall.function != null)
                                    {
                                        if (!string.IsNullOrEmpty(toolCall.function.name))
                                        {
                                            builder.function.name += toolCall.function.name;
                                        }
                                        if (!string.IsNullOrEmpty(toolCall.function.arguments))
                                        {
                                            builder.function.arguments += toolCall.function.arguments;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteLog("流式解析错误", $"解析流式响应失败: {ex.Message}\n原始行: {line}");
                        }
                    }
                }
                finally
                {
                    // 完成流式输出，移除思考占位符
                    RemoveThinkingPlaceholder();
                    
                    // 如果创建了流式聊天项，确保最终内容正确
                    if (streamingChatItem != null)
                    {
                        UpdateStreamingChatItem(streamingRichTextBox, fullResponse.ToString());
                        // 标记已经创建了聊天项，避免重复创建
                        _isStreamingChatItemCreated = true;
                    }
                }
                
                string finalResponse = fullResponse.ToString();
                WriteLog("流式响应完成", $"最终响应长度: {finalResponse.Length}字符");
                
                // 如果收集到了工具调用，处理它们
                if (collectedToolCalls.Count > 0)
                {
                    WriteLog("工具调用", $"检测到 {collectedToolCalls.Count} 个原生 Function Calling 工具调用");
                    
                    // 构建工具调用 JSON 格式（用于 Prompt Engineering 模式的解析）
                    var toolCallsArray = new List<object>();
                    foreach (var kv in collectedToolCalls.OrderBy(kv => kv.Key))
                    {
                        var builder = kv.Value;
                        
                        // 解析 arguments 字符串为 JSON 对象
                        object argumentsObj = new { };
                        try
                        {
                            if (!string.IsNullOrEmpty(builder.function.arguments))
                            {
                                // 尝试解析 arguments 字符串为 JsonElement
                                using (var argsDoc = JsonDocument.Parse(builder.function.arguments))
                                {
                                    // 将 JsonElement 转换为 .NET 对象
                                    argumentsObj = JsonElementToObject(argsDoc.RootElement);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            WriteLog("参数解析警告", $"无法解析参数 '{builder.function.arguments}': {ex.Message}");
                        }
                        
                        toolCallsArray.Add(new
                        {
                            name = builder.function.name,
                            arguments = argumentsObj
                        });
                    }
                    
                    var toolCallsJson = JsonSerializer.Serialize(toolCallsArray);
                    finalResponse = $"<tool_calls>\n{toolCallsJson}\n</tool_calls>";
                    
                    WriteLog("工具调用详情", $"工具调用 JSON: {toolCallsJson}");
                }
                
                return finalResponse;
            }
        }
        
        // 辅助类：用于构建工具调用
        private class ToolCallBuilder
        {
            public string id { get; set; }
            public string type { get; set; }
            public FunctionCallBuilder function { get; set; }
        }
        
        private class FunctionCallBuilder
        {
            public string name { get; set; }
            public string arguments { get; set; }
        }
        
        // 将 JsonElement 转换为 .NET 对象
        private object JsonElementToObject(JsonElement element)
        {
            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    var dict = new Dictionary<string, object>();
                    foreach (var prop in element.EnumerateObject())
                    {
                        dict[prop.Name] = JsonElementToObject(prop.Value);
                    }
                    return dict;
                case JsonValueKind.Array:
                    var list = new List<object>();
                    foreach (var item in element.EnumerateArray())
                    {
                        list.Add(JsonElementToObject(item));
                    }
                    return list;
                case JsonValueKind.String:
                    return element.GetString();
                case JsonValueKind.Number:
                    if (element.TryGetInt32(out int intVal))
                        return intVal;
                    if (element.TryGetInt64(out long longVal))
                        return longVal;
                    if (element.TryGetDouble(out double doubleVal))
                        return doubleVal;
                    return element.GetDouble();
                case JsonValueKind.True:
                    return true;
                case JsonValueKind.False:
                    return false;
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;
                default:
                    return element.ToString();
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
                // 检查是否为一次性工具且已执行过（使用工具名称+参数作为唯一标识）
                string toolCallKey = $"{toolCall.Name}:{toolCall.ArgumentsJson}";
                if (_oneTimeTools.Contains(toolCall.Name) && _executedOneTimeTools.Contains(toolCallKey))
                {
                    WriteLog("跳过重复工具", $"工具 {toolCall.Name} 参数 {toolCall.ArgumentsJson} 已在本次请求中执行过，跳过重复执行");
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
                        
                        // 记录一次性工具已执行（使用工具名称+参数作为唯一标识）
                        if (_oneTimeTools.Contains(toolCall.Name))
                        {
                            _executedOneTimeTools.Add(toolCallKey);
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

            // Kimi K2.5 模型要求 temperature=1
            double temperature = (_model.StartsWith("kimi-k2.5") || _model.StartsWith("kimi-k2")) ? 1.0 : 0.7;

            var requestBody = new Dictionary<string, object>
            {
                { "model", _model },
                { "messages", BuildMessages(true, userInput) },
                { "temperature", temperature },
                { "max_tokens", 2000 }
            };

            // 仅对Ollama API添加特有参数
            if (!_isCloudConnection && _isOllamaApi)
            {
                requestBody["options"] = new Dictionary<string, object>
                {
                    { "num_predict", 1000 },
                    { "temperature", temperature }
                };
                requestBody["think"] = false;
            }

            var requestContent = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");
            var response = await client.PostAsync(apiUrl, requestContent);
            var responseContent = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new HttpRequestException($"HTTP Error: {response.StatusCode}");
            }

            var jsonResponse = JsonSerializer.Deserialize<AIResponse>(responseContent);
            var finalChoice = jsonResponse?.choices != null && jsonResponse.choices.Length > 0 ? jsonResponse.choices[0] : null;
            var finalResponse = finalChoice?.message?.content ?? "";

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

            // 在UI线程上添加最终回复到聊天界面
            if (!string.IsNullOrWhiteSpace(finalResponse))
            {
                this.Invoke(new Action(() =>
                {
                    AddChatItem(finalResponse, false);
                }));
            }

            return finalResponse;
        }

        // 获取紧凑版系统提示词（用于提升性能）
        private string GetCompactSystemPrompt()
        {
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
            
            return $@"你是Excel操作助手，必须通过调用工具来操作Excel。
注意：你是Excel操作助手，不是Claude或其他AI模型。

核心规则：
1. 禁止仅用文字描述操作，必须实际调用工具函数
2. 对于多步操作，必须调用多次工具
3. 每次只输出一个工具调用，等待结果后再继续

当前环境：
- 当前工作簿：{(_activeWorkbook ?? "无")}
- 当前工作表：{(_activeWorksheet ?? "无")}
- 当前选中单元格：{currentCell}（行={currentRow}, 列={currentCol}即{colLetter}列）

重要提示：
- 当用户说""当前单元格""、""选中的单元格""时，指的是 {currentCell}
- 当用户说""当前表""、""这个表""时，指的是 {_activeWorksheet}
- 不要只是告诉用户你将要做什么，必须实际调用工具来执行操作";
        }

        // 获取详细版系统提示词（首次请求使用，包含完整规则说明）
        private string GetDetailedSystemPrompt(string currentCell, int currentRow, int currentCol, string colLetter)
        {
            return $@"你是一个Excel操作助手。你必须通过调用工具来操作Excel文件。
注意：你是Excel操作助手，不是Claude或其他AI模型。

**核心原则**：
🚫 **禁止仅用文字描述操作** - 例如：""我将在A1写入数据""、""现在我把名称写入A列""
✅ **必须实际调用工具函数** - 直接使用 set_cell_value、get_worksheet_names 等工具

**重要规则**：
1. **必须直接调用工具，不要只是描述要做什么**
2. **对于需要多步操作的任务，必须调用多次工具**
3. **""表""默认指工作表（worksheet）**
4. **理解""当前单元格""的含义**：当前选中单元格：{currentCell}（行={currentRow}, 列={currentCol}即{colLetter}列）

**当前环境**：
- 当前工作簿：{(_activeWorkbook ?? "无")}
- 当前工作表：{(_activeWorksheet ?? "无")}
- 当前选中单元格：{currentCell}

**重要提示**：
- 不要只是告诉用户你将要做什么，必须实际调用工具来执行操作
- 每个操作都必须对应一个工具调用，不能省略
- value参数必须填写用户实际指定的内容，不要使用示例中的占位符

请根据用户的自然语言指令，**立即调用**相应的工具完成任务。";
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
                    // 原生Function Calling模式
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
                    
                    // 性能优化：首次请求使用详细提示词，后续请求使用紧凑提示词
                    if (!_hasSentDetailedSystemPrompt)
                    {
                        // 首次请求：使用详细系统提示词
                        systemPrompt = GetDetailedSystemPrompt(currentCell, currentRow, currentCol, colLetter);
                        _hasSentDetailedSystemPrompt = true;
                    }
                    else
                    {
                        // 后续请求：使用紧凑系统提示词以提升性能
                        systemPrompt = GetCompactSystemPrompt();
                    }
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

        // AI API 响应模型
        public class AIResponse
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
                    Tag = isUser ? "user_container" : "model_container",
                    Visible = false  // 先隐藏，等所有内容准备好后再显示
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
            
            // 所有内容准备好后，显示对话气泡
            chatBubble.Visible = true;
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
            
            // 复制菜单项
            ToolStripMenuItem copyItem = new ToolStripMenuItem("复制");
            copyItem.Click += (s, e) =>
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
            menu.Items.Add(copyItem);
            
            // 删除菜单项（仅用户消息）
            if (isUserMessage)
            {
                ToolStripMenuItem deleteItem = new ToolStripMenuItem("删除");
                deleteItem.Click += (s, e) =>
                {
                    if (menu.SourceControl is RichTextBox rtb && flowLayoutPanelChat.Controls.Contains(rtb))
                    {
                        flowLayoutPanelChat.Controls.Remove(rtb);
                        rtb.Dispose();
                    }
                };
                menu.Items.Add(deleteItem);
            }
            
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

                // 记录配置变化
                WriteLog("配置更新", $"模型: {_model}\nAPI地址: {_apiUrl}\n是否云端: {_isCloudConnection}\n是否Ollama: {_isOllamaApi}");

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

                // 特殊处理：如果模型名称包含 :cloud 后缀，则认为是云端模型
                if (_model.Contains(":cloud") || _model.Contains(":Cloud"))
                {
                    _isCloudConnection = true;
                    WriteLog("模型检测", $"检测到云端模型后缀 ':cloud'，设置为云端连接");
                }

                // 读取模型提供商(如果配置文件中有的话)
                string provider = "";
                if (parts.Length >= 8)
                {
                    provider = parts[7].Split('^')[1];
                }

                // 检测是否为Ollama API（通过端口或URL特征判断）
                _isOllamaApi = IsOllamaApi(_apiUrl) || provider == "Ollama";

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

        // 初始化Skills系统
        private void InitializeSkills()
        {
            _skillManager = new SkillManager();
            
            if (_excelMcp != null)
            {
                // 加载内置技能
                _skillManager.LoadSkill(new ExcelBaseSkill(_excelMcp));
                _skillManager.LoadSkill(new ExcelWorkbookSkill(_excelMcp));
                _skillManager.LoadSkill(new ExcelCellSkill(_excelMcp));
                _skillManager.LoadSkill(new ExcelAnalysisSkill(_excelMcp));
                _skillManager.LoadSkill(new ExcelFinanceSkill(_excelMcp));
                _skillManager.LoadSkill(new ExcelFormatSkill(_excelMcp));
                _skillManager.LoadSkill(new ExcelSheetSkill(_excelMcp));
                _skillManager.LoadSkill(new ExcelRangeSkill(_excelMcp));
                _skillManager.LoadSkill(new ExcelChartSkill(_excelMcp));
                _skillManager.LoadSkill(new ExcelPivotSkill(_excelMcp));
                
                // 加载插件目录中的技能
                var skillsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "skills");
                var pluginSkills = PluginLoader.LoadSkillsFromDirectory(skillsDir, _excelMcp);
                foreach (var skill in pluginSkills)
                {
                    _skillManager.LoadSkill(skill);
                }
                
                // 记录加载的技能
                var loadedSkills = _skillManager.GetLoadedSkills();
                WriteLog("Skills初始化", $"加载了 {loadedSkills.Count} 个技能: {string.Join(", ", loadedSkills.Select(s => s.Name))}");
            }
        }

        // 创建流式输出的聊天项
        private Panel CreateStreamingChatItem()
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
                int cornerRadius = 12; // 圆角半径

                // 创建RichTextBox
                RichTextBox richTextBox = new RichTextBox
                {
                    BorderStyle = BorderStyle.None,
                    ReadOnly = true,
                    WordWrap = true,
                    Padding = new Padding(8),
                    ContextMenuStrip = CreateMessageContextMenu(false),
                    ScrollBars = RichTextBoxScrollBars.None,
                    Text = ""
                };

                // 创建圆角对话框容器Panel
                Panel chatBubble = new Panel
                {
                    Size = new Size(minWidth, 30),
                    BackColor = Color.LightGreen,
                    Tag = "model_container",
                    Visible = false  // 先隐藏，等有内容后再显示
                };

                // 设置圆角
                System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                path.AddArc(0, 0, cornerRadius, cornerRadius, 180, 90);
                path.AddArc(chatBubble.Width - cornerRadius, 0, cornerRadius, cornerRadius, 270, 90);
                path.AddArc(chatBubble.Width - cornerRadius, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 0, 90);
                path.AddArc(0, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 90, 90);
                path.CloseAllFigures();
                chatBubble.Region = new Region(path);

                // 配置RichTextBox
                int rtbWidth = chatBubble.Width - 4;
                int rtbHeight = chatBubble.Height - 4;
                
                richTextBox.Size = new Size(rtbWidth, rtbHeight);
                richTextBox.Location = new Point(2, 2);
                richTextBox.BackColor = chatBubble.BackColor;
                richTextBox.SelectionAlignment = HorizontalAlignment.Left;
                richTextBox.Tag = "model_message";

                chatBubble.Controls.Add(richTextBox);

                // 添加到聊天面板
                flowLayoutPanelChat.Controls.Add(chatBubble);
                flowLayoutPanelChat.ResumeLayout(false);
                flowLayoutPanelChat.PerformLayout();
                
                // 滚动到底部
                flowLayoutPanelChat.ScrollControlIntoView(chatBubble);
                
                return chatBubble;
            }
            catch
            {
                flowLayoutPanelChat.ResumeLayout(false);
                throw;
            }
        }

        // 更新流式聊天项内容
        private void UpdateStreamingChatItem(RichTextBox richTextBox, string content)
        {
            if (richTextBox.InvokeRequired)
            {
                richTextBox.Invoke(new Action<RichTextBox, string>(UpdateStreamingChatItem), richTextBox, content);
                return;
            }

            try
            {
                // 暂停布局更新
                flowLayoutPanelChat.SuspendLayout();
                
                // 更新内容
                richTextBox.Text = content;
                
                // 调整大小
                Panel chatBubble = (Panel)richTextBox.Parent;
                
                // 如果对话气泡还不可见，先显示它
                if (!chatBubble.Visible)
                {
                    chatBubble.Visible = true;
                }
                
                int scrollBarWidth = SystemInformation.VerticalScrollBarWidth;
                int availableWidth = flowLayoutPanelChat.ClientSize.Width - scrollBarWidth - 20;
                int maxWidth = (int)(availableWidth * 0.75);
                int minWidth = 80;
                int maxHeight = 300;
                
                using (Graphics g = flowLayoutPanelChat.CreateGraphics())
                {
                    // 计算文本尺寸
                    SizeF textSize = g.MeasureString(content, richTextBox.Font, maxWidth - richTextBox.Padding.Horizontal);
                    int textWidth = (int)Math.Ceiling(textSize.Width) + richTextBox.Padding.Horizontal + 10;
                    int textHeight = (int)Math.Ceiling(textSize.Height) + richTextBox.Padding.Vertical + 6;
                    
                    // 限制宽度
                    int finalWidth = Math.Max(minWidth, Math.Min(textWidth, maxWidth));
                    
                    // 限制高度
                    int finalHeight = Math.Min(textHeight, maxHeight);
                    bool needScroll = textHeight > maxHeight;
                    
                    // 更新RichTextBox大小
                    richTextBox.Size = new Size(finalWidth - 4, finalHeight - 4);
                    richTextBox.ScrollBars = needScroll ? RichTextBoxScrollBars.Vertical : RichTextBoxScrollBars.None;
                    
                    // 更新聊天气泡大小
                    chatBubble.Size = new Size(finalWidth, finalHeight);
                    
                    // 重新设置圆角
                    int cornerRadius = 12;
                    System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                    path.AddArc(0, 0, cornerRadius, cornerRadius, 180, 90);
                    path.AddArc(chatBubble.Width - cornerRadius, 0, cornerRadius, cornerRadius, 270, 90);
                    path.AddArc(chatBubble.Width - cornerRadius, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 0, 90);
                    path.AddArc(0, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 90, 90);
                    path.CloseAllFigures();
                    chatBubble.Region = new Region(path);
                }
                
                // 恢复布局并滚动到底部
                flowLayoutPanelChat.ResumeLayout(false);
                flowLayoutPanelChat.PerformLayout();
                flowLayoutPanelChat.ScrollControlIntoView(chatBubble);
            }
            catch (Exception ex)
            {
                WriteLog("流式更新错误", $"更新流式聊天项失败: {ex.Message}");
            }
        }

        // 移除思考占位符（如果存在）
        private void RemoveThinkingPlaceholder()
        {
            if (flowLayoutPanelChat.InvokeRequired)
            {
                flowLayoutPanelChat.Invoke(new Action(RemoveThinkingPlaceholder));
                return;
            }

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

        // 将思考占位符转换为AI回复容器，并填入内容
        private void ConvertThinkingPlaceholderToResponse(string response)
        {
            if (_thinkingPlaceholder == null)
                return;

            flowLayoutPanelChat.SuspendLayout();
            try
            {
                // 获取思考占位符中的控件
                Panel chatBubble = _thinkingPlaceholder.Controls[0] as Panel;
                if (chatBubble != null)
                {
                    RichTextBox richTextBox = chatBubble.Controls[0] as RichTextBox;
                    if (richTextBox != null)
                    {
                        // 重置字体和对齐方式
                        richTextBox.Font = new Font("微软雅黑", 9, FontStyle.Regular);
                        richTextBox.SelectionAlignment = HorizontalAlignment.Left;
                        // 更新Tag
                        richTextBox.Tag = "model_message";
                        
                        // 计算文本尺寸
                        int scrollBarWidth = SystemInformation.VerticalScrollBarWidth;
                        int availableWidth = flowLayoutPanelChat.ClientSize.Width - scrollBarWidth - 20;
                        int maxWidth = (int)(availableWidth * 0.75);
                        int minWidth = 80;
                        int maxHeight = 300;
                        
                        using (Graphics g = flowLayoutPanelChat.CreateGraphics())
                        {
                            SizeF textSize = g.MeasureString(response, richTextBox.Font, maxWidth - richTextBox.Padding.Horizontal);
                            int textWidth = (int)Math.Ceiling(textSize.Width) + richTextBox.Padding.Horizontal + 10;
                            int textHeight = (int)Math.Ceiling(textSize.Height) + richTextBox.Padding.Vertical + 6;
                            
                            int finalWidth = Math.Max(minWidth, Math.Min(textWidth, maxWidth));
                            int finalHeight = Math.Min(textHeight, maxHeight);
                            bool needScroll = textHeight > maxHeight;
                            
                            // 更新RichTextBox大小
                            richTextBox.Size = new Size(finalWidth - 4, finalHeight - 4);
                            richTextBox.ScrollBars = needScroll ? RichTextBoxScrollBars.Vertical : RichTextBoxScrollBars.None;
                            
                            // 更新聊天气泡大小
                            chatBubble.Size = new Size(finalWidth, finalHeight);
                            
                            // 重新设置圆角
                            int cornerRadius = 12;
                            System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                            path.AddArc(0, 0, cornerRadius, cornerRadius, 180, 90);
                            path.AddArc(chatBubble.Width - cornerRadius, 0, cornerRadius, cornerRadius, 270, 90);
                            path.AddArc(chatBubble.Width - cornerRadius, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 0, 90);
                            path.AddArc(0, chatBubble.Height - cornerRadius, cornerRadius, cornerRadius, 90, 90);
                            path.CloseAllFigures();
                            chatBubble.Region = new Region(path);
                            
                            // 更新行容器大小
                            _thinkingPlaceholder.Size = new Size(availableWidth, finalHeight);
                        }
                        
                        // 设置内容
                        richTextBox.Text = response;
                    }
                    // 更新Tag
                    chatBubble.Tag = "model_container";
                }
                // 更新Tag
                _thinkingPlaceholder.Tag = "model_row";
                
                // 滚动到底部
                flowLayoutPanelChat.ScrollControlIntoView(_thinkingPlaceholder);
            }
            finally
            {
                flowLayoutPanelChat.ResumeLayout(true);
            }
        }

        private void prompt_label_DoubleClick(object sender, EventArgs e)
        {
            Clipboard.SetText(prompt_label.Text);
            MessageBox.Show("文本已复制到剪贴板", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}

