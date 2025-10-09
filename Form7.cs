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

        private ExcelMcp _excelMcp = null;  // Excel MCP实例
        private string _activeWorkbook = string.Empty;  // 当前活跃的工作簿
        private string _activeWorksheet = string.Empty;  // 当前活跃的工作表

        // 缓存MCP工具定义，避免重复创建
        private List<object> _cachedMcpTools = null;

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
            if (string.IsNullOrEmpty(_apiKey) || string.IsNullOrEmpty(_model))
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
                                chartType = new { type = "string", description = "图表类型：line/bar/column/pie/scatter/area/radar" },
                                dataRange = new { type = "string", description = "数据源范围（如'A1:D10'）" },
                                chartPosition = new { type = "string", description = "图表位置（如'F1'）" },
                                title = new { type = "string", description = "图表标题（可选）" },
                                width = new { type = "integer", description = "图表宽度（默认400）" },
                                height = new { type = "integer", description = "图表高度（默认300）" }
                            },
                            required = new[] { "chartType", "dataRange", "chartPosition" }
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

        // 执行MCP工具调用
        private string ExecuteMcpTool(string toolName, JsonElement arguments)
        {
            try
            {
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
                            var value = arguments.GetProperty("value").GetString();

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
                            var dataJson = arguments.GetProperty("data").GetString();

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var range = worksheet.Range[rangeAddress];

                            // 解析JSON数组为二维数组
                            var dataList = JsonSerializer.Deserialize<List<List<object>>>(dataJson);
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
                            var rangeAddress = arguments.GetProperty("rangeAddress").GetString();
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
                            var chartType = arguments.GetProperty("chartType").GetString();
                            var dataRange = arguments.GetProperty("dataRange").GetString();
                            var chartPosition = arguments.GetProperty("chartPosition").GetString();
                            var title = arguments.TryGetProperty("title", out var titleProp) ? titleProp.GetString() : null;
                            var width = arguments.TryGetProperty("width", out var widthProp) ? widthProp.GetInt32() : 400;
                            var height = arguments.TryGetProperty("height", out var heightProp) ? heightProp.GetInt32() : 300;

                            var worksheet = GetWorksheet(fileName, sheetName);
                            var dataRangeObj = worksheet.Range[dataRange];
                            var chartPositionObj = worksheet.Range[chartPosition];

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

                // 构建请求体
                var requestBody = new Dictionary<string, object>
                {
                    { "model", _model },
                    { "messages", BuildMessages(useMcp) },
                    { "temperature", 0.7 },
                    { "max_tokens", 2000 }
                };

                // 如果启用MCP且ExcelMcp可用，添加工具定义
                if (useMcp && _excelMcp != null)
                {
                    requestBody["tools"] = GetMcpTools();
                }

                var response = await client.PostAsJsonAsync(apiUrl, requestBody);
                var responseContent = await response.Content.ReadAsStringAsync();

                if (!response.IsSuccessStatusCode)
                {
                    throw new HttpRequestException($"HTTP Error: {response.StatusCode}");
                }

                var jsonResponse = JsonSerializer.Deserialize<DeepSeekResponse>(responseContent);
                var choice = jsonResponse?.choices[0];

                // 调试信息：检查是否有工具调用
                System.Diagnostics.Debug.WriteLine($"AI响应内容: {choice?.message?.content}");
                System.Diagnostics.Debug.WriteLine($"工具调用数量: {choice?.message?.tool_calls?.Length ?? 0}");
                System.Diagnostics.Debug.WriteLine($"完整响应: {responseContent}");

                // 检查是否有工具调用
                if (choice?.message?.tool_calls != null && choice.message.tool_calls.Length > 0)
                {
                    // 处理工具调用
                    var toolCalls = choice.message.tool_calls;

                    System.Diagnostics.Debug.WriteLine($"开始执行 {toolCalls.Length} 个工具调用");
                    prompt_label.Text = $"正在执行 {toolCalls.Length} 个工具操作...";

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
                        var arguments = JsonSerializer.Deserialize<JsonElement>(toolCall.function.arguments);

                        System.Diagnostics.Debug.WriteLine($"执行工具: {functionName}");
                        System.Diagnostics.Debug.WriteLine($"参数: {toolCall.function.arguments}");
                        prompt_label.Text = $"正在执行工具: {functionName}...";

                        // 执行工具
                        var toolResult = ExecuteMcpTool(functionName, arguments);

                        System.Diagnostics.Debug.WriteLine($"工具执行结果: {toolResult}");

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
                            { "messages", BuildMessages(useMcp) },
                            { "temperature", 0.7 },
                            { "max_tokens", 2000 }
                        };

                        if (useMcp && _excelMcp != null)
                        {
                            finalRequestBody["tools"] = GetMcpTools();
                        }

                        prompt_label.Text = "等待AI响应...";
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
                            prompt_label.Text = $"正在执行 {moreToolCalls.Length} 个工具操作...";

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
                                var arguments = JsonSerializer.Deserialize<JsonElement>(toolCall.function.arguments);

                                System.Diagnostics.Debug.WriteLine($"执行工具: {functionName}");
                                System.Diagnostics.Debug.WriteLine($"参数: {toolCall.function.arguments}");
                                prompt_label.Text = $"正在执行工具: {functionName}...";

                                // 执行工具
                                var toolResult = ExecuteMcpTool(functionName, arguments);

                                System.Diagnostics.Debug.WriteLine($"工具执行结果: {toolResult}");

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

        // 构建消息列表（用于API请求）
        private List<object> BuildMessages(bool useMcp)
        {
            var messages = new List<object>();

            // 添加系统提示词（仅在使用MCP时）
            if (useMcp && _excelMcp != null)
            {
                var systemPrompt = @"你是一个Excel操作助手。你必须通过调用工具来操作Excel文件。

**核心原则**：
🚫 **禁止仅用文字描述操作** - 例如：""我将在A1写入数据""、""现在我把名称写入A列""
✅ **必须实际调用工具函数** - 直接使用 set_cell_value、get_worksheet_names 等工具

**重要规则**：
1. **必须直接调用工具，不要只是描述要做什么**：
   - 错误示例：""我将在A1单元格写入测试"" ❌
   - 错误示例：""现在我将这些工作表名称写入当前表的A列"" ❌
   - 正确示例：直接调用 set_cell_value 工具，参数为 row=1, column=1, value=""测试"" ✅
   - 正确示例：循环调用 set_cell_value 工具，将每个工作表名称写入 A1、A2、A3... ✅

2. **对于需要多步操作的任务，必须调用多次工具**：
   - 例如：要将5个工作表名称写入A1-A5，必须调用5次 set_cell_value 工具
   - 第一次：set_cell_value(row=1, column=1, value=""Sheet1"")
   - 第二次：set_cell_value(row=2, column=1, value=""Sheet2"")
   - ...以此类推

3. **""表""默认指工作表（worksheet）**：
   - 当用户说""新建一个表""、""创建表""时，指的是在当前工作簿中创建新的工作表（sheet），而不是创建新的工作簿
   - 当用户说""新建工作簿""、""创建Excel文件""时，才是创建工作簿
   - 例如：""新建一个销售表"" → 使用 create_worksheet，而不是 create_workbook
   - **重要**：create_worksheet 默认会在工作簿的最前面（第一个位置）创建新工作表
   - 除非用户明确说明""在某张表后面/前面新建""，否则默认就是在最前面新建

4. **创建目录表的正确方式**：
   - 当用户要求创建目录表并写入表名时，注意行号分配：
   - 如果需要添加标题（如""工作表目录""），标题应在A1，表名从A2开始
   - 例如：创建目录表 → 先在A1写入""工作表目录"" → 表名从A2、A3、A4...开始写入
   - **错误做法**：标题在A1，第一个表名也在A1 ❌
   - **正确做法**：标题在A1（row=1），第一个表名在A2（row=2），第二个在A3（row=3）✅

5. **理解""当前单元格""的含义**：
   - 当用户说""当前单元格""、""选中的单元格""、""这个单元格""时，指的是用户在Excel中当前选中的单元格或区域
   - 首先调用 get_current_selection 工具获取当前选中的单元格信息
   - 然后根据返回的行号、列号或地址来操作该单元格
   - 例如：""在当前单元格输入测试"" → 先调用 get_current_selection，获得行号和列号，再调用 set_cell_value

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
     * 用户说""在A1添加百度的链接"" → 使用 add_hyperlink
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
- 当前活跃工作簿：" + (string.IsNullOrEmpty(_activeWorkbook) ? "无" : _activeWorkbook) + @"
- 当前活跃工作表：" + (string.IsNullOrEmpty(_activeWorksheet) ? "无" : _activeWorksheet) + @"

**重要提示**：
- 如果当前活跃工作簿为""无""，请先使用 get_current_excel_info 工具获取最新的Excel环境信息
- 获取信息后，你就能知道用户当前打开的工作簿和工作表，然后可以直接对其进行操作
- 不要只是告诉用户你将要做什么，必须实际调用工具来执行操作
- 每个操作都必须对应一个工具调用，不能省略
- 当用户提到""当前单元格""时，先调用 get_current_selection 获取选中信息

**操作流程示例**：
用户：""请将当前工作簿中所有表的名称写入当前表的A列""
正确做法：
1. 调用 get_worksheet_names 获取所有工作表名称
2. 对每个工作表名称，调用 set_cell_value(row=行号, column=1, value=表名)
3. 完成后告诉用户操作完成

错误做法：
只回复""现在我将这些工作表名称写入当前表的A列""但不调用任何工具 ❌

用户：""在所有表前新建一个目录表，写入所有表名，并加上超链接""
正确做法（假设有Sheet1、Sheet2、Sheet3三个表）：
1. 调用 create_worksheet(sheetName=""目录"") → 自动在最前面创建目录表
2. 调用 get_worksheet_names() → 获取所有表名：[""目录"", ""Sheet1"", ""Sheet2"", ""Sheet3""]
3. 调用 set_cell_value(row=1, column=1, value=""工作表目录"") → 在A1写入标题
4. 调用 set_hyperlink_formula(cellAddress=""A2"", targetLocation=""Sheet1!A1"", displayText=""Sheet1"") → 第一个表名在A2
5. 调用 set_hyperlink_formula(cellAddress=""A3"", targetLocation=""Sheet2!A1"", displayText=""Sheet2"") → 第二个表名在A3
6. 调用 set_hyperlink_formula(cellAddress=""A4"", targetLocation=""Sheet3!A1"", displayText=""Sheet3"") → 第三个表名在A4
7. 告诉用户完成

**重要**：注意行号从2开始（跳过标题行A1），避免标题被覆盖

用户：""在当前单元格输入测试""
正确做法：
1. 调用 get_current_selection 获取当前选中的单元格信息
2. 从返回信息中提取行号和列号
3. 调用 set_cell_value(row=行号, column=列号, value=""测试"")
4. 告诉用户操作完成

用户：""在A1创建跳转到销售数据表的链接""
正确做法：
1. 调用 set_hyperlink_formula(cellAddress=""A1"", targetLocation=""销售数据!A1"", displayText=""查看销售数据"")
2. 告诉用户已创建工作簿内部跳转链接

错误做法：
使用 add_hyperlink 添加外部链接 ❌（这会导致无法正确跳转）

用户：""在B2添加百度搜索的链接""
正确做法：
1. 调用 add_hyperlink(cellAddress=""B2"", url=""https://www.baidu.com"", displayText=""百度搜索"")
2. 告诉用户已添加外部网址链接

错误做法：
使用 set_hyperlink_formula ❌（这只适用于工作簿内部跳转）

请根据用户的自然语言指令，**立即调用**相应的工具完成任务，而不是仅仅描述你要做什么。";

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
                    messages.Add(new
                    {
                        role = msg.Role,
                        content = msg.Content
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
            Form8 form8 = new Form8();
            form8.FormClosed += Form8_FormClosed;
            form8.ShowDialog();
        }

        private void Form8_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form7_Load(this, new EventArgs());
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
                _enterMode = decryptedContent.Split(';')[3].Split('^')[1];
                // 不在这里更新UI
            }
            catch (Exception ex)
            {
                // 记录错误日志，不更新UI
                System.Diagnostics.Debug.WriteLine($"解密配置失败：{ex.Message}");
                _apiKey = string.Empty;
                _model = string.Empty;
                _apiUrl = string.Empty;
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

