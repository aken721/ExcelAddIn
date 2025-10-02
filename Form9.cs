using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{


    public partial class Form9 : Form
    {

        //时间控制变量和任务空值变量
        private System.Windows.Forms.Timer _schedulerTimer;
        private List<ScheduledTask> _activeTasks = new List<ScheduledTask>();


        public Form9()
        {
            InitializeComponent();

            // 设置DataGridView的属性
            dataGridView1.AllowUserToAddRows = false;

            SubscribeToMouseEvents();

            // 初始化调度器定时器
            _schedulerTimer = new System.Windows.Forms.Timer();
            _schedulerTimer.Interval = 60000; // 每分钟检查一次
            _schedulerTimer.Tick += SchedulerTimer_Tick;
            _schedulerTimer.Start();

            dataGridView1.CurrentCellDirtyStateChanged += dataGridView1_CurrentCellDirtyStateChanged;
            dataGridView1.DataError += DataGridView1_DataError;
            dataGridView1.CellValueChanged += DataGridView1_CellValueChanged;
            dataGridView1.CellContentClick += DataGridView1_CellContentClick;

            LoadActiveTasks();

        }

        private void Form9_Load(object sender, EventArgs e)
        {
            panelWeek.Parent = splitContainer1.Panel2;
            panelWeek.Visible = false;
            panelMonth.Parent = splitContainer1.Panel2;
            labelMonth.Parent = panelMonth;
            labelDay.Parent = panelMonth;
            textBoxMonth.Parent = panelMonth;
            textBoxDay.Parent = panelMonth;
            buttonDropDown1.Parent = panelMonth;
            buttonDropDown2.Parent = panelMonth;
            panelMonth.Visible = false;
            btnRefresh.Visible = false;

            //触发器页面初始化
            radioButtonOnce.Checked = true;
            checkBoxStop.Visible = false;
            dateTimePicker2.Format = DateTimePickerFormat.Custom;
            dateTimePicker2.CustomFormat = "HH:mm:ss";
            dateTimePicker3.Visible = false;
            dateTimePicker4.Visible = false;
            dateTimePicker4.Format = DateTimePickerFormat.Custom;
            dateTimePicker4.CustomFormat = "HH:mm:ss";
            textBoxInterval.Text = "1";

            //操作页面初始化
            comboBoxProjectType.SelectedIndex = 0;
            textBoxCMD.Text = "请输入CMD命令";
            textBoxCMD.ForeColor = System.Drawing.Color.DarkGray; // 设置字体颜色为灰色
            textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Italic); // 设置字体为斜体                    
            textBoxCMD.ReadOnly = false;
            labelVBA.Visible = false;
            labelScript.Visible = false;
            comboBoxVBA.Visible = false;
            textBoxScript.Visible = false;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    btnRefresh.Visible = false;
                    break;
                case 1:
                    btnRefresh.Visible = false;
                    if (radioButtonOnce.Checked)
                    {
                        panelWeek.Visible = false;
                    }
                    if (radioButtonDay.Checked)
                    {
                        panelWeek.Visible = true;
                        labelFrequency.Text = "天发生一次";
                        if (textBoxInterval.Text == "" || int.Parse(textBoxInterval.Text) < 1) textBoxInterval.Text = "1";
                        flowLayoutPanelWeekDay.Visible = false;
                    }
                    if (radioButtonWeek.Checked)
                    {
                        panelWeek.Visible = true;
                        labelFrequency.Text = "周发生一次";
                        panelMonth.Visible = false;
                        if (textBoxInterval.Text == "" || int.Parse(textBoxInterval.Text) < 1) textBoxInterval.Text = "1";
                        flowLayoutPanelWeekDay.Visible = true;
                    }
                    if (radioButtonMonth.Checked)
                    {
                        panelWeek.Visible = false;
                        panelMonth.Visible = true;
                    }
                    break;
                case 2:
                    btnRefresh.Visible = false;
                    break;
                case 3:
                    if (IsSheetExist("_定时任务"))
                    {
                        System.Data.DataTable dt = ReadPlanSheet("_定时任务");
                        dataGridView1.DataSource = dt;
                        if (dataGridView1.DataSource != null)
                        {
                            ConfigureDataGridView();
                        }
                    }
                    else
                    {
                        dataGridView1.DataSource = null;
                        RemoveCustomColumns();
                    }
                    btnRefresh.Visible = true;
                    break;
            }
        }


        // 任务调度器类
        private class ScheduledTask
        {
            public string TaskName { get; set; } = string.Empty;
            public DateTime StartTime { get; set; }
            public DateTime NextRunTime { get; set; } = DateTime.MaxValue;
            public DateTime? EndTime { get; set; }
            public string Frequency { get; set; } = "一次性";
            public int Interval { get; set; } = 1;
            public string[] WeekDays { get; set; } = Array.Empty<string>();
            public string[] MonthDays { get; set; } = Array.Empty<string>();
            public string[] Months { get; set; } = Array.Empty<string>();
            public string TaskType { get; set; } = "CMD";
            public string Command { get; set; } = string.Empty;
            public DateTime? LastRunTime { get; set; }
            public string TaskStatus { get; set; } = "计划中";
            public bool IsExpired { get; set; }
            public bool IsEnabled { get; set; }
        }


        /// <summary>
        /// 1"任务名称",2"任务描述",3"开始时间",4"重复频率",5"间隔周期",
        /// 6"详细安排（月）",7"详细安排（日）",8"任务类型",9"任务内容",10"是否启用",
        /// 11"计划停止时间",12"实际执行时间",13"下次执行时间",14"上次运行结果",15"任务状态"
        /// </summary>
        /// <param name="startTime"></param>
        /// <param name="lastRunTime"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>  

        // datagridview控件中添加"是否启用"和"操作"列
        private void ConfigureDataGridView()
        {
            try
            {
                // 禁用自动生成列
                dataGridView1.AutoGenerateColumns = false;
                dataGridView1.Columns.Clear();

                // 手动添加所有需要显示的列（包括复选框列）
                // (1) 是否启用列（复选框）
                DataGridViewCheckBoxColumn enableColumn = new DataGridViewCheckBoxColumn();
                enableColumn.Name = "是否启用";
                enableColumn.DataPropertyName = "是否启用";
                enableColumn.HeaderText = "是否启用";
                enableColumn.Width = 60;
                enableColumn.TrueValue = true;
                enableColumn.FalseValue = false;
                dataGridView1.Columns.Add(enableColumn);
                // (2) 其他列
                DataGridViewTextBoxColumn taskNameColumn = new DataGridViewTextBoxColumn();
                taskNameColumn.Name = "任务名称";
                taskNameColumn.DataPropertyName = "任务名称"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(taskNameColumn);

                DataGridViewTextBoxColumn taskDescriptionColumn = new DataGridViewTextBoxColumn();
                taskDescriptionColumn.Name = "任务描述";
                taskDescriptionColumn.DataPropertyName = "任务描述"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(taskDescriptionColumn);

                DataGridViewTextBoxColumn startTimeColumn = new DataGridViewTextBoxColumn();
                startTimeColumn.Name = "开始时间";
                startTimeColumn.DataPropertyName = "开始时间"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(startTimeColumn);
                DataGridViewTextBoxColumn frequencyColumn = new DataGridViewTextBoxColumn();
                frequencyColumn.Name = "重复频率";
                frequencyColumn.DataPropertyName = "重复频率"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(frequencyColumn);

                DataGridViewTextBoxColumn intervalColumn = new DataGridViewTextBoxColumn();
                intervalColumn.Name = "间隔周期";
                intervalColumn.DataPropertyName = "间隔周期"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(intervalColumn);

                DataGridViewTextBoxColumn monthDetailColumn = new DataGridViewTextBoxColumn();
                monthDetailColumn.Name = "详细安排（月）";
                monthDetailColumn.DataPropertyName = "详细安排（月）"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(monthDetailColumn);

                DataGridViewTextBoxColumn dayDetailColumn = new DataGridViewTextBoxColumn();
                dayDetailColumn.Name = "详细安排（日）";
                dayDetailColumn.DataPropertyName = "详细安排（日）"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(dayDetailColumn);

                DataGridViewTextBoxColumn taskTypeColumn = new DataGridViewTextBoxColumn();
                taskTypeColumn.Name = "任务类型";
                taskTypeColumn.DataPropertyName = "任务类型"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(taskTypeColumn);

                DataGridViewTextBoxColumn taskContentColumn = new DataGridViewTextBoxColumn();
                taskContentColumn.Name = "任务内容";
                taskContentColumn.DataPropertyName = "任务内容"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(taskContentColumn);

                DataGridViewTextBoxColumn planEndTimeColumn = new DataGridViewTextBoxColumn();
                planEndTimeColumn.Name = "计划停止时间";
                planEndTimeColumn.DataPropertyName = "计划停止时间"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(planEndTimeColumn);

                DataGridViewTextBoxColumn actualRunTimeColumn = new DataGridViewTextBoxColumn();
                actualRunTimeColumn.Name = "实际执行时间";
                actualRunTimeColumn.DataPropertyName = "实际执行时间"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(actualRunTimeColumn);

                DataGridViewTextBoxColumn nextRunTimeColumn = new DataGridViewTextBoxColumn();
                nextRunTimeColumn.Name = "下次执行时间";
                nextRunTimeColumn.DataPropertyName = "下次执行时间"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(nextRunTimeColumn);

                DataGridViewTextBoxColumn lastRunResultColumn = new DataGridViewTextBoxColumn();
                lastRunResultColumn.Name = "上次运行结果";
                lastRunResultColumn.DataPropertyName = "上次运行结果"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(lastRunResultColumn);

                DataGridViewTextBoxColumn taskStatusColumn = new DataGridViewTextBoxColumn();
                taskStatusColumn.Name = "任务状态";
                taskStatusColumn.DataPropertyName = "任务状态"; // 绑定到 DataTable 的列名
                dataGridView1.Columns.Add(taskStatusColumn);

                // (3) 操作列（按钮）
                DataGridViewButtonColumn deleteColumn = new DataGridViewButtonColumn();
                deleteColumn.Name = "操作列";
                deleteColumn.HeaderText = "操作";
                deleteColumn.Text = "删除";
                deleteColumn.UseColumnTextForButtonValue = true;
                deleteColumn.Width = 60;
                dataGridView1.Columns.Add(deleteColumn);

                // (3) 修改列（按钮）
                DataGridViewButtonColumn editColumn = new DataGridViewButtonColumn();
                editColumn.Name = "修改列";
                editColumn.HeaderText = "修改";
                editColumn.Text = "修改";
                editColumn.UseColumnTextForButtonValue = true;
                editColumn.Width = 60;
                dataGridView1.Columns.Add(editColumn);

                // 绑定事件
                //dataGridView1.CellValueChanged -= DataGridView1_CellValueChanged;
                //dataGridView1.CellValueChanged += DataGridView1_CellValueChanged;

                //dataGridView1.CellContentClick -= DataGridView1_CellContentClick;
                //dataGridView1.CellContentClick += DataGridView1_CellContentClick;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"列初始化失败: {ex.Message}");
            }
        }


        private void DataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }



        // 更新任务的下次执行时间
        private void UpdateTaskNextRunTime(ScheduledTask task)
        {
            task.NextRunTime = CalculateNextRunTime(task);
        }


        private DateTime ParseTaskTime(string timeString, DateTime fallback)
        {
            if (string.IsNullOrWhiteSpace(timeString))
                return fallback;

            // 支持多种时间格式解析
            string[] formats = {
                                    "yyyy-MM-dd HH:mm:ss",
                                    "yyyy/MM/dd HH:mm:ss",
                                    "yyyy-MM-dd",
                                    "yyyy/MM/dd"
                                };
            if (DateTime.TryParseExact(timeString,
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out DateTime result))
            {
                return result;
            }

            return fallback;
        }

        private void DataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 处理删除按钮点击
            if (e.RowIndex >= 0 && dataGridView1.Columns[e.ColumnIndex].Name == "是否启用") DataGridView1_CellValueChanged(sender, e);

            if (e.RowIndex >= 0 && dataGridView1.Columns[e.ColumnIndex].Name == "操作列")
            {
                string taskName = dataGridView1.Rows[e.RowIndex].Cells["任务名称"].Value.ToString();

                if (System.Windows.Forms.MessageBox.Show($"确定要删除任务 '{taskName}' 吗？",
                                      "确认删除",
                                      MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    // 从Excel删除
                    DeleteTaskFromExcel(taskName);

                    // 从内存删除
                    _activeTasks.RemoveAll(t => t.TaskName == taskName);

                    // 刷新表格
                    RefreshDataGrid();
                }
            }
            else if (e.RowIndex >= 0 && dataGridView1.Columns[e.ColumnIndex].Name == "修改列")
            {
                string taskName = dataGridView1.Rows[e.RowIndex].Cells["任务名称"].Value.ToString();
                LoadTaskToForm(taskName);
                tabControl1.SelectedIndex = 0; // 跳转到第一个标签页
            }
        }

        private void LoadTaskToForm(string taskName)
        {
            try
            {
                _isEditing = true;
                _originalTaskName = taskName;

                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == taskName)
                    {
                        // 加载基本信息
                        textBoxTaskName.Text = usedRange.Cells[i, 1].Value?.ToString();
                        textBoxTaskDescription.Text = usedRange.Cells[i, 2].Value?.ToString();

                        // 解析开始时间
                        if (DateTime.TryParse(usedRange.Cells[i, 3].Value?.ToString(), out DateTime startTime))
                        {
                            dateTimePicker1.Value = startTime;
                            dateTimePicker2.Value = startTime;
                        }

                        // 设置重复频率
                        string frequency = usedRange.Cells[i, 4].Value?.ToString();
                        switch (frequency)
                        {
                            case "一次性":
                                radioButtonOnce.Checked = true;
                                break;
                            case "每天":
                                radioButtonDay.Checked = true;
                                break;
                            case "每周":
                                radioButtonWeek.Checked = true;
                                break;
                            case "每月":
                                radioButtonMonth.Checked = true;
                                break;
                        }

                        // 设置间隔周期
                        string interval = usedRange.Cells[i, 5].Value?.ToString();
                        textBoxInterval.Text = Regex.Match(interval ?? "", @"\d+").Value;

                        // 设置详细安排
                        if (frequency == "每周")
                        {
                            string[] weekDays = (usedRange.Cells[i, 6].Value?.ToString() ?? "").Split(',');
                            SetCheckedDays(flowLayoutPanelWeekDay, weekDays);
                        }
                        else if (frequency == "每月")
                        {
                            textBoxMonth.Text = usedRange.Cells[i, 6].Value?.ToString();
                            textBoxDay.Text = usedRange.Cells[i, 7].Value?.ToString();
                        }

                        // 设置任务类型
                        string taskType = usedRange.Cells[i, 8].Value?.ToString();
                        comboBoxProjectType.SelectedIndex = taskType switch
                        {
                            "运行CMD命令" => 0,
                            "运行VBA宏" => 1,
                            "运行BAT批处理" => 2,
                            "运行Python脚本" => 3,
                            "启动程序" => 4,
                            _ => 0
                        };

                        // 设置任务内容
                        string content = usedRange.Cells[i, 9].Value?.ToString();
                        switch (comboBoxProjectType.SelectedIndex)
                        {
                            case 0:
                                textBoxCMD.Text = content;
                                break;
                            case 1:
                                comboBoxVBA.SelectedItem = content;
                                break;
                            case 2:
                            case 3:
                            case 4:
                                textBoxScript.Text = content;
                                break;
                        }

                        // 设置启用状态
                        checkBoxAvailable.Checked = usedRange.Cells[i, 10].Value?.ToString() == "启用";

                        // 设置停止时间
                        if (DateTime.TryParse(usedRange.Cells[i, 11].Value?.ToString(), out DateTime endTime))
                        {
                            checkBoxStop.Checked = true;
                            dateTimePicker3.Value = endTime;
                            dateTimePicker4.Value = endTime;
                        }

                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"加载任务失败: {ex.Message}");
            }
        }

        private void SetCheckedDays(FlowLayoutPanel panel, string[] days)
        {
            foreach (Control control in panel.Controls)
            {
                if (control is System.Windows.Forms.CheckBox checkBox)
                {
                    checkBox.Checked = days.Contains(checkBox.Text);
                }
            }
        }

        private void dataGridView1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dataGridView1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void DataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex == dataGridView1.Columns["是否启用"].Index)
            {
                bool isEnabled = (bool)dataGridView1.Rows[e.RowIndex].Cells["是否启用"].Value;
                string taskName = dataGridView1.Rows[e.RowIndex].Cells["任务名称"].Value?.ToString();

                if (!string.IsNullOrEmpty(taskName))
                {
                    // 更新Excel和内存状态
                    UpdateExcelEnableStatus(taskName, isEnabled);

                    // 当启用时检查过期状态
                    if (isEnabled)
                    {
                        var task = _activeTasks.FirstOrDefault(t => t.TaskName == taskName);
                        if (task != null)
                        {
                            UpdateTaskStatus(task);
                            task.NextRunTime = CalculateNextRunTime(task);
                        }
                    }
                }
            }
        }

        // 更新Excel中的任务状态
        private void UpdateExcelEnableStatus(string taskName, bool isEnabled)
        {
            try
            {
                var sheet = (Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                var usedRange = sheet.UsedRange;
                object[,] data = (object[,])usedRange.Value2;

                for (int i = 2; i <= data.GetLength(0); i++)
                {
                    if (data[i, 1]?.ToString() == taskName)
                    {
                        sheet.Cells[i, 10] = isEnabled ? "启用" : "不启用";
                        Marshal.ReleaseComObject(usedRange);
                        Marshal.ReleaseComObject(sheet);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("更新Excel状态失败: " + ex.Message);
            }
        }

        private void DeleteTaskFromExcel(string taskName)
        {
            try
            {
                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == taskName)
                    {
                        ((Range)sheet.Rows[i]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("删除任务失败: " + ex.Message);
            }
        }




        private void DataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (e.ColumnIndex == dataGridView1.Columns["是否启用"].Index)
            {
                System.Windows.MessageBox.Show("请为'是否启用'列输入有效的布尔值(true/false)");
                e.ThrowException = false;
            }
        }

        // 移除自定义列的方法
        private void RemoveCustomColumns()
        {
            if (dataGridView1.Columns.Contains("启用列"))
            {
                dataGridView1.Columns.Remove("启用列");
            }
            if (dataGridView1.Columns.Contains("操作列"))
            {
                dataGridView1.Columns.Remove("操作列");
            }
        }


        // 加载活动任务
        private void LoadActiveTasks()
        {
            Worksheet sheet = null;
            Range usedRange = null;
            try
            {
                _activeTasks.Clear();
                if (!IsSheetExist("_定时任务")) return;

                sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                usedRange = sheet.UsedRange;
                int rowCount = usedRange.Rows.Count;

                for (int i = 2; i <= rowCount; i++)
                {
                    try
                    {
                        var statusCell = usedRange.Cells[i, 10].Value?.ToString();
                        if (statusCell != "启用") continue;

                        // 解析基础信息
                        var task = new ScheduledTask
                        {
                            TaskName = usedRange.Cells[i, 1].Value?.ToString() ?? string.Empty,
                            Frequency = usedRange.Cells[i, 4].Value?.ToString() ?? string.Empty,
                            TaskType = usedRange.Cells[i, 8].Value?.ToString() ?? string.Empty,
                            Command = usedRange.Cells[i, 9].Value?.ToString() ?? string.Empty,
                            LastRunTime = DateTime.TryParse(usedRange.Cells[i, 12].Value?.ToString(), out DateTime lrt) ? lrt : (DateTime?)null,
                            TaskStatus = usedRange.Cells[i, 15].Value?.ToString() ?? "计划中"
                        };

                        // 设置 IsEnabled 和 IsExpired
                        task.IsEnabled = usedRange.Cells[i, 10].Value?.ToString() == "启用";
                        task.IsExpired = CalculateIsExpired(task);

                        // 状态验证
                        if (!string.IsNullOrEmpty(task.TaskStatus) &&
                           (task.TaskStatus == "已失效" || task.TaskStatus == "已停用"))
                        {
                            continue;
                        }

                        // 解析开始时间
                        DateTime startTime;
                        if (DateTime.TryParse(usedRange.Cells[i, 3].Value?.ToString(), out startTime))
                        {
                            task.StartTime = startTime;
                        }
                        else
                        {
                            task.StartTime = DateTime.Now; // 解析失败时使用当前时间作为默认值
                        }

                        // 解析时间信息
                        DateTime tempDate;
                        if (DateTime.TryParse(usedRange.Cells[i, 13].Value?.ToString(), out tempDate))
                            task.NextRunTime = tempDate;
                        else
                            task.NextRunTime = DateTime.MaxValue;

                        var cellValue = usedRange.Cells[i, 5].Value?.ToString() ?? "";
                        var digits = new System.Collections.Generic.List<char>();
                        foreach (char c in cellValue)
                        {
                            if (char.IsDigit(c))
                                digits.Add(c);
                        }
                        var intervalString = new string(digits.ToArray());
                        task.Interval = string.IsNullOrEmpty(intervalString) ? 1 : int.Parse(intervalString);

                        // 根据任务类型解析不同参数
                        switch (task.Frequency)
                        {
                            case "每周":
                                task.WeekDays = (usedRange.Cells[i, 6].Value?.ToString() ?? "")
                                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                break;

                            case "每月":
                                task.Months = (usedRange.Cells[i, 6].Value?.ToString() ?? "")
                                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                                task.MonthDays = (usedRange.Cells[i, 7].Value?.ToString() ?? "")
                                    .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                                break;
                        }

                        // 自动更新状态
                        if (task.NextRunTime >= DateTime.MaxValue.AddDays(-1) ||
                           (task.EndTime.HasValue && DateTime.Now > task.EndTime))
                        {
                            UpdateTaskStatus(task, isExpired: true);
                        }

                        _activeTasks.Add(task);
                    }
                    catch (COMException ex)
                    {
                        System.Windows.MessageBox.Show($"COM 错误: {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        // 使用临时task对象调用原有LogError方法
                        var tempTask = new ScheduledTask { TaskName = $"第{i}行任务" };
                        LogError(tempTask, ex);
                    }
                }

                // 加载后刷新状态
                RefreshTaskStatuses();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("加载任务列表失败: " + ex.Message);
            }
        }

        //辅助方法：检查一次性任务是否过期
        private bool CalculateIsExpired(ScheduledTask task)
        {
            bool isOneTime = task.Frequency == "一次性";
            bool hasRun = task.LastRunTime.HasValue;
            bool endTimePassed = task.EndTime.HasValue && DateTime.Now > task.EndTime;
            return (isOneTime && hasRun) || endTimePassed;
        }

        //辅助方法：状态更新
        private void RefreshTaskStatus(ScheduledTask task)
        {
            bool isEnabled = IsTaskEnabled(task.TaskName);
            DateTime? endTime = GetTaskEndTime(task.TaskName);

            if (!isEnabled)
            {
                task.TaskStatus = "已停用";
            }
            else if (endTime.HasValue && DateTime.Now > endTime.Value)
            {
                task.TaskStatus = "已失效";
            }
            else
            {
                task.TaskStatus = "计划中";
            }

            UpdateTaskStatus(task);
        }

        private bool IsTaskEnabled(string taskName)
        {
            try
            {
                if (!IsSheetExist("_定时任务")) return false;

                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    string currentTaskName = usedRange.Cells[i, 1]?.Value?.ToString();
                    if (currentTaskName == taskName)
                    {
                        string status = usedRange.Cells[i, 10]?.Value?.ToString();
                        return status == "启用";
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取启用状态失败: {ex.Message}");
                return false;
            }
        }

        private DateTime? GetTaskEndTime(string taskName)
        {
            try
            {
                if (!IsSheetExist("_定时任务")) return null;

                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    string currentTaskName = usedRange.Cells[i, 1]?.Value?.ToString();
                    if (currentTaskName == taskName)
                    {
                        object endTimeValue = usedRange.Cells[i, 11]?.Value;
                        if (endTimeValue == null || endTimeValue.ToString() == "-") return null;

                        if (DateTime.TryParse(endTimeValue.ToString(), out DateTime endTime))
                        {
                            return endTime;
                        }
                        return null;
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取结束时间失败: {ex.Message}");
                return null;
            }
        }

        private static readonly Dictionary<DayOfWeek, string> _dayOfWeekMapping = new()
        {
            { DayOfWeek.Monday, "周一" },
            { DayOfWeek.Tuesday, "周二" },
            { DayOfWeek.Wednesday, "周三" },
            { DayOfWeek.Thursday, "周四" },
            { DayOfWeek.Friday, "周五" },
            { DayOfWeek.Saturday, "周六" },
            { DayOfWeek.Sunday, "周日" }
        };

        // 错误日志记录
        private void LogError(ScheduledTask task, Exception ex)
        {
            try
            {
                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == task.TaskName)
                    {
                        sheet.Cells[i, 12].Value = $"错误: {ex.Message}";
                        break;
                    }
                }
            }
            catch { /* 防止日志记录失败导致崩溃 */ }
        }

        // 定时器触发事件
        private void SchedulerTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                LoadActiveTasks();

                foreach (var task in _activeTasks.ToList())
                {
                    // 跳过已禁用/过期任务
                    if (!task.IsEnabled || task.IsExpired)
                        continue;

                    // 跳过非计划中状态的任务
                    if (task.TaskStatus != "计划中")
                        continue;

                    if (DateTime.Now >= task.NextRunTime.AddSeconds(-30) &&
                        DateTime.Now <= task.NextRunTime.AddSeconds(30))
                    {
                        string result = ExecuteTask(task);
                        // 计算并更新下次执行时间
                        task.NextRunTime = CalculateNextRunTime(task);
                        UpdateTaskStatus(task);
                        task.LastRunTime = DateTime.Now;
                    }
                }

                // 定时刷新状态
                RefreshTaskStatuses();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"定时器错误: {ex.Message}");
            }
        }

        // 执行任务
        private string ExecuteTask(ScheduledTask task)
        {
            string result = "成功";
            try
            {
                switch (task.TaskType)
                {
                    case "运行CMD命令":
                        ExecuteScheduledCMD(task.Command);
                        break;
                    case "运行VBA宏":
                        ExecuteScheduledMacro(task.Command);
                        break;
                    case "运行BAT批处理":
                        ExecuteScheduledBat(task.Command);
                        break;
                    case "运行Python脚本":
                        ExecuteScheduledPython(task.Command);
                        break;
                    case "启动程序":
                        ExecuteScheduledExe(task.Command);
                        break;
                }
            }
            catch (Exception ex)
            {
                result = $"失败: {ex.Message}";
                LogError(task, ex);
                System.Windows.MessageBox.Show($"执行任务 {task.TaskName} 失败: {ex.Message}");
            }
            finally
            {
                UpdateLastResult(task, result);

                // 处理一次性任务状态
                if (task.Frequency == "一次性")
                {
                    UpdateExcelEnableStatus(task.TaskName, false);
                    UpdateTaskStatus(task, isExpired: true);
                }

                // 任务执行后刷新状态
                RefreshTaskStatuses();
            }
            return result;
        }

        //更新上次运行结果
        private void UpdateLastResult(ScheduledTask task, string result)
        {
            try
            {
                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == task.TaskName)
                    {
                        sheet.Cells[i, 14] = result; // 更新第14列
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("更新上次运行结果失败：" + ex.Message);
            }
        }

        // 更新任务状态
        // 更新任务状态
        private void UpdateTaskStatus(ScheduledTask task)
        {
            try
            {
                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                Range usedRange = sheet.UsedRange;

                if (usedRange.Rows.Count >= 2)
                {
                    for (int i = 2; i <= usedRange.Rows.Count; i++)
                    {
                        if (usedRange.Cells[i, 1].Value?.ToString() == task.TaskName)
                        {
                            // 获取任务信息
                            bool isEnabled = usedRange.Cells[i, 10].Value?.ToString() == "启用";
                            DateTime? endTime = null;

                            // 解析计划停止时间
                            if (DateTime.TryParse(usedRange.Cells[i, 11].Value?.ToString(), out DateTime tempEnd))
                            {
                                endTime = tempEnd;
                            }

                            // 计算过期状态（新增逻辑）
                            bool isOneTimeTask = usedRange.Cells[i, 4].Value?.ToString() == "一次性";
                            bool hasRun = usedRange.Cells[i, 12].Value != null; // 实际执行时间存在
                            bool isExpired = (isOneTimeTask && hasRun) ||
                                            (endTime.HasValue && DateTime.Now > endTime);

                            // 更新实际执行时间
                            sheet.Cells[i, 12] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                            // 更新下次执行时间
                            sheet.Cells[i, 13] = isExpired ? "N/A" : task.NextRunTime.ToString("yyyy-MM-dd HH:mm:ss");

                            // 自动更新状态（新增核心逻辑）
                            string newStatus = "计划中";
                            if (!isEnabled)
                            {
                                newStatus = "已停用";
                            }
                            else if (isExpired)
                            {
                                newStatus = "已失效";
                                // 自动禁用过期任务
                                sheet.Cells[i, 10] = "不启用";
                            }

                            // 更新状态列（第15列）
                            sheet.Cells[i, 15] = newStatus;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("更新任务状态失败: " + ex.Message);
            }
        }

        // 计算初始下次执行时间
        private DateTime CalculateInitialNextRun(DateTime startTime, DateTime? lastRunTime)
        {
            TimeSpan originalTime = startTime.TimeOfDay;
            DateTime now = DateTime.Now;

            // 判断是否是首次运行且需要等待开始时间
            bool isFirstRunAndNotStarted = (lastRunTime == null) && (DateTime.Now < startTime);

            // 核心逻辑：若首次运行且未到开始时间，直接返回开始时间
            if (isFirstRunAndNotStarted)
            {
                return startTime;
            }


            DateTime baseTime = lastRunTime.HasValue ? lastRunTime.Value : startTime;

            if (DateTime.Now < startTime)
            {
                baseTime = startTime;
            }

            if (!int.TryParse(textBoxInterval.Text, out int interval))
            {
                interval = 1;
            }

            if (radioButtonDay.Checked)
            {
                DateTime nextDate = baseTime.AddDays(interval);
                return nextDate.Date.Add(originalTime);
            }
            else if (radioButtonWeek.Checked)
            {
                // 转换中文星期名称到DayOfWeek的映射
                var dayNameMapping = new Dictionary<string, DayOfWeek>
                {
                    { "周一", DayOfWeek.Monday }, { "周二", DayOfWeek.Tuesday },
                    { "周三", DayOfWeek.Wednesday }, { "周四", DayOfWeek.Thursday },
                    { "周五", DayOfWeek.Friday }, { "周六", DayOfWeek.Saturday },
                    { "周日", DayOfWeek.Sunday }
                };

                // 获取目标周几（用户选择或开始日期的周几）
                var selectedDays = IsCheckedBoxChecked(flowLayoutPanelWeekDay);
                var targetDays = selectedDays.Any()
                    ? selectedDays.Select(d => dayNameMapping[d]).ToList()
                    : new List<DayOfWeek> { startTime.DayOfWeek };

                // 计算基准日期（确保在开始时间之后）
                DateTime referenceDate = (now > startTime) ? now : startTime;
                referenceDate = referenceDate.Date.Add(originalTime);

                // 核心计算逻辑
                DateTime nextDate = referenceDate;
                int weeksToAdd = 0;
                bool found = false;

                // 最多检查3个周期防止无限循环
                for (int i = 0; i < 3; i++)
                {
                    // 计算周期开始日
                    DateTime periodStart = startTime.Date.AddDays(weeksToAdd * 7 * interval);

                    // 遍历周期内的每一天
                    foreach (var day in targetDays.OrderBy(d => (int)d))
                    {
                        // 计算候选日期
                        DateTime candidate = periodStart.AddDays((int)day - (int)startTime.DayOfWeek);

                        // 处理跨周情况
                        if (candidate < periodStart)
                            candidate = candidate.AddDays(7);

                        // 添加时间部分
                        candidate = candidate.Date.Add(originalTime);

                        // 有效性检查
                        if (candidate >= referenceDate)
                        {
                            nextDate = candidate;
                            found = true;
                            break;
                        }
                    }

                    if (found) break;
                    weeksToAdd++;
                }

                return found ? nextDate : DateTime.MaxValue;
            }
            else if (radioButtonMonth.Checked)
            {
                // 解析用户选择的月份
                var selectedMonths = textBoxMonth.Text.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(MonthNameToNumber)
                    .Where(m => m != -1)
                    .OrderBy(m => m)
                    .ToList();

                if (selectedMonths.Count == 0)
                    throw new InvalidOperationException("请选择有效的月份");

                // 解析日期设置（最后一天或具体日期）
                var daySpecs = new List<object>();
                var dayStrings = textBoxDay.Text.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var dayStr in dayStrings)
                {
                    string trimmed = dayStr.Trim();
                    if (trimmed == "最后一天")
                    {
                        daySpecs.Add("LastDay");
                    }
                    else if (int.TryParse(trimmed, out int day) && day >= 1 && day <= 31)
                    {
                        daySpecs.Add(day);
                    }
                }

                DateTime baseDate = lastRunTime.HasValue ? lastRunTime.Value : startTime;
                if (DateTime.Now < startTime) baseDate = startTime;

                // 检查未来10年内的所有选中月份
                for (int yearOffset = 0; yearOffset < 10; yearOffset++)
                {
                    int currentYear = baseDate.Year + yearOffset;

                    foreach (int month in selectedMonths)
                    {
                        // 跳过过去年份中已经检查过的月份
                        if (yearOffset == 0 && month < baseDate.Month && currentYear == baseDate.Year)
                            continue;

                        foreach (var daySpec in daySpecs)
                        {
                            int day = daySpec switch
                            {
                                "LastDay" => DateTime.DaysInMonth(currentYear, month),
                                int d => Math.Min(d, DateTime.DaysInMonth(currentYear, month)),
                                _ => -1
                            };

                            if (day == -1) continue;

                            DateTime candidateDate = new DateTime(currentYear, month, day)
                                .Add(baseDate.TimeOfDay);

                            if (candidateDate >= baseDate)
                            {
                                return candidateDate;
                            }
                        }
                    }
                }
                throw new InvalidOperationException("未来10年内未找到符合条件的执行时间");
            }

            // 默认返回一次性任务的时间
            return DateTime.MaxValue;
        }

        // 计算下次执行时间
        private DateTime CalculateNextRunTime(ScheduledTask task)
        {
            DateTime baseDate = task.NextRunTime;

            switch (task.Frequency)
            {
                case "一次性":
                    return DateTime.MaxValue;

                case "每天":
                    return baseDate.AddDays(task.Interval);

                case "每周":
                    // 转换中文星期名称到DayOfWeek的映射
                    var dayNameMapping = new Dictionary<string, DayOfWeek>
                    {
                        { "周一", DayOfWeek.Monday },
                        { "周二", DayOfWeek.Tuesday },
                        { "周三", DayOfWeek.Wednesday },
                        { "周四", DayOfWeek.Thursday },
                        { "周五", DayOfWeek.Friday },
                        { "周六", DayOfWeek.Saturday },
                        { "周日", DayOfWeek.Sunday }
                    };

                    // 获取目标周几集合（如果用户未选择则使用开始日期的周几）
                    var targetDays = task.WeekDays.Any()
                        ? task.WeekDays.Select(d => dayNameMapping[d]).ToList()
                        : new List<DayOfWeek> { task.StartTime.DayOfWeek };

                    // 计算基准日期（使用最后一次运行时间或开始时间）
                    DateTime baseWeekDate = task.LastRunTime?.Date ?? task.StartTime.Date;

                    // 计算间隔周数
                    int intervalWeeks = task.Interval;

                    // 计算完整周间隔天数
                    int intervalDays = intervalWeeks * 7;

                    // 查找下一个符合条件的日期
                    DateTime nextDate = baseWeekDate.AddDays(1);
                    DateTime maxCheckDate = nextDate.AddYears(1); // 最多检查1年

                    while (nextDate <= maxCheckDate)
                    {
                        // 计算与开始日期的完整周数差
                        int weeksPassed = (int)(nextDate - task.StartTime.Date).TotalDays / 7;

                        // 检查是否满足周间隔条件
                        bool isIntervalMatch = weeksPassed % intervalWeeks == 0;

                        // 检查星期几是否匹配
                        bool isDayMatch = targetDays.Contains(nextDate.DayOfWeek);

                        if (isIntervalMatch && isDayMatch)
                        {
                            // 保留原始时间部分
                            return nextDate.Date.Add(task.StartTime.TimeOfDay);
                        }

                        nextDate = nextDate.AddDays(1);
                    }

                    return DateTime.MaxValue; // 未找到符合条件的日期

                case "每月":
                    // 获取所有有效日期（升序排列）
                    var validDays = task.MonthDays
                        .Select(dayStr => int.TryParse(dayStr, out int day) ? day : -1)
                        .Where(day => day > 0 && day <= 31)
                        .OrderBy(day => day)
                        .ToList();

                    if (validDays.Count == 0)
                    {
                        return DateTime.MaxValue;
                    }

                    // 当前月份候选日期
                    var currentMonthCandidates = validDays
                        .Select(day => GetSafeDate(baseDate.Year, baseDate.Month, day))
                        .Where(d => d > baseDate)
                        .OrderBy(d => d)
                        .ToList();

                    // 优先检查当前月份剩余日期
                    if (currentMonthCandidates.Any())
                    {
                        return currentMonthCandidates.First();
                    }

                    // 当前月无有效日期时，检查后续月份
                    DateTime nextMonth = baseDate.AddMonths(task.Interval);
                    int maxCheckMonths = 12; // 最多检查12个月

                    for (int i = 0; i < maxCheckMonths; i++)
                    {
                        var monthToCheck = nextMonth.AddMonths(i);
                        var candidates = validDays
                            .Select(day => GetSafeDate(monthToCheck.Year, monthToCheck.Month, day))
                            .Where(d => d > baseDate)
                            .OrderBy(d => d)
                            .ToList();

                        if (candidates.Any())
                        {
                            return candidates.First();
                        }
                    }
                    return DateTime.MaxValue;
            }

            return DateTime.MaxValue;
        }

        // 安全日期加法（防止溢出）
        private DateTime SafeAddDays(DateTime date, int days)
        {
            try
            {
                return date.AddDays(days);
            }
            catch (ArgumentOutOfRangeException)
            {
                return DateTime.MaxValue;
            }
        }

        // 辅助方法：生成安全日期
        private DateTime GetSafeDate(int year, int month, int day)
        {
            int lastDay = DateTime.DaysInMonth(year, month);
            return new DateTime(year, month, Math.Min(day, lastDay));
        }

        //状态更新
        private void UpdateTaskStatus(ScheduledTask task, bool isExpired = false)
        {
            try
            {
                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == task.TaskName)
                    {
                        // 更新下次执行时间
                        sheet.Cells[i, 13] = task.IsExpired ? "N/A" : task.NextRunTime.ToString("yyyy-MM-dd HH:mm:ss");
                        sheet.Cells[i, 15] = isExpired ? "已失效" : "计划中";
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("更新任务状态失败: " + ex.Message);
            }
        }

        private void RefreshTaskStatuses()
        {
            Parallel.ForEach(_activeTasks, task =>
            {
                task.IsExpired = CalculateIsExpired(task);
                task.TaskStatus = task.IsExpired ? "已失效" : "计划中";
            });

            // UI线程安全调用
            if (dataGridView1.InvokeRequired)
            {
                dataGridView1.Invoke(new System.Action(() => dataGridView1.Refresh()));
            }
            else
            {
                dataGridView1.Refresh();
            }
        }

        // 订阅窗体和所有子控件的 MouseDown 事件
        private void SubscribeToMouseEvents()
        {
            // 监听窗体点击
            this.MouseDown += Form9_MouseDown;
            // 递归监听所有控件的点击（确保覆盖所有区域）
            SubscribeAllControls(this);
        }

        // 递归订阅所有控件的 MouseDown 事件
        private void SubscribeAllControls(Control parent)
        {
            foreach (Control c in parent.Controls)
            {
                c.MouseDown += Control_MouseDown;
                SubscribeAllControls(c); // 递归子控件
            }
        }

        private void Form9_MouseDown(object sender, MouseEventArgs e)
        {
            // 获取鼠标点击的屏幕坐标
            System.Drawing.Point screenClickPoint = this.PointToScreen(e.Location);
            // 获取 FlowLayoutPanel 的屏幕区域
            System.Drawing.Rectangle panelMonthScreenBounds = flowLayoutPanelMonth.RectangleToScreen(flowLayoutPanelMonth.ClientRectangle);
            System.Drawing.Rectangle panelDayScreenBounds = flowLayoutPanelDay.RectangleToScreen(flowLayoutPanelDay.ClientRectangle);
            // 如果点击位置不在面板内，且面板可见，则隐藏
            if (!panelMonthScreenBounds.Contains(screenClickPoint) && flowLayoutPanelMonth.Visible)
            {
                flowLayoutPanelMonth.Visible = false;
            }

            if (!panelDayScreenBounds.Contains(screenClickPoint) && flowLayoutPanelDay.Visible)
            {
                flowLayoutPanelDay.Visible = false;
            }
        }

        // 处理控件点击事件
        private void Control_MouseDown(object sender, MouseEventArgs e)
        {
            // 将控件坐标转换为窗体坐标
            Control control = (Control)sender;
            System.Drawing.Point screenPoint = control.PointToScreen(e.Location);
            System.Drawing.Point formPoint = this.PointToClient(screenPoint);
            Form9_MouseDown(this, new MouseEventArgs(e.Button, e.Clicks, formPoint.X, formPoint.Y, e.Delta));
        }

        private void radioButtonOnce_Click(object sender, EventArgs e)
        {
            panelWeek.Visible = false;
            panelMonth.Visible = false;
        }

        private void radioButtonDay_Click(object sender, EventArgs e)
        {
            panelMonth.Visible = false;
            panelWeek.Visible = true;
            labelFrequency.Text = "天发生一次";
            if (textBoxInterval.Text == "" || int.Parse(textBoxInterval.Text) < 1) textBoxInterval.Text = "1";
            flowLayoutPanelWeekDay.Visible = false;
        }

        private void radioButtonWeek_Click(object sender, EventArgs e)
        {
            panelMonth.Visible = false;
            panelWeek.Visible = true;
            labelFrequency.Text = "周发生一次";
            if (int.Parse(textBoxInterval.Text) < 1) textBoxInterval.Text = "1";
            flowLayoutPanelWeekDay.Visible = true;
        }

        private void radioButtonMonth_Click(object sender, EventArgs e)
        {
            panelWeek.Visible = false;
            panelMonth.Visible = true;
            labelMonth.Visible = true;
            labelDay.Visible = true;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }



        private string[] IsCheckedBoxChecked(FlowLayoutPanel flowLayoutPanel)
        {

            return flowLayoutPanel.Controls.OfType<System.Windows.Forms.CheckBox>()
                .Where(c => c.Checked && c.Text != "选择全部")
                .Select(c => c.Text)
                .ToArray();
        }

        private void flowLayoutPanelDay_VisibleChanged(object sender, EventArgs e)
        {
            if (flowLayoutPanelDay.Visible)
            {
                if (textBoxDay.Text != "")
                {
                    string[] dayStringSelected = textBoxDay.Text.Split(',').ToArray();

                    foreach (var control in flowLayoutPanelDay.Controls)
                    {
                        if (control is System.Windows.Forms.CheckBox checkBox && dayStringSelected.Contains(checkBox.Text))
                        {
                            checkBox.Checked = true;
                        }
                    }
                }
            }
            else
            {
                textBoxDay.Text = String.Join(",", IsCheckedBoxChecked(flowLayoutPanelDay));
            }
        }

        private void flowLayoutPanelMonth_VisibleChanged(object sender, EventArgs e)
        {
            if (flowLayoutPanelMonth.Visible)
            {
                if (textBoxMonth.Text != "")
                {
                    string[] monthStringSelected = textBoxMonth.Text.Split(',').ToArray();

                    foreach (var control in flowLayoutPanelMonth.Controls)
                    {
                        if (control is System.Windows.Forms.CheckBox checkBox && monthStringSelected.Contains(checkBox.Text))
                        {
                            checkBox.Checked = true;
                        }
                    }
                }
            }
            else
            {
                textBoxMonth.Text = String.Join(", ", IsCheckedBoxChecked(flowLayoutPanelMonth));
            }
        }

        private void checkBoxMonth13_Click(object sender, EventArgs e)
        {
            if (checkBoxMonth13.Checked)
            {
                checkBoxMonth13.Checked = true;
                checkBoxMonth1.Checked = true;
                checkBoxMonth2.Checked = true;
                checkBoxMonth3.Checked = true;
                checkBoxMonth4.Checked = true;
                checkBoxMonth5.Checked = true;
                checkBoxMonth6.Checked = true;
                checkBoxMonth7.Checked = true;
                checkBoxMonth8.Checked = true;
                checkBoxMonth9.Checked = true;
                checkBoxMonth10.Checked = true;
                checkBoxMonth11.Checked = true;
                checkBoxMonth12.Checked = true;
            }
            else
            {
                checkBoxMonth13.Checked = false;
                checkBoxMonth1.Checked = false;
                checkBoxMonth2.Checked = false;
                checkBoxMonth3.Checked = false;
                checkBoxMonth4.Checked = false;
                checkBoxMonth5.Checked = false;
                checkBoxMonth6.Checked = false;
                checkBoxMonth7.Checked = false;
                checkBoxMonth8.Checked = false;
                checkBoxMonth9.Checked = false;
                checkBoxMonth10.Checked = false;
                checkBoxMonth11.Checked = false;
                checkBoxMonth12.Checked = false;

            }
        }

        private void comboBoxProjectType_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBoxProjectType.SelectedIndex)
            {
                //cmd
                case 0:
                    textBoxCMD.Text = "请输入CMD命令";
                    textBoxCMD.ReadOnly = false;
                    labelVBA.Visible = false;
                    comboBoxVBA.Visible = false;
                    labelScript.Visible = false;
                    textBoxScript.Visible = false;
                    break;

                //vba
                case 1:
                    textBoxCMD.Text = "请输入CMD命令";
                    textBoxCMD.ForeColor = System.Drawing.Color.DarkGray; // 设置字体颜色为灰色
                    textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Italic); // 设置字体为斜体                    
                    textBoxCMD.ReadOnly = true;
                    labelScript.Visible = false;
                    textBoxScript.Visible = false;
                    PopulateVBAProceduresOptimized();
                    if (comboBoxVBA.Items.Count > 0)
                    {
                        comboBoxVBA.SelectedIndex = 0;

                        labelVBA.Visible = true;
                        comboBoxVBA.Visible = true;
                    }
                    else
                    {
                        System.Windows.MessageBox.Show("Excel中没有可使用的VBA模块");
                    }
                    break;

                //bat
                case 2:
                    textBoxCMD.Text = "请输入CMD命令";
                    textBoxCMD.ForeColor = System.Drawing.Color.DarkGray; // 设置字体颜色为灰色
                    textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Italic); // 设置字体为斜体                    
                    textBoxCMD.ReadOnly = true;
                    labelVBA.Visible = false;
                    comboBoxVBA.Visible = false;
                    labelScript.Text = "请双击文本框选择脚本文件";
                    labelScript.Visible = true;
                    textBoxScript.Visible = true;
                    break;

                //python
                case 3:
                    textBoxCMD.Text = "请输入CMD命令";
                    textBoxCMD.ForeColor = System.Drawing.Color.DarkGray; // 设置字体颜色为灰色
                    textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Italic); // 设置字体为斜体                    
                    textBoxCMD.ReadOnly = true;
                    labelVBA.Visible = false;
                    comboBoxVBA.Visible = false;
                    labelScript.Text = "请双击文本框选择脚本文件";
                    labelScript.Visible = true;
                    textBoxScript.Visible = true;
                    break;

                //program
                case 4:
                    textBoxCMD.Text = "请输入CMD命令";
                    textBoxCMD.ForeColor = System.Drawing.Color.DarkGray; // 设置字体颜色为灰色
                    textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Italic); // 设置字体为斜体                    
                    textBoxCMD.ReadOnly = true;
                    labelVBA.Visible = false;
                    comboBoxVBA.Visible = false;
                    labelScript.Text = "请双击文本框选择可执行程序";
                    labelScript.Visible = true;
                    textBoxScript.Visible = true;
                    break;
            }
        }


        //VBA宏列表加载
        private void PopulateVBAProceduresOptimized()
        {
            try
            {
                comboBoxVBA.BeginUpdate();
                comboBoxVBA.Items.Clear();

                Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                if (workbook == null) return;

                // 检查是否支持访问 VBA 工程
                if (!IsVBAProjectAccessEnabled())
                {
                    System.Windows.MessageBox.Show("请启用 Excel 的 VBA 工程访问权限！");
                    return;
                }

                var project = workbook.VBProject;
                foreach (VBComponent component in project.VBComponents)
                {
                    var module = component.CodeModule;
                    int count = module.CountOfLines;
                    if (count == 0) continue;

                    for (int i = 1; i <= count;)
                    {
                        string procName = module.get_ProcOfLine(i, out vbext_ProcKind kind);
                        if (kind == vbext_ProcKind.vbext_pk_Proc)
                        {
                            comboBoxVBA.Items.Add($"{component.Name}.{procName}");
                            i += module.get_ProcCountLines(procName, kind);
                        }
                        else
                        {
                            i++;
                        }
                    }
                }
            }
            catch (COMException ex) when (ex.ErrorCode == unchecked((int)0x800AC472))
            {
                System.Windows.MessageBox.Show("无法访问 VBA 工程，请检查 Excel 权限设置！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载 VBA 宏失败: {ex.Message}");
            }
            finally
            {
                comboBoxVBA.EndUpdate();
            }
        }

        private bool IsVBAProjectAccessEnabled()
        {
            try
            {
                // 尝试访问 VBA 工程属性，检测权限
                var project = Globals.ThisAddIn.Application.ActiveWorkbook.VBProject;
                return project != null;
            }
            catch (COMException)
            {
                return false;
            }
        }


        private void textBoxCMD_TextChanged(object sender, EventArgs e)
        {
            if (textBoxCMD.Text == "请输入CMD命令")
            {
                textBoxCMD.ForeColor = System.Drawing.Color.DarkGray; // 设置字体颜色为灰色
                textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Italic);// 设置字体为斜体
            }
            else
            {
                textBoxCMD.ForeColor = System.Drawing.Color.Black; // 设置字体颜色为黑色
                textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Regular); // 设置字体为常规

            }
        }

        private void textBoxScript_DoubleClick(object sender, EventArgs e)
        {
            switch (comboBoxProjectType.SelectedIndex)
            {
                case 0:
                    break;
                case 1:
                    break;
                case 2:
                    // 选择脚本文件
                    using (OpenFileDialog openFileDialog = new OpenFileDialog())
                    {
                        openFileDialog.Title = "请选择bat批处理文件";
                        openFileDialog.Filter = "批处理文件 (*.bat)|*.bat";
                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            textBoxScript.Text = openFileDialog.FileName;
                        }
                    }
                    break;
                case 3:
                    // 选择Python脚本文件
                    using (OpenFileDialog openFileDialog = new OpenFileDialog())
                    {
                        openFileDialog.Title = "请选择Python脚本文件";
                        openFileDialog.Filter = "Python文件 (*.py)|*.py";
                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            textBoxScript.Text = openFileDialog.FileName;
                        }
                    }
                    break;
                case 4:
                    // 选择可执行程序文件
                    using (OpenFileDialog openFileDialog = new OpenFileDialog())
                    {
                        openFileDialog.Title = "请选择可执行程序文件";
                        openFileDialog.Filter = "可执行文件 (*.exe)|*.exe";
                        if (openFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            textBoxScript.Text = openFileDialog.FileName;
                        }
                    }
                    break;
                default:
                    break;
            }
        }

        private void textBoxCMD_Enter(object sender, EventArgs e)
        {
            if (!textBoxCMD.ReadOnly)
            {
                if (textBoxCMD.Text == "请输入CMD命令")
                {
                    textBoxCMD.Text = ""; // 清空文本框
                    textBoxCMD.ForeColor = System.Drawing.Color.Black; // 设置字体颜色为黑色
                    textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Regular); // 设置字体为常规
                }
            }

        }

        private void textBoxCMD_Leave(object sender, EventArgs e)
        {
            if (!textBoxCMD.ReadOnly)
            {
                if (textBoxCMD.Text == "")
                {
                    textBoxCMD.Text = "请输入CMD命令";
                    textBoxCMD.ForeColor = System.Drawing.Color.DarkGray; // 设置字体颜色为灰色
                    textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Italic); // 设置字体为斜体
                }
            }
        }


        private bool _isEditing = false;                  //判断是修改编辑还是新增任务
        private string _originalTaskName = string.Empty;  //原始任务名称


        private void btnConfirm_Click(object sender, EventArgs e)
        {
            if (!IsCorrectDetail())
            {
                return;
            }

            // 修改后的名称检查逻辑
            if (IsTaskNameExists(textBoxTaskName.Text) &&
               !(_isEditing && textBoxTaskName.Text == _originalTaskName))
            {
                System.Windows.MessageBox.Show("任务名称已存在，请重新设置");
                tabControl1.SelectedIndex = 0;
                return;
            }

            SetPlanSheet();

            System.Data.DataTable dt = ReadPlanSheet("_定时任务");
            dataGridView1.DataSource = dt;
            tabControl1.SelectedIndex = 3;

            // 重置编辑状态
            _isEditing = false;
            _originalTaskName = string.Empty;

            detailClear();
            Thread.Sleep(1000);
            this.Dispose();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            if (!IsCorrectDetail())
            {
                return;
            }

            // 修改后的名称检查逻辑
            if (IsTaskNameExists(textBoxTaskName.Text) &&
               !(_isEditing && textBoxTaskName.Text == _originalTaskName))
            {
                System.Windows.MessageBox.Show("任务名称已存在，请重新设置");
                return;
            }

            SetPlanSheet();
            RefreshDataGrid();
            tabControl1.SelectedIndex = 3;

            // 重置编辑状态
            _isEditing = false;
            _originalTaskName = string.Empty;

            detailClear();
        }

        // 检查任务名称是否存在
        private bool IsTaskNameExists(string taskName)
        {
            if (!IsSheetExist("_定时任务")) return false;

            Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
            Range usedRange = sheet.UsedRange;

            for (int i = 2; i <= usedRange.Rows.Count; i++)
            {
                string existingName = usedRange.Cells[i, 1].Value?.ToString();
                if (existingName == taskName)
                {
                    return true;
                }
            }
            return false;
        }

        //excel中添加定时任务表
        private void SetPlanSheet()
        {
            // 检查是否是编辑模式
            if (_isEditing)
            {
                UpdateExistingTask();
            }
            else
            {
                AddNewTask();
            }
        }

        private void AddNewTask()
        {
            string[] planSheetTitle =
                [
                    "任务名称",
                "任务描述",
                "开始时间",
                "重复频率",
                "间隔周期",
                "详细安排（月）",
                "详细安排（日）",
                "任务类型",
                "任务内容",
                "是否启用",
                "计划停止时间",
                "实际执行时间",
                "下次执行时间",
                "上次运行结果",
                "任务状态"
                ];
            string _sheetName = "_定时任务";
            if (!IsSheetExist(_sheetName))
            {
                ThisAddIn.app.Application.ScreenUpdating = false;
                Excel.Worksheet newSheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Add();
                newSheet.Name = _sheetName;
                for (int i = 0; i < planSheetTitle.Length; i++)
                {
                    newSheet.Cells[1, i + 1].Value = planSheetTitle[i];
                }
                newSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                ThisAddIn.app.Application.ScreenUpdating = true;
            }
            int rowCount = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[_sheetName].UsedRange.Rows.Count;
            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[_sheetName];
            sheet.Cells[rowCount + 1, 1].Value = textBoxTaskName.Text;          //任务名称
            sheet.Cells[rowCount + 1, 2].Value = textBoxTaskDescription.Text;  //任务描述
            string date = dateTimePicker1.Value.ToString("yyyy-MM-dd");       //开始日期
            string time = dateTimePicker2.Value.ToString("HH:mm:ss");        //开始时间
            DateTime dateTime = DateTime.Parse(date + " " + time);
            sheet.Cells[rowCount + 1, 3].Value = dateTime.ToString("yyyy-MM-dd HH:mm:ss");     //开始日期+时间
            sheet.Cells[rowCount + 1, 4].Value = radioButtonOnce.Checked ? "一次性" : radioButtonDay.Checked ? "每天" : radioButtonWeek.Checked ? "每周" : "每月";     //重复频率

            //间隔周期
            if (radioButtonDay.Checked)
            {
                sheet.Cells[rowCount + 1, 5].Value = $"每{textBoxInterval.Text}天";
            }
            else if (radioButtonWeek.Checked)
            {
                sheet.Cells[rowCount + 1, 5].Value = $"每{textBoxInterval.Text}周";
            }
            else
            {
                sheet.Cells[rowCount + 1, 5].Value = "-";
            }
            //详细安排
            if (radioButtonWeek.Checked)
            {
                sheet.Cells[rowCount + 1, 6].Value = String.Join(",", IsCheckedBoxChecked(flowLayoutPanelWeekDay));
                sheet.Cells[rowCount + 1, 7].Value = "-";
            }
            else if (radioButtonMonth.Checked)
            {
                sheet.Cells[rowCount + 1, 6].Value = textBoxMonth.Text;
                sheet.Cells[rowCount + 1, 7].Value = textBoxDay.Text;
            }
            else
            {
                sheet.Cells[rowCount + 1, 6].Value = "-";
                sheet.Cells[rowCount + 1, 7].Value = "-";
            }
            sheet.Cells[rowCount + 1, 8].Value = comboBoxProjectType.SelectedItem.ToString();  //任务类型

            //任务内容
            switch (comboBoxProjectType.SelectedIndex)
            {
                case 0:
                    sheet.Cells[rowCount + 1, 9].Value = textBoxCMD.Text;
                    break;
                case 1:
                    sheet.Cells[rowCount + 1, 9].Value = comboBoxVBA.SelectedItem.ToString();
                    break;
                case 2:
                    sheet.Cells[rowCount + 1, 9].Value = textBoxScript.Text;
                    break;
                case 3:
                    sheet.Cells[rowCount + 1, 9].Value = textBoxScript.Text;
                    break;
                case 4:
                    sheet.Cells[rowCount + 1, 9].Value = textBoxScript.Text;
                    break;
            }
            //是否启用
            sheet.Cells[rowCount + 1, 10].Value = checkBoxAvailable.Checked ? "启用" : "不启用";

            //计划停止时间
            if (checkBoxStop.Checked)
            {
                string stopDate = dateTimePicker3.Value.ToString("yyyy-MM-dd");
                string stopTime = dateTimePicker4.Value.ToString("HH:mm:ss");
                DateTime stopDateTime = DateTime.Parse(stopDate + " " + stopTime);
                sheet.Cells[rowCount + 1, 11].Value = stopDateTime.ToString("yyyy-MM-dd HH:mm:ss");
            }
            else
            {
                sheet.Cells[rowCount + 1, 11].Value = "-";
            }

            // 获取实际执行时间
            DateTime? lastRunTime = null;
            var lastRunCellValue = sheet.Cells[rowCount + 1, 12].Value?.ToString();

            // 修正解析逻辑
            if (!string.IsNullOrEmpty(lastRunCellValue) && lastRunCellValue != "-")
            {
                if (DateTime.TryParse(lastRunCellValue, out DateTime temp))
                {
                    lastRunTime = temp;
                }
                else
                {
                    lastRunTime = null;
                }
            }
            else
            {
                lastRunTime = null;
            }

            // 下次执行时间
            DateTime nextRunTime = CalculateInitialNextRun(dateTime, lastRunTime);
            sheet.Cells[rowCount + 1, 13].Value = nextRunTime.ToString("yyyy-MM-dd HH:mm:ss");
        }

        private void UpdateExistingTask()
        {
            try
            {
                Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == _originalTaskName)
                    {
                        // 先删除旧记录
                        ((Range)sheet.Rows[i]).Delete(XlDeleteShiftDirection.xlShiftUp);
                        // 添加新记录（使用AddNewTask方法）
                        AddNewTask();
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"更新任务失败: {ex.Message}");
            }
        }


        // 将中文月份转换为数字
        private int MonthNameToNumber(string chineseMonth)
        {
            var monthMap = new Dictionary<string, int>
            {
                {"一月", 1}, {"二月", 2}, {"三月", 3}, {"四月", 4},
                {"五月", 5}, {"六月", 6}, {"七月", 7}, {"八月", 8},
                {"九月", 9}, {"十月", 10}, {"十一月", 11}, {"十二月", 12}
            };
            return monthMap.TryGetValue(chineseMonth, out int month) ? month : -1;
        }



        //读取excel中定时任务表
        private System.Data.DataTable ReadPlanSheet(string sheetName)
        {
            var dataTable = new System.Data.DataTable();
            try
            {
                // 清空旧数据
                dataTable.Clear();
                dataTable.Columns.Clear();

                Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[sheetName];
                Excel.Range usedRange = sheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int columnCount = usedRange.Columns.Count;

                // 添加所有列（保持与Excel相同的顺序）
                for (int i = 1; i <= columnCount; i++)
                {
                    string columnName = usedRange.Cells[1, i].Value?.ToString() ?? $"Column{i}";
                    dataTable.Columns.Add(columnName);
                }

                // 添加数据行
                for (int i = 2; i <= rowCount; i++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int j = 1; j <= columnCount; j++)
                    {
                        dataRow[j - 1] = usedRange.Cells[i, j].Value?.ToString() ?? "";
                    }
                    dataTable.Rows.Add(dataRow);
                }

                // 添加布尔类型的"是否启用"列（替换原有的字符串列）
                if (dataTable.Columns.Contains("是否启用"))
                {
                    // 临时保存原始列数据
                    var enableStatusData = dataTable.AsEnumerable()
                        .Select(r => r["是否启用"].ToString() == "启用")
                        .ToList();

                    // 移除原有列
                    dataTable.Columns.Remove("是否启用");

                    // 添加新的布尔列
                    DataColumn boolColumn = new DataColumn("是否启用", typeof(bool));
                    dataTable.Columns.Add(boolColumn);

                    // 填充数据
                    for (int i = 0; i < enableStatusData.Count; i++)
                    {
                        dataTable.Rows[i]["是否启用"] = enableStatusData[i];
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("读取定时任务表失败: " + ex.Message);
            }
            return dataTable;
        }

        private bool IsSheetExist(string sheetName)
        {
            Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    return true;
                }
            }
            return false;
        }

        private void ResetAllControls()
        {
            textBoxTaskName.Text = "";
            textBoxTaskDescription.Text = "";
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;
            radioButtonOnce.Checked = true;
            radioButtonDay.Checked = false;
            radioButtonWeek.Checked = false;
            radioButtonMonth.Checked = false;
            textBoxInterval.Text = "1";
            textBoxMonth.Text = "";
            textBoxDay.Text = "";
            comboBoxProjectType.SelectedIndex = 0;
            comboBoxVBA.Items.Clear();
            textBoxScript.Text = "";
            textBoxCMD.Text = "请输入CMD命令";
            textBoxCMD.ForeColor = System.Drawing.Color.DarkGray; // 设置字体颜色为灰色
            textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Italic); // 设置字体为斜体                    
            textBoxCMD.ReadOnly = false;
            checkBoxAvailable.Checked = true;
            checkBoxStop.Checked = false;
            dateTimePicker3.Value = DateTime.Now;
            dateTimePicker4.Value = DateTime.Now;
            foreach (Control control in flowLayoutPanelDay.Controls)
            {
                if (control is System.Windows.Forms.CheckBox checkBox)
                {
                    checkBox.Checked = false;
                }
            }
            foreach (Control control in flowLayoutPanelMonth.Controls)
            {
                if (control is System.Windows.Forms.CheckBox checkBox)
                {
                    checkBox.Checked = false;
                }
            }
            foreach (Control control in flowLayoutPanelWeekDay.Controls)
            {
                if (control is System.Windows.Forms.CheckBox checkBox)
                {
                    checkBox.Checked = false;
                }
            }
        }

        private void radioButtonOnce_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButtonOnce.Checked)
            {
                checkBoxStop.Visible = false;
            }
            else
            {
                checkBoxStop.Visible = true;
            }
        }

        private void checkBoxStop_VisibleChanged(object sender, EventArgs e)
        {
            if (!checkBoxStop.Visible)
            {
                checkBoxStop.Checked = false;
                dateTimePicker3.Visible = false;
                dateTimePicker4.Visible = false;
            }
            else
            {
                dateTimePicker3.Visible = true;
                dateTimePicker4.Visible = true;
            }
        }



        //运行VBA
        private void ExecuteScheduledMacro(string macroName)
        {
            try
            {
                ThisAddIn.app.Application.Run(macroName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("执行宏时出错: " + ex.Message);
            }
        }

        //运行CMD
        private void ExecuteScheduledCMD(string cmd)
        {
            try
            {
                System.Diagnostics.Process.Start("cmd.exe", "/C " + cmd);
            }
            catch (Exception ex)
            {
                Console.WriteLine("执行CMD命令时出错: " + ex.Message);
            }
        }

        //运行bat
        private async void ExecuteScheduledBat(string batFilePath)
        {
            try
            {
                await Task.Run(() => System.Diagnostics.Process.Start(batFilePath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("执行bat文件时出错: " + ex.Message);
            }
        }

        //运行python
        private async void ExecuteScheduledPython(string pythonFilePath)
        {
            try
            {
                await Task.Run(() => System.Diagnostics.Process.Start("python", pythonFilePath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("执行Python脚本时出错: " + ex.Message);
            }
        }

        //运行exe
        private async void ExecuteScheduledExe(string exeFilePath)
        {
            try
            {
                await Task.Run(() => System.Diagnostics.Process.Start(exeFilePath));
            }
            catch (Exception ex)
            {
                Console.WriteLine("执行exe文件时出错: " + ex.Message);
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            LoadActiveTasks();
            RefreshDataGrid();
            RefreshTaskStatuses();
        }

        private void RefreshDataGrid()
        {
            RemoveCustomColumns(); // 先清除旧列

            if (dataGridView1.InvokeRequired)
            {
                dataGridView1.Invoke(new System.Action(() => RefreshDataGrid()));
                return;
            }

            if (IsSheetExist("_定时任务"))
            {
                System.Data.DataTable dt = ReadPlanSheet("_定时任务");
                dataGridView1.DataSource = dt;
                ConfigureDataGridView(); // 每次刷新都重新配置列
            }
            else
            {
                dataGridView1.DataSource = null;
            }
        }

        private void textBoxMonth_Click(object sender, EventArgs e)
        {
            if (!flowLayoutPanelMonth.Visible)
            {
                flowLayoutPanelDay.Visible = false;
                flowLayoutPanelMonth.Visible = true;
                flowLayoutPanelMonth.BringToFront();
            }
            else
            {
                flowLayoutPanelDay.Visible = false;
                flowLayoutPanelMonth.Visible = false;
                flowLayoutPanelMonth.SendToBack();

            }
        }

        private void textBoxDay_Click(object sender, EventArgs e)
        {
            if (!flowLayoutPanelDay.Visible)
            {
                flowLayoutPanelMonth.Visible = false;
                flowLayoutPanelDay.Visible = true;
                flowLayoutPanelDay.BringToFront();
            }
            else
            {
                flowLayoutPanelMonth.Visible = false;
                flowLayoutPanelDay.Visible = false;
                flowLayoutPanelDay.SendToBack();
            }
        }

        private void buttonDropDown1_Click(object sender, EventArgs e)
        {
            if (!flowLayoutPanelMonth.Visible)
            {
                flowLayoutPanelDay.Visible = false;
                flowLayoutPanelMonth.Visible = true;
                flowLayoutPanelMonth.BringToFront();
            }
            else
            {
                flowLayoutPanelDay.Visible = false;
                flowLayoutPanelMonth.Visible = false;
                flowLayoutPanelMonth.SendToBack();

            }
        }

        private void buttonDropDown2_Click(object sender, EventArgs e)
        {
            if (!flowLayoutPanelDay.Visible)
            {
                flowLayoutPanelMonth.Visible = false;
                flowLayoutPanelDay.Visible = true;
                flowLayoutPanelDay.BringToFront();
            }
            else
            {
                flowLayoutPanelMonth.Visible = false;
                flowLayoutPanelDay.Visible = false;
                flowLayoutPanelDay.SendToBack();
            }
        }


        private bool IsCorrectDetail()
        {
            //检查任务名称
            if (textBoxTaskName.Text == "")
            {
                System.Windows.MessageBox.Show("任务名称不能为空");
                tabControl1.SelectedIndex = 0;
                return false;
            }

            //检查触发器
            DateTime selectedStartDate = dateTimePicker1.Value.Date;
            DateTime selectedStartTime = dateTimePicker2.Value;
            DateTime selectedStopDate = dateTimePicker3.Value.Date;
            DateTime selectedStopTime = dateTimePicker4.Value;
            DateTime startDateTime = new DateTime(selectedStartDate.Year, selectedStartDate.Month, selectedStartDate.Day, selectedStartTime.Hour, selectedStartTime.Minute, selectedStartTime.Second);
            DateTime stopDateTime = new DateTime(selectedStopDate.Year, selectedStopDate.Month, selectedStopDate.Day, selectedStopTime.Hour, selectedStopTime.Minute, selectedStopTime.Second);
            DateTime latestValidTime = startDateTime > DateTime.Now ? startDateTime : DateTime.Now;
            if (radioButtonOnce.Checked && startDateTime.AddMinutes(10) < DateTime.Now)
            {
                System.Windows.MessageBox.Show("开始时间不能早于当前时间");
                tabControl1.SelectedIndex = 1;
                return false;
            }
            else if (radioButtonDay.Checked && int.Parse(textBoxInterval.Text) < 1)
            {
                System.Windows.MessageBox.Show("间隔日应大于等于1");
                tabControl1.SelectedIndex = 1;
                return false;
            }
            else if (radioButtonWeek.Checked && int.Parse(textBoxInterval.Text) < 1)
            {
                System.Windows.MessageBox.Show("间隔周应大于等于1");
                tabControl1.SelectedIndex = 1;
                return false;
            }
            else if (radioButtonMonth.Checked && int.Parse(textBoxInterval.Text) <= 0)
            {
                System.Windows.MessageBox.Show("间隔时间应大于等于1");
                tabControl1.SelectedIndex = 1;
                return false;
            }
            else if (radioButtonMonth.Checked && textBoxDay.Text == "" && textBoxMonth.Text == "")
            {
                System.Windows.MessageBox.Show("请设置详细安排");
                tabControl1.SelectedIndex = 1;
                return false;
            }
            else if (checkBoxStop.Checked && stopDateTime < latestValidTime.AddMinutes(10))
            {
                System.Windows.MessageBox.Show("停止时间不能早于开始时间或当前时间");
                tabControl1.SelectedIndex = 1;
                return false;
            }

            //检查任务设置
            if (comboBoxProjectType.SelectedIndex == 0 && textBoxCMD.Text == "请输入CMD命令")
            {
                System.Windows.MessageBox.Show("cmd命令不能为空");
                tabControl1.SelectedIndex = 2;
                return false;
            }
            else if (comboBoxProjectType.SelectedIndex == 0 && textBoxCMD.Text == "")
            {
                System.Windows.MessageBox.Show("cmd命令不能为空");
                tabControl1.SelectedIndex = 2;
                return false;
            }
            else if (comboBoxProjectType.SelectedIndex == 1 && comboBoxVBA.SelectedItem == null)
            {
                System.Windows.MessageBox.Show("请选择要执行的VBA宏命令");
                tabControl1.SelectedIndex = 2;
                return false;
            }
            else if (comboBoxProjectType.SelectedIndex == 2 && !File.Exists(textBoxScript.Text))
            {
                System.Windows.MessageBox.Show("请选择要执行的bat批处理文件");
                tabControl1.SelectedIndex = 2;
                return false;
            }
            else if (comboBoxProjectType.SelectedIndex == 3 && !File.Exists(textBoxScript.Text))
            {
                System.Windows.MessageBox.Show("请选择要执行的python脚本文件");
                tabControl1.SelectedIndex = 2;
                return false;
            }
            else if (comboBoxProjectType.SelectedIndex == 4 && !File.Exists(textBoxScript.Text))
            {
                System.Windows.MessageBox.Show("请选择要打开的可执行程序");
                tabControl1.SelectedIndex = 2;
                return false;
            }
            return true;
        }

        //恢复初始化
        private void detailClear()
        {
            //任务页恢复初始化
            if (textBoxTaskName.Text != "")
            {
                textBoxTaskName.Text = "";
            }
            if (textBoxTaskDescription.Text != "")
            {
                textBoxTaskDescription.Text = "";
            }

            //触发器页恢复初始化
            //开始时间
            dateTimePicker1.Value = DateTime.Now;
            dateTimePicker2.Value = DateTime.Now;

            //停止时间
            dateTimePicker3.Value = DateTime.Now;
            dateTimePicker4.Value = DateTime.Now.AddMinutes(10);
            //是否启用
            checkBoxAvailable.Checked = true;
            //停止计划
            checkBoxStop.Checked = false;
            dateTimePicker3.Visible = false;
            dateTimePicker4.Visible = false;
            //一次性
            radioButtonOnce.Checked = true;
            //每天
            radioButtonDay.Checked = false;
            //每周
            radioButtonWeek.Checked = false;
            //每月
            radioButtonMonth.Checked = false;
            textBoxInterval.Text = "1";
            //周详细安排
            foreach (Control control in flowLayoutPanelWeekDay.Controls)
            {
                if (control is System.Windows.Forms.CheckBox checkBox)
                {
                    checkBox.Checked = false;
                }
            }
            //月详细安排（月份）
            textBoxMonth.Text = "";
            //月详细安排（日期）
            textBoxDay.Text = "";


            //操作页恢复初始化
            comboBoxProjectType.SelectedIndex = 0;
            textBoxCMD.Text = "请输入CMD命令";
            textBoxCMD.ForeColor = System.Drawing.Color.DarkGray; // 设置字体颜色为灰色
            textBoxCMD.Font = new System.Drawing.Font(textBoxCMD.Font, System.Drawing.FontStyle.Italic); // 设置字体为斜体
            textBoxCMD.ReadOnly = false;
            textBoxScript.Text = "";
        }

        private void Form9_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 取消订阅事件
            this.MouseDown -= Form9_MouseDown;
            foreach (Control c in this.Controls)
            {
                c.MouseDown -= Control_MouseDown;
            }
            foreach (Control c in flowLayoutPanelDay.Controls)
            {
                c.MouseDown -= Control_MouseDown;
            }
            foreach (Control c in flowLayoutPanelMonth.Controls)
            {
                c.MouseDown -= Control_MouseDown;
            }

            dataGridView1.CurrentCellDirtyStateChanged -= dataGridView1_CurrentCellDirtyStateChanged;
            dataGridView1.DataError -= DataGridView1_DataError;
            dataGridView1.CellValueChanged -= DataGridView1_CellValueChanged;
            dataGridView1.CellContentClick -= DataGridView1_CellContentClick;

            if (IsSheetExist("_定时任务"))
            {
                Worksheet worksheet = Globals.ThisAddIn.Application.ActiveWorkbook.Sheets["_定时任务"];
                if (worksheet.UsedRange.Rows.Count == 1)
                {
                    worksheet.Delete();
                }
            }
        }
    }
}
