using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelScheduleSkill : ISkill
    {
        public string Name => "ExcelSchedule";
        public string Description => "定时任务技能，支持创建和管理定时执行的任务";

        private List<ScheduledTask> _activeTasks = new List<ScheduledTask>();
        private static System.Timers.Timer _schedulerTimer;
        private static ExcelScheduleSkill _instance;

        public ExcelScheduleSkill()
        {
            _instance = this;
            if (_schedulerTimer == null)
            {
                _schedulerTimer = new System.Timers.Timer(60000);
                _schedulerTimer.Elapsed += SchedulerTimer_Elapsed;
                _schedulerTimer.Start();
            }
        }

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "create_scheduled_task",
                    Description = "创建定时任务。当用户要求创建定时任务、设置定时执行时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "taskName", new { type = "string", description = "任务名称" } },
                                { "taskDescription", new { type = "string", description = "任务描述（可选）" } },
                                { "startTime", new { type = "string", description = "开始时间（格式：yyyy-MM-dd HH:mm:ss）" } },
                                { "frequency", new { type = "string", description = "重复频率：一次性/每天/每周/每月" } },
                                { "interval", new { type = "integer", description = "间隔周期（天/周数，可选，默认1）" } },
                                { "weekDays", new { type = "string", description = "每周几执行（JSON数组，如[\"周一\",\"周三\"]，仅每周类型）" } },
                                { "months", new { type = "string", description = "月份安排（如\"1,3,6\"，仅每月类型）" } },
                                { "monthDays", new { type = "string", description = "日期安排（如\"1,15,最后一天\"，仅每月类型）" } },
                                { "taskType", new { type = "string", description = "任务类型：运行CMD命令/运行VBA宏/运行BAT批处理/运行Python脚本/启动程序" } },
                                { "command", new { type = "string", description = "任务内容（命令/宏名/文件路径）" } },
                                { "enabled", new { type = "boolean", description = "是否启用（默认true）" } },
                                { "stopTime", new { type = "string", description = "计划停止时间（格式：yyyy-MM-dd HH:mm:ss，可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "taskName", "startTime", "frequency", "taskType", "command" }
                },
                new SkillTool
                {
                    Name = "list_scheduled_tasks",
                    Description = "列出所有定时任务。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "delete_scheduled_task",
                    Description = "删除指定的定时任务。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "taskName", new { type = "string", description = "任务名称" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "taskName" }
                },
                new SkillTool
                {
                    Name = "enable_task",
                    Description = "启用定时任务。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "taskName", new { type = "string", description = "任务名称" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "taskName" }
                },
                new SkillTool
                {
                    Name = "disable_task",
                    Description = "禁用定时任务。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "taskName", new { type = "string", description = "任务名称" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "taskName" }
                },
                new SkillTool
                {
                    Name = "run_task_now",
                    Description = "立即执行指定的定时任务。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "taskName", new { type = "string", description = "任务名称" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "taskName" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "create_scheduled_task":
                        return CreateScheduledTask(arguments);
                    case "list_scheduled_tasks":
                        return ListScheduledTasks();
                    case "delete_scheduled_task":
                        return DeleteScheduledTask(arguments);
                    case "enable_task":
                        return EnableTask(arguments);
                    case "disable_task":
                        return DisableTask(arguments);
                    case "run_task_now":
                        return RunTaskNow(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelScheduleSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private void SchedulerTimer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                LoadActiveTasks();

                foreach (var task in _activeTasks.ToList())
                {
                    if (!task.IsEnabled || task.IsExpired) continue;
                    if (task.TaskStatus != "计划中") continue;

                    if (DateTime.Now >= task.NextRunTime.AddSeconds(-30) &&
                        DateTime.Now <= task.NextRunTime.AddSeconds(30))
                    {
                        string result = ExecuteTask(task);
                        task.NextRunTime = CalculateNextRunTime(task);
                        UpdateTaskStatus(task, result);
                        task.LastRunTime = DateTime.Now;
                    }
                }
            }
            catch { }
        }

        private SkillResult CreateScheduledTask(Dictionary<string, object> arguments)
        {
            var taskName = arguments["taskName"].ToString();
            var startTime = DateTime.Parse(arguments["startTime"].ToString());
            var frequency = arguments["frequency"].ToString();
            var taskType = arguments["taskType"].ToString();
            var command = arguments["command"].ToString();
            var taskDescription = arguments.ContainsKey("taskDescription") ? arguments["taskDescription"].ToString() : "";
            var interval = arguments.ContainsKey("interval") ? Convert.ToInt32(arguments["interval"]) : 1;
            var enabled = !arguments.ContainsKey("enabled") || Convert.ToBoolean(arguments["enabled"]);

            if (IsTaskNameExists(taskName))
                return new SkillResult { Success = false, Error = $"任务名称已存在: {taskName}" };

            var weekDays = arguments.ContainsKey("weekDays")
                ? Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["weekDays"].ToString()).ToArray()
                : Array.Empty<string>();
            var months = arguments.ContainsKey("months") ? arguments["months"].ToString() : "-";
            var monthDays = arguments.ContainsKey("monthDays") ? arguments["monthDays"].ToString() : "-";
            var stopTime = arguments.ContainsKey("stopTime") ? arguments["stopTime"].ToString() : "-";

            string intervalStr = frequency switch
            {
                "每天" => $"每{interval}天",
                "每周" => $"每{interval}周",
                _ => "-"
            };

            string weekDaysStr = frequency == "每周" ? string.Join(",", weekDays) : "-";
            string monthStr = frequency == "每月" ? months : "-";
            string monthDayStr = frequency == "每月" ? monthDays : "-";

            AddNewTaskToExcel(taskName, taskDescription, startTime, frequency, intervalStr,
                weekDaysStr, monthDayStr, taskType, command, enabled, stopTime);

            return new SkillResult { Success = true, Content = $"定时任务 '{taskName}' 创建成功\n类型: {taskType}\n调度: {frequency}\n开始时间: {startTime:yyyy-MM-dd HH:mm:ss}" };
        }

        private void AddNewTaskToExcel(string taskName, string taskDescription, DateTime startTime, string frequency,
            string intervalStr, string weekDaysStr, string monthDayStr, string taskType, string command,
            bool enabled, string stopTime)
        {
            string sheetName = "_定时任务";

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            try
            {
                if (!IsSheetExist(sheetName))
                {
                    Excel.Worksheet newSheet = ThisAddIn.app.ActiveWorkbook.Sheets.Add();
                    newSheet.Name = sheetName;
                    var headers = new[] {
                        "任务名称", "任务描述", "开始时间", "重复频率", "间隔周期",
                        "详细安排（月）", "详细安排（日）", "任务类型", "任务内容", "是否启用",
                        "计划停止时间", "实际执行时间", "下次执行时间", "上次运行结果", "任务状态"
                    };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        newSheet.Cells[1, i + 1].Value = headers[i];
                    }
                    newSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                }

                Excel.Worksheet sheet = ThisAddIn.app.ActiveWorkbook.Sheets[sheetName];
                int rowCount = sheet.UsedRange.Rows.Count;

                sheet.Cells[rowCount + 1, 1].Value = taskName;
                sheet.Cells[rowCount + 1, 2].Value = taskDescription;
                sheet.Cells[rowCount + 1, 3].Value = startTime.ToString("yyyy-MM-dd HH:mm:ss");
                sheet.Cells[rowCount + 1, 4].Value = frequency;
                sheet.Cells[rowCount + 1, 5].Value = intervalStr;
                sheet.Cells[rowCount + 1, 6].Value = weekDaysStr;
                sheet.Cells[rowCount + 1, 7].Value = monthDayStr;
                sheet.Cells[rowCount + 1, 8].Value = taskType;
                sheet.Cells[rowCount + 1, 9].Value = command;
                sheet.Cells[rowCount + 1, 10].Value = enabled ? "启用" : "不启用";
                sheet.Cells[rowCount + 1, 11].Value = stopTime;

                DateTime nextRunTime = CalculateInitialNextRun(startTime, null, frequency, intervalStr, weekDaysStr, months: weekDaysStr, monthDays: monthDayStr);
                sheet.Cells[rowCount + 1, 13].Value = nextRunTime.ToString("yyyy-MM-dd HH:mm:ss");
                sheet.Cells[rowCount + 1, 15].Value = "计划中";
            }
            finally
            {
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
            }
        }

        private SkillResult ListScheduledTasks()
        {
            if (!IsSheetExist("_定时任务"))
                return new SkillResult { Success = true, Content = "暂无定时任务" };

            var dt = ReadPlanSheet("_定时任务");
            if (dt == null || dt.Rows.Count == 0)
                return new SkillResult { Success = true, Content = "暂无定时任务" };

            var content = "定时任务列表：\n" + string.Join("\n", dt.Rows.Cast<DataRow>().Select(r =>
                $"  - {r["任务名称"]} [{r["任务类型"]}] {r["重复频率"]} " +
                $"开始: {r["开始时间"]} " +
                $"状态: {r["是否启用"]} " +
                $"执行结果: {r["上次运行结果"]}"));

            return new SkillResult { Success = true, Content = content };
        }

        private SkillResult DeleteScheduledTask(Dictionary<string, object> arguments)
        {
            var taskName = arguments["taskName"].ToString();

            if (!IsSheetExist("_定时任务"))
                return new SkillResult { Success = false, Error = "任务表不存在" };

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            try
            {
                Excel.Worksheet sheet = ThisAddIn.app.ActiveWorkbook.Sheets["_定时任务"];
                Excel.Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == taskName)
                    {
                        ((Excel.Range)sheet.Rows[i]).Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                        break;
                    }
                }
            }
            finally
            {
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
            }

            return new SkillResult { Success = true, Content = $"任务 '{taskName}' 已删除" };
        }

        private SkillResult EnableTask(Dictionary<string, object> arguments)
        {
            var taskName = arguments["taskName"].ToString();
            return UpdateTaskEnableStatus(taskName, true);
        }

        private SkillResult DisableTask(Dictionary<string, object> arguments)
        {
            var taskName = arguments["taskName"].ToString();
            return UpdateTaskEnableStatus(taskName, false);
        }

        private SkillResult UpdateTaskEnableStatus(string taskName, bool enable)
        {
            if (!IsSheetExist("_定时任务"))
                return new SkillResult { Success = false, Error = "任务表不存在" };

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            try
            {
                Excel.Worksheet sheet = ThisAddIn.app.ActiveWorkbook.Sheets["_定时任务"];
                Excel.Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == taskName)
                    {
                        sheet.Cells[i, 10].Value = enable ? "启用" : "不启用";
                        sheet.Cells[i, 15].Value = enable ? "计划中" : "已停用";
                        break;
                    }
                }
            }
            finally
            {
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
            }

            return new SkillResult { Success = true, Content = $"任务 '{taskName}' 已{(enable ? "启用" : "禁用")}" };
        }

        private SkillResult RunTaskNow(Dictionary<string, object> arguments)
        {
            var taskName = arguments["taskName"].ToString();

            LoadActiveTasks();
            var task = _activeTasks.FirstOrDefault(t => t.TaskName == taskName);
            if (task == null)
                return new SkillResult { Success = false, Error = $"任务不存在或未启用: {taskName}" };

            string result = ExecuteTask(task);
            UpdateTaskStatus(task, result);

            return new SkillResult { Success = true, Content = $"任务 '{taskName}' 已执行\n结果: {result}" };
        }

        private string ExecuteTask(ScheduledTask task)
        {
            string result = "成功";
            try
            {
                switch (task.TaskType)
                {
                    case "运行CMD命令":
                        System.Diagnostics.Process.Start("cmd.exe", "/C " + task.Command);
                        break;
                    case "运行VBA宏":
                        ThisAddIn.app.Run(task.Command);
                        break;
                    case "运行BAT批处理":
                        System.Diagnostics.Process.Start(task.Command);
                        break;
                    case "运行Python脚本":
                        System.Diagnostics.Process.Start("python", task.Command);
                        break;
                    case "启动程序":
                        System.Diagnostics.Process.Start(task.Command);
                        break;
                }
            }
            catch (Exception ex)
            {
                result = $"失败: {ex.Message}";
            }
            return result;
        }

        private void UpdateTaskStatus(ScheduledTask task, string result)
        {
            if (!IsSheetExist("_定时任务")) return;

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            try
            {
                Excel.Worksheet sheet = ThisAddIn.app.ActiveWorkbook.Sheets["_定时任务"];
                Excel.Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == task.TaskName)
                    {
                        sheet.Cells[i, 12].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        sheet.Cells[i, 14].Value = result;

                        bool isExpired = task.Frequency == "一次性";
                        if (isExpired)
                        {
                            sheet.Cells[i, 10].Value = "不启用";
                            sheet.Cells[i, 15].Value = "已失效";
                            sheet.Cells[i, 13].Value = "N/A";
                        }
                        else
                        {
                            task.NextRunTime = CalculateNextRunTime(task);
                            sheet.Cells[i, 13].Value = task.NextRunTime.ToString("yyyy-MM-dd HH:mm:ss");
                            sheet.Cells[i, 15].Value = "计划中";
                        }
                        break;
                    }
                }
            }
            finally
            {
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
            }
        }

        private void LoadActiveTasks()
        {
            _activeTasks.Clear();
            if (!IsSheetExist("_定时任务")) return;

            try
            {
                Excel.Worksheet sheet = ThisAddIn.app.ActiveWorkbook.Sheets["_定时任务"];
                Excel.Range usedRange = sheet.UsedRange;
                int rowCount = usedRange.Rows.Count;

                for (int i = 2; i <= rowCount; i++)
                {
                    var statusCell = usedRange.Cells[i, 10].Value?.ToString();
                    if (statusCell != "启用") continue;

                    var task = new ScheduledTask
                    {
                        TaskName = usedRange.Cells[i, 1].Value?.ToString() ?? "",
                        Frequency = usedRange.Cells[i, 4].Value?.ToString() ?? "",
                        TaskType = usedRange.Cells[i, 8].Value?.ToString() ?? "",
                        Command = usedRange.Cells[i, 9].Value?.ToString() ?? "",
                        IsEnabled = statusCell == "启用",
                        TaskStatus = usedRange.Cells[i, 15].Value?.ToString() ?? "计划中"
                    };

                    if (DateTime.TryParse(usedRange.Cells[i, 3].Value?.ToString(), out DateTime startTime))
                        task.StartTime = startTime;

                    if (DateTime.TryParse(usedRange.Cells[i, 12].Value?.ToString(), out DateTime lastRun))
                        task.LastRunTime = lastRun;

                    if (DateTime.TryParse(usedRange.Cells[i, 13].Value?.ToString(), out DateTime nextRun))
                        task.NextRunTime = nextRun;

                    if (task.TaskStatus == "已失效" || task.TaskStatus == "已停用") continue;

                    task.IsExpired = task.Frequency == "一次性" && task.LastRunTime.HasValue;

                    _activeTasks.Add(task);
                }
            }
            catch { }
        }

        private DateTime CalculateNextRunTime(ScheduledTask task)
        {
            return CalculateInitialNextRun(task.StartTime, task.LastRunTime, task.Frequency, "", "", "", "");
        }

        private DateTime CalculateInitialNextRun(DateTime startTime, DateTime? lastRunTime, string frequency,
            string intervalStr, string weekDaysStr, string months, string monthDays)
        {
            TimeSpan originalTime = startTime.TimeOfDay;
            DateTime now = DateTime.Now;

            bool isFirstRunAndNotStarted = (lastRunTime == null) && (now < startTime);
            if (isFirstRunAndNotStarted) return startTime;

            DateTime baseTime = lastRunTime ?? startTime;
            if (now < startTime) baseTime = startTime;

            int interval = 1;
            var match = System.Text.RegularExpressions.Regex.Match(intervalStr ?? "", @"\d+");
            if (match.Success) int.TryParse(match.Value, out interval);
            if (interval < 1) interval = 1;

            switch (frequency)
            {
                case "一次性":
                    return startTime;

                case "每天":
                    {
                        DateTime nextDate = baseTime.Date.AddDays(interval);
                        while (nextDate <= now.Date)
                        {
                            nextDate = nextDate.AddDays(interval);
                        }
                        return nextDate.Add(originalTime);
                    }

                case "每周":
                    {
                        var dayNameMapping = new Dictionary<string, DayOfWeek>
                        {
                            { "周一", DayOfWeek.Monday }, { "周二", DayOfWeek.Tuesday },
                            { "周三", DayOfWeek.Wednesday }, { "周四", DayOfWeek.Thursday },
                            { "周五", DayOfWeek.Friday }, { "周六", DayOfWeek.Saturday },
                            { "周日", DayOfWeek.Sunday }
                        };

                        var selectedDays = (weekDaysStr ?? "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                            .Where(d => dayNameMapping.ContainsKey(d.Trim()))
                            .Select(d => dayNameMapping[d.Trim()])
                            .ToList();

                        if (selectedDays.Count == 0) selectedDays.Add(startTime.DayOfWeek);

                        DateTime referenceDate = now > startTime ? now : startTime;
                        referenceDate = referenceDate.Date.Add(originalTime);

                        for (int i = 0; i < 52; i++)
                        {
                            DateTime periodStart = startTime.Date.AddDays(i * 7 * interval);
                            foreach (var day in selectedDays.OrderBy(d => (int)d))
                            {
                                DateTime candidate = periodStart.AddDays((int)day - (int)startTime.DayOfWeek);
                                if (candidate < periodStart) candidate = candidate.AddDays(7);
                                candidate = candidate.Date.Add(originalTime);
                                if (candidate >= referenceDate) return candidate;
                            }
                        }
                        return DateTime.MaxValue;
                    }

                case "每月":
                    return baseTime.AddMonths(1);
            }

            return startTime;
        }

        private bool IsTaskNameExists(string taskName)
        {
            if (!IsSheetExist("_定时任务")) return false;

            try
            {
                Excel.Worksheet sheet = ThisAddIn.app.ActiveWorkbook.Sheets["_定时任务"];
                Excel.Range usedRange = sheet.UsedRange;

                for (int i = 2; i <= usedRange.Rows.Count; i++)
                {
                    if (usedRange.Cells[i, 1].Value?.ToString() == taskName)
                        return true;
                }
            }
            catch { }

            return false;
        }

        private bool IsSheetExist(string sheetName)
        {
            try
            {
                foreach (Excel.Worksheet sheet in ThisAddIn.app.ActiveWorkbook.Sheets)
                {
                    if (sheet.Name == sheetName) return true;
                }
            }
            catch { }
            return false;
        }

        private DataTable ReadPlanSheet(string sheetName)
        {
            var dataTable = new DataTable();
            try
            {
                Excel.Worksheet sheet = ThisAddIn.app.ActiveWorkbook.Sheets[sheetName];
                Excel.Range usedRange = sheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int columnCount = usedRange.Columns.Count;

                for (int c = 1; c <= columnCount; c++)
                {
                    var colName = usedRange.Cells[1, c].Value?.ToString() ?? $"列{c}";
                    dataTable.Columns.Add(colName);
                }

                for (int r = 2; r <= rowCount; r++)
                {
                    var row = dataTable.NewRow();
                    for (int c = 1; c <= columnCount; c++)
                    {
                        row[c - 1] = usedRange.Cells[r, c].Value?.ToString() ?? "";
                    }
                    dataTable.Rows.Add(row);
                }
            }
            catch { }

            return dataTable;
        }

        private class ScheduledTask
        {
            public string TaskName { get; set; } = "";
            public DateTime StartTime { get; set; }
            public DateTime NextRunTime { get; set; } = DateTime.MaxValue;
            public DateTime? LastRunTime { get; set; }
            public string Frequency { get; set; } = "一次性";
            public string TaskType { get; set; } = "运行CMD命令";
            public string Command { get; set; } = "";
            public bool IsEnabled { get; set; }
            public bool IsExpired { get; set; }
            public string TaskStatus { get; set; } = "计划中";
        }
    }
}
