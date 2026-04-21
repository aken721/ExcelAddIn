using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace TableMagic.Skills
{
    public class ExcelSheetSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public ExcelSheetSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "ExcelSheet";
        public string Description => "Excel工作表技能，提供工作表管理功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "activate_worksheet",
                    Description = "激活/切换到指定工作表",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "要激活的工作表名称" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sheetName" }
                },
                new SkillTool
                {
                    Name = "create_worksheet",
                    Description = "在当前工作簿中创建新的工作表（Sheet）。当用户要求新建表、创建工作表、添加工作表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "新工作表的名称（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sheetName" }
                },
                new SkillTool
                {
                    Name = "rename_worksheet",
                    Description = "重命名工作表",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "oldSheetName", new { type = "string", description = "原工作表名称" } },
                                { "newSheetName", new { type = "string", description = "新工作表名称" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "oldSheetName", "newSheetName" }
                },
                new SkillTool
                {
                    Name = "delete_worksheet",
                    Description = "删除工作表",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "要删除的工作表名称" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sheetName" }
                },
                new SkillTool
                {
                    Name = "copy_worksheet",
                    Description = "复制工作表",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sourceSheetName", new { type = "string", description = "源工作表名称" } },
                                { "targetSheetName", new { type = "string", description = "目标工作表名称" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sourceSheetName", "targetSheetName" }
                },
                new SkillTool
                {
                    Name = "move_worksheet",
                    Description = "移动工作表到指定位置",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "工作表名称" } },
                                { "position", new { type = "integer", description = "目标位置（从1开始）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sheetName", "position" }
                },
                new SkillTool
                {
                    Name = "set_worksheet_visible",
                    Description = "设置工作表可见性",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "工作表名称" } },
                                { "visible", new { type = "boolean", description = "是否可见" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sheetName", "visible" }
                },
                new SkillTool
                {
                    Name = "get_worksheet_index",
                    Description = "获取工作表索引位置",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "工作表名称" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "sheetName" }
                },
                new SkillTool
                {
                    Name = "freeze_panes",
                    Description = "冻结窗格",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "rangeAddress", new { type = "string", description = "冻结位置，如'B2'表示冻结A列和第1行" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "rangeAddress" }
                },
                new SkillTool
                {
                    Name = "unfreeze_panes",
                    Description = "取消冻结窗格",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "activate_worksheet":
                        {
                            var sheetName = arguments["sheetName"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;

                            _excelMcp.ActivateWorksheet(fileName, sheetName);
                            return new SkillResult { Success = true, Content = "激活工作表成功" };
                        }
                    case "rename_worksheet":
                        {
                            var oldSheetName = arguments["oldSheetName"].ToString();
                            var newSheetName = arguments["newSheetName"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;

                            try
                            {
                                _excelMcp.RenameWorksheet(fileName, oldSheetName, newSheetName);
                                return new SkillResult { Success = true, Content = "重命名工作表成功" };
                            }
                            catch (ArgumentException ex) when (ex.Message.Contains("已存在"))
                            {
                                var existingNames = _excelMcp.GetWorksheetNames(fileName);
                                string existingActualName = existingNames.FirstOrDefault(n => string.Equals(n, newSheetName, StringComparison.OrdinalIgnoreCase)) ?? newSheetName;
                                string altName = newSheetName;
                                int suffix = 2;
                                while (existingNames.Any(n => string.Equals(n, altName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    altName = $"{newSheetName}_{suffix}";
                                    suffix++;
                                }
                                return SkillResult.FromError(
                                    $"工作表名称 '{newSheetName}' 已存在（实际名称：'{existingActualName}'），无法重命名",
                                    new List<string>
                                    {
                                        $"1. 使用其他名称重命名，例如：'{altName}'",
                                        $"2. 先删除现有的 '{existingActualName}' 工作表，再重命名",
                                        $"3. 重命名现有的 '{existingActualName}' 工作表为其他名称后再操作"
                                    },
                                    requiresUserDecision: true);
                            }
                            catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.Contains("已被使用") || ex.Message.Contains("already") || ex.Message.Contains("重名"))
                            {
                                var existingNames = _excelMcp.GetWorksheetNames(fileName);
                                string existingActualName = existingNames.FirstOrDefault(n => string.Equals(n, newSheetName, StringComparison.OrdinalIgnoreCase)) ?? newSheetName;
                                string altName = newSheetName;
                                int suffix = 2;
                                while (existingNames.Any(n => string.Equals(n, altName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    altName = $"{newSheetName}_{suffix}";
                                    suffix++;
                                }
                                return SkillResult.FromError(
                                    $"工作表名称 '{newSheetName}' 已存在（实际名称：'{existingActualName}'），无法重命名",
                                    new List<string>
                                    {
                                        $"1. 使用其他名称重命名，例如：'{altName}'",
                                        $"2. 先删除现有的 '{existingActualName}' 工作表，再重命名",
                                        $"3. 重命名现有的 '{existingActualName}' 工作表为其他名称后再操作"
                                    },
                                    requiresUserDecision: true);
                            }
                        }
                    case "delete_worksheet":
                        {
                            var sheetName = arguments["sheetName"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;

                            try
                            {
                                _excelMcp.DeleteWorksheet(fileName, sheetName);
                                return new SkillResult { Success = true, Content = "删除工作表成功" };
                            }
                            catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.Contains("不能删除") || ex.Message.Contains("cannot") || ex.Message.Contains("无法删除"))
                            {
                                return SkillResult.FromError(
                                    $"无法删除工作表 '{sheetName}'",
                                    new List<string>
                                    {
                                        "1. 工作簿至少需要保留一个可见工作表，请确认是否要删除",
                                        "2. 如果是隐藏工作表无法删除，请先取消隐藏再删除"
                                    },
                                    requiresUserDecision: true);
                            }
                        }
                    case "copy_worksheet":
                        {
                            var sourceSheetName = arguments["sourceSheetName"].ToString();
                            var targetSheetName = arguments["targetSheetName"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;

                            try
                            {
                                _excelMcp.CopyWorksheet(fileName, sourceSheetName, targetSheetName);
                                return new SkillResult { Success = true, Content = "复制工作表成功" };
                            }
                            catch (ArgumentException ex) when (ex.Message.Contains("已存在"))
                            {
                                var existingNames = _excelMcp.GetWorksheetNames(fileName);
                                string altName = targetSheetName;
                                int suffix = 2;
                                while (existingNames.Any(n => string.Equals(n, altName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    altName = $"{targetSheetName}_{suffix}";
                                    suffix++;
                                }
                                string existingActualName = existingNames.FirstOrDefault(n => string.Equals(n, targetSheetName, StringComparison.OrdinalIgnoreCase)) ?? targetSheetName;
                                return SkillResult.FromError(
                                    $"目标工作表名称 '{targetSheetName}' 已存在（实际名称：'{existingActualName}'），无法复制",
                                    new List<string>
                                    {
                                        $"1. 使用其他名称复制，例如：'{altName}'",
                                        $"2. 先删除现有的 '{existingActualName}' 工作表，再复制",
                                        $"3. 重命名现有的 '{existingActualName}' 工作表后再操作"
                                    },
                                    requiresUserDecision: true);
                            }
                            catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.Contains("已被使用") || ex.Message.Contains("already") || ex.Message.Contains("重名"))
                            {
                                var existingNames = _excelMcp.GetWorksheetNames(fileName);
                                string altName = targetSheetName;
                                int suffix = 2;
                                while (existingNames.Any(n => string.Equals(n, altName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    altName = $"{targetSheetName}_{suffix}";
                                    suffix++;
                                }
                                string existingActualName = existingNames.FirstOrDefault(n => string.Equals(n, targetSheetName, StringComparison.OrdinalIgnoreCase)) ?? targetSheetName;
                                return SkillResult.FromError(
                                    $"目标工作表名称 '{targetSheetName}' 已存在（实际名称：'{existingActualName}'），无法复制",
                                    new List<string>
                                    {
                                        $"1. 使用其他名称复制，例如：'{altName}'",
                                        $"2. 先删除现有的 '{existingActualName}' 工作表，再复制",
                                        $"3. 重命名现有的 '{existingActualName}' 工作表后再操作"
                                    },
                                    requiresUserDecision: true);
                            }
                        }
                    case "move_worksheet":
                        {
                            var sheetName = arguments["sheetName"].ToString();
                            var position = Convert.ToInt32(arguments["position"]);
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;

                            _excelMcp.MoveWorksheet(fileName, sheetName, position);
                            return new SkillResult { Success = true, Content = "移动工作表成功" };
                        }
                    case "set_worksheet_visible":
                        {
                            var sheetName = arguments["sheetName"].ToString();
                            var visible = Convert.ToBoolean(arguments["visible"]);
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;

                            _excelMcp.SetWorksheetVisible(fileName, sheetName, visible);
                            return new SkillResult { Success = true, Content = $"{ (visible ? "显示" : "隐藏") }工作表成功" };
                        }
                    case "create_worksheet":
                        {
                            string sheetName = null;
                            if (arguments.ContainsKey("name"))
                            {
                                sheetName = arguments["name"].ToString();
                            }
                            else if (arguments.ContainsKey("sheetName"))
                            {
                                sheetName = arguments["sheetName"].ToString();
                            }
                            
                            if (string.IsNullOrEmpty(sheetName))
                            {
                                return SkillResult.FromError("缺少必需参数：sheetName 或 name",
                                    new List<string> { "请提供工作表名称" });
                            }
                            
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;

                            try
                            {
                                _excelMcp.CreateWorksheet(fileName, sheetName);
                                return new SkillResult { Success = true, Content = $"创建工作表 '{sheetName}' 成功" };
                            }
                            catch (ArgumentException ex) when (ex.Message.Contains("已存在"))
                            {
                                var existingNames = _excelMcp.GetWorksheetNames(fileName);
                                string newName = sheetName;
                                int suffix = 2;
                                while (existingNames.Any(n => string.Equals(n, newName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    newName = $"{sheetName}_{suffix}";
                                    suffix++;
                                }
                                string existingActualName = existingNames.FirstOrDefault(n => string.Equals(n, sheetName, StringComparison.OrdinalIgnoreCase)) ?? sheetName;
                                return SkillResult.FromError(
                                    $"工作表名称 '{sheetName}' 已存在（实际名称：'{existingActualName}'），无法创建同名工作表",
                                    new List<string>
                                    {
                                        $"1. 使用其他名称创建，例如：'{newName}'",
                                        $"2. 先删除现有的 '{existingActualName}' 工作表，再重新创建",
                                        $"3. 重命名现有的 '{existingActualName}' 工作表为其他名称，再创建"
                                    },
                                    requiresUserDecision: true);
                            }
                            catch (System.Runtime.InteropServices.COMException ex) when (ex.Message.Contains("已被使用") || ex.Message.Contains("already") || ex.Message.Contains("重名"))
                            {
                                var existingNames = _excelMcp.GetWorksheetNames(fileName);
                                string newName = sheetName;
                                int suffix = 2;
                                while (existingNames.Any(n => string.Equals(n, newName, StringComparison.OrdinalIgnoreCase)))
                                {
                                    newName = $"{sheetName}_{suffix}";
                                    suffix++;
                                }
                                string existingActualName = existingNames.FirstOrDefault(n => string.Equals(n, sheetName, StringComparison.OrdinalIgnoreCase)) ?? sheetName;
                                return SkillResult.FromError(
                                    $"工作表名称 '{sheetName}' 已存在（实际名称：'{existingActualName}'），无法创建同名工作表",
                                    new List<string>
                                    {
                                        $"1. 使用其他名称创建，例如：'{newName}'",
                                        $"2. 先删除现有的 '{existingActualName}' 工作表，再重新创建",
                                        $"3. 重命名现有的 '{existingActualName}' 工作表为其他名称，再创建"
                                    },
                                    requiresUserDecision: true);
                            }
                        }
                    case "get_worksheet_index":
                        {
                            var sheetName = arguments["sheetName"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;

                            var index = _excelMcp.GetWorksheetIndex(fileName, sheetName);
                            return new SkillResult { Success = true, Content = $"工作表 '{sheetName}' 的索引位置是 {index}" };
                        }
                    case "freeze_panes":
                        {
                            var rangeAddress = arguments["rangeAddress"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            // 解析 rangeAddress 为 row 和 column
                            // 例如 "B2" 表示冻结第1行和第A列，即 row=2, column=2
                            int row = 1, column = 1;
                            if (!string.IsNullOrEmpty(rangeAddress))
                            {
                                var match = System.Text.RegularExpressions.Regex.Match(rangeAddress.ToUpper(), @"^([A-Z]+)(\d+)$");
                                if (match.Success)
                                {
                                    string colPart = match.Groups[1].Value;
                                    row = int.Parse(match.Groups[2].Value);
                                    // 将列字母转换为数字 (A=1, B=2, ...)
                                    column = 0;
                                    for (int i = 0; i < colPart.Length; i++)
                                    {
                                        column = column * 26 + (colPart[i] - 'A' + 1);
                                    }
                                }
                            }

                            _excelMcp.FreezePanes(fileName, sheetName, row, column);
                            return new SkillResult { Success = true, Content = $"冻结窗格成功，位置：{rangeAddress}（行={row}, 列={column}）" };
                        }
                    case "unfreeze_panes":
                        {
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            _excelMcp.UnfreezePanes(fileName, sheetName);
                            return new SkillResult { Success = true, Content = "取消冻结窗格成功" };
                        }
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelSheetSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }
    }
}