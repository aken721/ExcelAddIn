using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelSpotlightSkill : ISkill
    {
        public string Name => "ExcelSpotlight";
        public string Description => "聚光灯技能，高亮显示当前选中单元格所在的行和列";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "enable_spotlight",
                    Description = "开启聚光灯效果，高亮显示当前选中单元格所在的整行和整列。当用户要求开启聚光灯、打开聚光灯、高亮行列时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "colorIndex", new { type = "integer", description = "聚光灯颜色索引（35=浅绿, 37=浅蓝, 24=浅紫, 36=浅黄, 15=浅灰, 44=浅橙，默认35）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "disable_spotlight",
                    Description = "关闭聚光灯效果。当用户要求关闭聚光灯时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "toggle_spotlight",
                    Description = "切换聚光灯效果开关状态。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "get_spotlight_status",
                    Description = "获取聚光灯当前状态。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "set_spotlight_color",
                    Description = "设置聚光灯颜色。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "colorIndex", new { type = "integer", description = "聚光灯颜色索引（35=浅绿, 37=浅蓝, 24=浅紫, 36=浅黄, 15=浅灰, 44=浅橙）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "colorIndex" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "enable_spotlight":
                        return EnableSpotlight(arguments);
                    case "disable_spotlight":
                        return DisableSpotlight();
                    case "toggle_spotlight":
                        return ToggleSpotlight();
                    case "get_spotlight_status":
                        return GetSpotlightStatus();
                    case "set_spotlight_color":
                        return SetSpotlightColor(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelSpotlightSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private SkillResult EnableSpotlight(Dictionary<string, object> arguments)
        {
            if (ThisAddIn.Global.spotlight == 1)
                return new SkillResult { Success = true, Content = "聚光灯已经处于开启状态" };

            if (arguments.ContainsKey("colorIndex"))
            {
                ThisAddIn.Global.spotlightColorIndex = Convert.ToInt32(arguments["colorIndex"]);
                UpdateSpotlightDropDown(ThisAddIn.Global.spotlightColorIndex);
            }

            Excel.Worksheet currentWorksheet = ThisAddIn.app.ActiveSheet;
            Excel.Range usedRange = currentWorksheet.UsedRange;

            ThisAddIn.Global.cellColor = ThisAddIn.Global.GetColorDictionary(usedRange);

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.Global.spotlight = 1;

            Excel.Range activeCell = ThisAddIn.app.ActiveCell;
            currentWorksheet.Cells.Interior.ColorIndex = 0;
            activeCell.EntireRow.Interior.ColorIndex = ThisAddIn.Global.spotlightColorIndex;
            activeCell.EntireColumn.Interior.ColorIndex = ThisAddIn.Global.spotlightColorIndex;
            ThisAddIn.app.ScreenUpdating = true;

            UpdateRibbonButton(true);

            return new SkillResult { Success = true, Content = "聚光灯已开启" };
        }

        private SkillResult DisableSpotlight()
        {
            if (ThisAddIn.Global.spotlight == 0)
                return new SkillResult { Success = true, Content = "聚光灯已经处于关闭状态" };

            ThisAddIn.Global.spotlight = 0;

            Excel.Worksheet currentWorksheet = ThisAddIn.app.ActiveSheet;
            Excel.Range selectCell = ThisAddIn.app.ActiveCell;

            ThisAddIn.app.ScreenUpdating = false;
            selectCell.EntireRow.Interior.ColorIndex = 0;
            selectCell.EntireColumn.Interior.ColorIndex = 0;

            if (ThisAddIn.Global.cellColor.Count > 0)
            {
                foreach (var cellColorEntry in ThisAddIn.Global.cellColor)
                {
                    string cellAddress = cellColorEntry.Key;
                    int cellColorIndex = cellColorEntry.Value;
                    Excel.Range cell = currentWorksheet.Range[cellAddress];
                    cell.Interior.ColorIndex = cellColorIndex;
                }
            }
            ThisAddIn.Global.cellColor.Clear();

            ThisAddIn.app.ScreenUpdating = true;

            UpdateRibbonButton(false);

            return new SkillResult { Success = true, Content = "聚光灯已关闭" };
        }

        private SkillResult ToggleSpotlight()
        {
            if (ThisAddIn.Global.spotlight == 1)
                return DisableSpotlight();
            else
                return EnableSpotlight(new Dictionary<string, object>());
        }

        private SkillResult GetSpotlightStatus()
        {
            var status = ThisAddIn.Global.spotlight == 1 ? "已开启" : "已关闭";
            var colorName = ThisAddIn.Global.spotlightColorIndex switch
            {
                35 => "浅绿",
                37 => "浅蓝",
                24 => "浅紫",
                36 => "浅黄",
                15 => "浅灰",
                44 => "浅橙",
                _ => $"自定义(索引{ThisAddIn.Global.spotlightColorIndex})"
            };

            return new SkillResult { Success = true, Content = $"聚光灯状态: {status}\n当前颜色: {colorName}" };
        }

        private SkillResult SetSpotlightColor(Dictionary<string, object> arguments)
        {
            var colorIndex = Convert.ToInt32(arguments["colorIndex"]);
            ThisAddIn.Global.spotlightColorIndex = colorIndex;
            UpdateSpotlightDropDown(colorIndex);

            if (ThisAddIn.Global.spotlight == 1)
            {
                Excel.Worksheet currentWorksheet = ThisAddIn.app.ActiveSheet;
                Excel.Range activeCell = ThisAddIn.app.ActiveCell;

                ThisAddIn.app.ScreenUpdating = false;
                currentWorksheet.Cells.Interior.ColorIndex = 0;
                activeCell.EntireRow.Interior.ColorIndex = ThisAddIn.Global.spotlightColorIndex;
                activeCell.EntireColumn.Interior.ColorIndex = ThisAddIn.Global.spotlightColorIndex;
                ThisAddIn.app.ScreenUpdating = true;
            }

            var colorName = colorIndex switch
            {
                35 => "浅绿",
                37 => "浅蓝",
                24 => "浅紫",
                36 => "浅黄",
                15 => "浅灰",
                44 => "浅橙",
                _ => $"自定义(索引{colorIndex})"
            };

            return new SkillResult { Success = true, Content = $"聚光灯颜色已设置为: {colorName}" };
        }

        private void UpdateRibbonButton(bool isOn)
        {
            try
            {
                var ribbon = Globals.Ribbons.OfType<Ribbon1>().FirstOrDefault();
                if (ribbon != null)
                {
                    ribbon.confirm_spotlight.Checked = isOn;
                    ribbon.confirm_spotlight.Label = isOn ? "关闭聚光灯" : "打开聚光灯";
                    ribbon.confirm_spotlight.Image = isOn
                        ? TableMagic.Properties.Resources.spotlight_open
                        : TableMagic.Properties.Resources.spotlight_close;
                }
            }
            catch { }
        }

        private void UpdateSpotlightDropDown(int colorIndex)
        {
            try
            {
                var ribbon = Globals.Ribbons.OfType<Ribbon1>().FirstOrDefault();
                if (ribbon != null && ribbon.spotlightDropDown != null)
                {
                    int selectedIndex = colorIndex switch
                    {
                        35 => 0,
                        37 => 1,
                        24 => 2,
                        36 => 3,
                        15 => 4,
                        44 => 5,
                        _ => 0
                    };
                    ribbon.spotlightDropDown.SelectedItemIndex = selectedIndex;
                }
            }
            catch { }
        }
    }
}
