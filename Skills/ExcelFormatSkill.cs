using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelAddIn.Skills
{
    public class ExcelFormatSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public ExcelFormatSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "ExcelFormat";
        public string Description => "Excel格式化技能，提供单元格格式化、边框设置等功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "set_cell_format",
                    Description = "设置单元格格式。当用户要求设置字体颜色、背景色、字号、加粗、斜体、对齐方式时使用此工具。注意：如果不指定sheetName，则使用当前活动的工作表。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，不指定则使用当前工作簿）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选，不指定则使用当前活动的工作表）" } },
                                { "rangeAddress", new { type = "string", description = "单元格区域地址（必需）" } },
                                { "fontColor", new { type = "string", description = "字体颜色（可选），如红色、#FF0000" } },
                                { "backgroundColor", new { type = "string", description = "背景色（可选），如黄色、#FFFF00" } },
                                { "fontSize", new { type = "integer", description = "字号（可选），如12" } },
                                { "bold", new { type = "boolean", description = "是否加粗（可选）" } },
                                { "italic", new { type = "boolean", description = "是否斜体（可选）" } },
                                { "horizontalAlignment", new { type = "string", description = "水平对齐（可选）：left/center/right" } },
                                { "verticalAlignment", new { type = "string", description = "垂直对齐（可选）：top/center/bottom" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "rangeAddress" }
                },
                new SkillTool
                {
                    Name = "set_border",
                    Description = "设置边框。当用户要求添加边框、设置边框线、画边框时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "rangeAddress", new { type = "string", description = "单元格区域地址（必需）" } },
                                { "borderType", new { type = "string", description = "边框类型（必需）：all/left/right/top/bottom/outline" } },
                                { "lineStyle", new { type = "string", description = "线条样式（可选）：solid/dashed/dotted" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "rangeAddress", "borderType" }
                },
                new SkillTool
                {
                    Name = "merge_cells",
                    Description = "合并单元格。当用户要求合并单元格、合并区域时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "rangeAddress", new { type = "string", description = "单元格区域地址（必需），如A1:C3" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "rangeAddress" }
                },
                new SkillTool
                {
                    Name = "unmerge_cells",
                    Description = "取消合并单元格。当用户要求取消合并、拆分单元格时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "rangeAddress", new { type = "string", description = "单元格区域地址（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "rangeAddress" }
                },
                new SkillTool
                {
                    Name = "set_cell_text_wrap",
                    Description = "设置单元格自动换行。当用户要求自动换行、设置换行时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "rangeAddress", new { type = "string", description = "单元格区域地址（必需）" } },
                                { "wrap", new { type = "boolean", description = "是否自动换行（必需）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "rangeAddress", "wrap" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "set_cell_format":
                        {
                            var rangeAddress = arguments["rangeAddress"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                            var fontColor = arguments.ContainsKey("fontColor") ? arguments["fontColor"].ToString() : null;
                            var backgroundColor = arguments.ContainsKey("backgroundColor") ? arguments["backgroundColor"].ToString() : null;
                            var fontSize = arguments.ContainsKey("fontSize") ? Convert.ToInt32(arguments["fontSize"]) : (int?)null;
                            var bold = arguments.ContainsKey("bold") ? Convert.ToBoolean(arguments["bold"]) : (bool?)null;
                            var italic = arguments.ContainsKey("italic") ? Convert.ToBoolean(arguments["italic"]) : (bool?)null;
                            var horizontalAlignment = arguments.ContainsKey("horizontalAlignment") ? arguments["horizontalAlignment"].ToString() : null;
                            var verticalAlignment = arguments.ContainsKey("verticalAlignment") ? arguments["verticalAlignment"].ToString() : null;

                            _excelMcp.SetCellFormat(fileName, sheetName, rangeAddress, fontColor, backgroundColor, fontSize, bold, italic, horizontalAlignment, verticalAlignment);
                            return new SkillResult { Success = true, Content = "设置单元格格式成功" };
                        }
                    case "set_border":
                        {
                            var rangeAddress = arguments["rangeAddress"].ToString();
                            var borderType = arguments["borderType"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
                            var lineStyle = arguments.ContainsKey("lineStyle") ? arguments["lineStyle"].ToString() : null;

                            _excelMcp.SetBorder(fileName, sheetName, rangeAddress, borderType, lineStyle);
                            return new SkillResult { Success = true, Content = "设置边框成功" };
                        }
                    case "merge_cells":
                        {
                            var rangeAddress = arguments["rangeAddress"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            _excelMcp.MergeCells(fileName, sheetName, rangeAddress);
                            return new SkillResult { Success = true, Content = "合并单元格成功" };
                        }
                    case "unmerge_cells":
                        {
                            var rangeAddress = arguments["rangeAddress"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            _excelMcp.UnmergeCells(fileName, sheetName, rangeAddress);
                            return new SkillResult { Success = true, Content = "取消合并单元格成功" };
                        }
                    case "set_cell_text_wrap":
                        {
                            var rangeAddress = arguments["rangeAddress"].ToString();
                            var wrap = Convert.ToBoolean(arguments["wrap"]);
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            _excelMcp.SetCellTextWrap(fileName, sheetName, rangeAddress, wrap);
                            return new SkillResult { Success = true, Content = "设置自动换行成功" };
                        }
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelFormatSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }
    }
}