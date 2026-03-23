using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelAddIn.Skills
{
    public class ExcelFinanceSkill : ISkill
    {
        private ExcelMcp _excelMcp;

        public ExcelFinanceSkill(ExcelMcp excelMcp)
        {
            _excelMcp = excelMcp;
        }

        public string Name => "ExcelFinance";
        public string Description => "Excel财务分析技能，提供财务指标计算等功能";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "calculate_financial_ratio",
                    Description = "计算财务比率",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选，默认使用当前活跃工作表）" } },
                                { "revenueRange", new { type = "string", description = "收入数据范围，如B2:B12" } },
                                { "costRange", new { type = "string", description = "成本数据范围，如C2:C12" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "revenueRange", "costRange" }
                },
                new SkillTool
                {
                    Name = "calculate_profit_margin",
                    Description = "计算利润率",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "fileName", new { type = "string", description = "工作簿文件名（可选，默认使用当前活跃工作簿）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选，默认使用当前活跃工作表）" } },
                                { "revenueRange", new { type = "string", description = "收入数据范围，如B2:B12" } },
                                { "profitRange", new { type = "string", description = "利润数据范围，如D2:D12" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "revenueRange", "profitRange" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "calculate_financial_ratio":
                        {
                            var revenueRange = arguments["revenueRange"].ToString();
                            var costRange = arguments["costRange"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            var revenueData = _excelMcp.GetRangeValues(fileName, sheetName, revenueRange);
                            var costData = _excelMcp.GetRangeValues(fileName, sheetName, costRange);

                            var result = CalculateFinancialRatio(revenueData, costData);
                            return new SkillResult { Success = true, Content = result };
                        }
                    case "calculate_profit_margin":
                        {
                            var revenueRange = arguments["revenueRange"].ToString();
                            var profitRange = arguments["profitRange"].ToString();
                            var fileName = arguments.ContainsKey("fileName") ? arguments["fileName"].ToString() : null;
                            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;

                            var revenueData = _excelMcp.GetRangeValues(fileName, sheetName, revenueRange);
                            var profitData = _excelMcp.GetRangeValues(fileName, sheetName, profitRange);

                            var result = CalculateProfitMargin(revenueData, profitData);
                            return new SkillResult { Success = true, Content = result };
                        }
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelFinanceSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private string CalculateFinancialRatio(object[,] revenueData, object[,] costData)
        {
            if (revenueData == null || costData == null)
            {
                return "数据为空";
            }

            double totalRevenue = 0;
            double totalCost = 0;

            // 计算总收入
            for (int i = 0; i < revenueData.GetLength(0); i++)
            {
                for (int j = 0; j < revenueData.GetLength(1); j++)
                {
                    if (revenueData[i, j] is double d) totalRevenue += d;
                    else if (revenueData[i, j] is int n) totalRevenue += n;
                }
            }

            // 计算总成本
            for (int i = 0; i < costData.GetLength(0); i++)
            {
                for (int j = 0; j < costData.GetLength(1); j++)
                {
                    if (costData[i, j] is double d) totalCost += d;
                    else if (costData[i, j] is int n) totalCost += n;
                }
            }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"总收入: {totalRevenue}");
            sb.AppendLine($"总成本: {totalCost}");
            sb.AppendLine($"毛利润: {totalRevenue - totalCost}");
            if (totalRevenue > 0)
            {
                sb.AppendLine($"毛利率: {((totalRevenue - totalCost) / totalRevenue) * 100:F2}%");
            }

            return sb.ToString();
        }

        private string CalculateProfitMargin(object[,] revenueData, object[,] profitData)
        {
            if (revenueData == null || profitData == null)
            {
                return "数据为空";
            }

            double totalRevenue = 0;
            double totalProfit = 0;

            // 计算总收入
            for (int i = 0; i < revenueData.GetLength(0); i++)
            {
                for (int j = 0; j < revenueData.GetLength(1); j++)
                {
                    if (revenueData[i, j] is double d) totalRevenue += d;
                    else if (revenueData[i, j] is int n) totalRevenue += n;
                }
            }

            // 计算总利润
            for (int i = 0; i < profitData.GetLength(0); i++)
            {
                for (int j = 0; j < profitData.GetLength(1); j++)
                {
                    if (profitData[i, j] is double d) totalProfit += d;
                    else if (profitData[i, j] is int n) totalProfit += n;
                }
            }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"总收入: {totalRevenue}");
            sb.AppendLine($"总利润: {totalProfit}");
            if (totalRevenue > 0)
            {
                sb.AppendLine($"利润率: {((totalProfit) / totalRevenue) * 100:F2}%");
            }

            return sb.ToString();
        }
    }
}