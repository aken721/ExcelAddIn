using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SQLite;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelDatabaseSkill : ISkill
    {
        public string Name => "ExcelDatabase";
        public string Description => "数据库连接技能，支持连接多种数据库并提取数据到Excel";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "connect_database",
                    Description = "连接数据库并获取表名列表。支持SQL Server、MySQL、PostgreSQL、Access、SQLite等数据库。当用户要求连接数据库时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "dbType", new { type = "string", description = "数据库类型：sqlserver/mysql/postgresql/access/sqlite" } },
                                { "connectionString", new { type = "string", description = "数据库连接字符串或文件路径" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "dbType", "connectionString" }
                },
                new SkillTool
                {
                    Name = "execute_query",
                    Description = "执行SQL查询并将结果写入Excel。当用户要求查询数据库、执行SQL时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "dbType", new { type = "string", description = "数据库类型：sqlserver/mysql/postgresql/access/sqlite" } },
                                { "connectionString", new { type = "string", description = "数据库连接字符串" } },
                                { "query", new { type = "string", description = "SQL查询语句" } },
                                { "outputSheetName", new { type = "string", description = "输出工作表名称（可选，默认'查询结果'）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "dbType", "connectionString", "query" }
                },
                new SkillTool
                {
                    Name = "export_table_to_excel",
                    Description = "将数据库表导出到Excel工作表。当用户要求导出数据库表时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "dbType", new { type = "string", description = "数据库类型" } },
                                { "connectionString", new { type = "string", description = "数据库连接字符串" } },
                                { "tableName", new { type = "string", description = "要导出的表名" } },
                                { "outputSheetName", new { type = "string", description = "输出工作表名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "dbType", "connectionString", "tableName" }
                },
                new SkillTool
                {
                    Name = "get_table_structure",
                    Description = "获取数据库表的结构信息（字段名、类型等）。当用户要求查看表结构时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "dbType", new { type = "string", description = "数据库类型" } },
                                { "connectionString", new { type = "string", description = "数据库连接字符串" } },
                                { "tableName", new { type = "string", description = "表名" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "dbType", "connectionString", "tableName" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "connect_database":
                        return await ConnectDatabaseAsync(arguments);
                    case "execute_query":
                        return await ExecuteQueryAsync(arguments);
                    case "export_table_to_excel":
                        return await ExportTableToExcelAsync(arguments);
                    case "get_table_structure":
                        return await GetTableStructureAsync(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelDatabaseSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private async Task<SkillResult> ConnectDatabaseAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var dbType = arguments["dbType"].ToString().ToLower();
                var connectionString = arguments["connectionString"].ToString();

                try
                {
                    var tableNames = GetTableNames(dbType, connectionString);
                    return new SkillResult 
                    { 
                        Success = true, 
                        Content = $"数据库连接成功，共 {tableNames.Count} 张表：\n{string.Join("\n", tableNames)}" 
                    };
                }
                catch (Exception ex)
                {
                    return new SkillResult { Success = false, Error = $"数据库连接失败：{ex.Message}" };
                }
            });
        }

        private async Task<SkillResult> ExecuteQueryAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var dbType = arguments["dbType"].ToString().ToLower();
                var connectionString = arguments["connectionString"].ToString();
                var query = arguments["query"].ToString();
                var outputSheetName = arguments.ContainsKey("outputSheetName") 
                    ? arguments["outputSheetName"].ToString() 
                    : "查询结果";

                try
                {
                    var dataTable = ExecuteQuery(dbType, connectionString, query);
                    WriteDataTableToExcel(dataTable, outputSheetName);
                    return new SkillResult 
                    { 
                        Success = true, 
                        Content = $"查询执行成功，共 {dataTable.Rows.Count} 行数据已写入工作表 '{outputSheetName}'" 
                    };
                }
                catch (Exception ex)
                {
                    return new SkillResult { Success = false, Error = $"查询执行失败：{ex.Message}" };
                }
            });
        }

        private async Task<SkillResult> ExportTableToExcelAsync(Dictionary<string, object> arguments)
        {
            var dbType = arguments["dbType"].ToString().ToLower();
            var connectionString = arguments["connectionString"].ToString();
            var tableName = arguments["tableName"].ToString();
            var outputSheetName = arguments.ContainsKey("outputSheetName") 
                ? arguments["outputSheetName"].ToString() 
                : tableName;

            return await ExecuteQueryAsync(new Dictionary<string, object>
            {
                { "dbType", dbType },
                { "connectionString", connectionString },
                { "query", $"SELECT * FROM [{tableName}]" },
                { "outputSheetName", outputSheetName }
            });
        }

        private async Task<SkillResult> GetTableStructureAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                var dbType = arguments["dbType"].ToString().ToLower();
                var connectionString = arguments["connectionString"].ToString();
                var tableName = arguments["tableName"].ToString();

                try
                {
                    var query = dbType switch
                    {
                        "sqlserver" => $"SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tableName}'",
                        "mysql" => $"DESCRIBE {tableName}",
                        "postgresql" => $"SELECT column_name, data_type, character_maximum_length, is_nullable FROM information_schema.columns WHERE table_name = '{tableName}'",
                        "access" or "sqlite" => $"PRAGMA table_info({tableName})",
                        _ => throw new NotSupportedException($"不支持的数据库类型: {dbType}")
                    };

                    var dataTable = ExecuteQuery(dbType, connectionString, query);
                    WriteDataTableToExcel(dataTable, $"{tableName}_结构");
                    return new SkillResult 
                    { 
                        Success = true, 
                        Content = $"表 '{tableName}' 结构已导出，共 {dataTable.Rows.Count} 个字段" 
                    };
                }
                catch (Exception ex)
                {
                    return new SkillResult { Success = false, Error = $"获取表结构失败：{ex.Message}" };
                }
            });
        }

        private List<string> GetTableNames(string dbType, string connectionString)
        {
            var tableNames = new List<string>();

            switch (dbType)
            {
                case "access":
                    using (var conn = new OleDbConnection(connectionString))
                    {
                        conn.Open();
                        var schema = conn.GetSchema("Tables");
                        foreach (DataRow row in schema.Rows)
                        {
                            if (row["TABLE_TYPE"].ToString() == "TABLE")
                                tableNames.Add(row["TABLE_NAME"].ToString());
                        }
                    }
                    break;

                case "sqlite":
                    using (var conn = new SQLiteConnection(connectionString))
                    {
                        conn.Open();
                        using (var cmd = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table'", conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                                tableNames.Add(reader["name"].ToString());
                        }
                    }
                    break;

                case "sqlserver":
                case "mysql":
                case "postgresql":
                    var query = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'";
                    var dt = ExecuteQuery(dbType, connectionString, query);
                    foreach (DataRow row in dt.Rows)
                        tableNames.Add(row[0].ToString());
                    break;

                default:
                    throw new NotSupportedException($"不支持的数据库类型: {dbType}");
            }

            return tableNames;
        }

        private DataTable ExecuteQuery(string dbType, string connectionString, string query)
        {
            var dataTable = new DataTable();

            switch (dbType)
            {
                case "access":
                    using (var conn = new OleDbConnection(connectionString))
                    using (var adapter = new OleDbDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                case "sqlite":
                    using (var conn = new SQLiteConnection(connectionString))
                    using (var adapter = new SQLiteDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                case "sqlserver":
                    using (var conn = new System.Data.SqlClient.SqlConnection(connectionString))
                    using (var adapter = new System.Data.SqlClient.SqlDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                case "mysql":
                    using (var conn = new MySql.Data.MySqlClient.MySqlConnection(connectionString))
                    using (var adapter = new MySql.Data.MySqlClient.MySqlDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                case "postgresql":
                    using (var conn = new Npgsql.NpgsqlConnection(connectionString))
                    using (var adapter = new Npgsql.NpgsqlDataAdapter(query, conn))
                    {
                        adapter.Fill(dataTable);
                    }
                    break;

                default:
                    throw new NotSupportedException($"不支持的数据库类型: {dbType}");
            }

            return dataTable;
        }

        private void WriteDataTableToExcel(DataTable dataTable, string sheetName)
        {
            var workbook = ThisAddIn.app.ActiveWorkbook;

            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            Excel.Worksheet sheet;
            
            var existingNames = new List<string>();
            foreach (Excel.Worksheet ws in workbook.Worksheets) existingNames.Add(ws.Name);
            
            string actualSheetName = sheetName;
            if (existingNames.Any(n => string.Equals(n, sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                int suffix = 2;
                while (existingNames.Any(n => string.Equals(n, $"{sheetName}_{suffix}", StringComparison.OrdinalIgnoreCase)))
                {
                    suffix++;
                }
                actualSheetName = $"{sheetName}_{suffix}";
            }
            
            sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
            sheet.Name = actualSheetName;

            for (int c = 0; c < dataTable.Columns.Count; c++)
            {
                sheet.Cells[1, c + 1].Value = dataTable.Columns[c].ColumnName;
                sheet.Cells[1, c + 1].Font.Bold = true;
            }

            for (int r = 0; r < dataTable.Rows.Count; r++)
            {
                for (int c = 0; c < dataTable.Columns.Count; c++)
                {
                    sheet.Cells[r + 2, c + 1].Value = dataTable.Rows[r][c];
                }
            }

            sheet.UsedRange.Columns.AutoFit();
            sheet.Activate();

            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;
        }
    }
}
