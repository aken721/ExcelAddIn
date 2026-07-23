using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelDatabaseSkill : ISkill
{
    private readonly IExcelProvider _provider;
    public ExcelDatabaseSkill(IExcelProvider provider) { _provider = provider; }
    public string Name => "ExcelDatabase";
    public string Description => "数据库连接技能，支持SQLite/MySQL/PostgreSQL/SQL Server";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new()
            {
                Name = "connect_database",
                Description = "连接数据库并获取表名列表。支持sqlserver/mysql/postgresql/sqlite",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "dbType", new { type = "string", description = "数据库类型：sqlserver/mysql/postgresql/sqlite" } },
                            { "connectionString", new { type = "string", description = "数据库连接字符串或文件路径" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "dbType", "connectionString" }
            },
            new()
            {
                Name = "execute_query",
                Description = "执行SQL查询并将结果写入Excel",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "dbType", new { type = "string", description = "数据库类型" } },
                            { "connectionString", new { type = "string", description = "连接字符串" } },
                            { "query", new { type = "string", description = "SQL查询语句" } },
                            { "outputFileName", new { type = "string", description = "输出工作簿文件名（可选）" } },
                            { "outputSheetName", new { type = "string", description = "输出工作表名称（默认'查询结果'）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "dbType", "connectionString", "query" }
            },
            new()
            {
                Name = "export_table_to_excel",
                Description = "将数据库表导出到Excel工作表",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "dbType", new { type = "string", description = "数据库类型" } },
                            { "connectionString", new { type = "string", description = "连接字符串" } },
                            { "tableName", new { type = "string", description = "要导出的表名" } },
                            { "outputFileName", new { type = "string", description = "输出工作簿文件名（可选）" } },
                            { "outputSheetName", new { type = "string", description = "输出工作表名称（可选）" } }
                        }
                    }
                },
                RequiredParameters = new List<string> { "dbType", "connectionString", "tableName" }
            },
            new()
            {
                Name = "get_table_structure",
                Description = "获取数据库表的结构信息（字段名、类型等）",
                Parameters = new Dictionary<string, object>
                {
                    { "type", "object" },
                    { "properties", new Dictionary<string, object>
                        {
                            { "dbType", new { type = "string", description = "数据库类型" } },
                            { "connectionString", new { type = "string", description = "连接字符串" } },
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
            return toolName switch
            {
                "connect_database" => await ConnectDatabaseAsync(arguments),
                "execute_query" => await ExecuteQueryAsync(arguments),
                "export_table_to_excel" => await ExportTableAsync(arguments),
                "get_table_structure" => await GetTableStructureAsync(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private async Task<SkillResult> ConnectDatabaseAsync(Dictionary<string, object> args)
    {
        return await Task.Run(() =>
        {
            var dbType = GetStr(args, "dbType")?.ToLower();
            var connStr = GetStr(args, "connectionString");
            var tables = GetTableNames(dbType!, connStr!);
            return SkillResult.Ok($"数据库连接成功，共 {tables.Count} 张表：\n{string.Join("\n", tables)}");
        });
    }

    private async Task<SkillResult> ExecuteQueryAsync(Dictionary<string, object> args)
    {
        return await Task.Run(() =>
        {
            var dbType = GetStr(args, "dbType")?.ToLower();
            var connStr = GetStr(args, "connectionString");
            var query = GetStr(args, "query");
            var outputFn = GetStr(args, "outputFileName");
            var outputSn = GetStr(args, "outputSheetName") ?? "查询结果";

            var dt = ExecuteDbQuery(dbType!, connStr!, query!);
            WriteDataTableToExcel(dt, outputFn, outputSn);
            return SkillResult.Ok($"查询执行成功，共 {dt.Rows.Count} 行数据已写入工作表 '{outputSn}'");
        });
    }

    private async Task<SkillResult> ExportTableAsync(Dictionary<string, object> args)
    {
        var newArgs = new Dictionary<string, object>(args)
        {
            ["query"] = $"SELECT * FROM [{GetStr(args, "tableName")}]"
        };
        if (!newArgs.ContainsKey("outputSheetName"))
            newArgs["outputSheetName"] = GetStr(args, "tableName") ?? "导出表";
        return await ExecuteQueryAsync(newArgs);
    }

    private async Task<SkillResult> GetTableStructureAsync(Dictionary<string, object> args)
    {
        return await Task.Run(() =>
        {
            var dbType = GetStr(args, "dbType")?.ToLower();
            var connStr = GetStr(args, "connectionString");
            var tableName = GetStr(args, "tableName");
            var query = dbType switch
            {
                "sqlite" => $"PRAGMA table_info({tableName})",
                "mysql" => $"DESCRIBE {tableName}",
                "postgresql" => $"SELECT column_name, data_type, character_maximum_length, is_nullable FROM information_schema.columns WHERE table_name = '{tableName}'",
                "sqlserver" => $"SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH, IS_NULLABLE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{tableName}'",
                _ => throw new NotSupportedException($"不支持的数据库类型: {dbType}")
            };
            var dt = ExecuteDbQuery(dbType, connStr, query);
            var outputFn = GetStr(args, "outputFileName");
            WriteDataTableToExcel(dt, outputFn, $"{tableName}_结构");
            return SkillResult.Ok($"表 '{tableName}' 结构已导出，共 {dt.Rows.Count} 个字段");
        });
    }

    private List<string> GetTableNames(string dbType, string connStr)
    {
        var tables = new List<string>();
        switch (dbType)
        {
            case "sqlite":
                using (var conn = new SQLiteConnection(connStr))
                {
                    conn.Open();
                    using var cmd = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table'", conn);
                    using var reader = cmd.ExecuteReader();
                    while (reader.Read()) tables.Add(reader["name"].ToString()!);
                }
                break;
            case "mysql":
            case "postgresql":
            case "sqlserver":
                var dt = ExecuteDbQuery(dbType, connStr, "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'");
                foreach (DataRow row in dt.Rows) tables.Add(row[0].ToString()!);
                break;
            default: throw new NotSupportedException($"不支持的数据库类型: {dbType}");
        }
        return tables;
    }

    private DataTable ExecuteDbQuery(string dbType, string connStr, string query)
    {
        var dt = new DataTable();
        switch (dbType)
        {
            case "sqlite":
                using (var conn = new SQLiteConnection(connStr))
                using (var adapter = new SQLiteDataAdapter(query, conn)) { adapter.Fill(dt); }
                break;
            case "mysql":
                using (var conn = new MySql.Data.MySqlClient.MySqlConnection(connStr))
                using (var adapter = new MySql.Data.MySqlClient.MySqlDataAdapter(query, conn)) { adapter.Fill(dt); }
                break;
            case "postgresql":
                using (var conn = new Npgsql.NpgsqlConnection(connStr))
                using (var adapter = new Npgsql.NpgsqlDataAdapter(query, conn)) { adapter.Fill(dt); }
                break;
            case "sqlserver":
                using (var conn = new Microsoft.Data.SqlClient.SqlConnection(connStr))
                using (var adapter = new Microsoft.Data.SqlClient.SqlDataAdapter(query, conn)) { adapter.Fill(dt); }
                break;
            default: throw new NotSupportedException($"不支持的数据库类型: {dbType}");
        }
        return dt;
    }

    private void WriteDataTableToExcel(DataTable dt, string fileName, string sheetName)
    {
        var fn = fileName ?? _provider.GetOpenWorkbooks().FirstOrDefault();
        if (fn == null)
        {
            fn = "query_result.xlsx";
            _provider.CreateWorkbook(fn, sheetName);
        }
        else
        {
            if (!_provider.GetOpenWorkbooks().Contains(fn))
                _provider.OpenWorkbook(fn);
            _provider.CreateWorksheet(fn, sheetName);
        }

        for (int c = 0; c < dt.Columns.Count; c++)
            _provider.SetCellValue(fn, sheetName, 1, c + 1, dt.Columns[c].ColumnName);

        for (int r = 0; r < dt.Rows.Count; r++)
            for (int c = 0; c < dt.Columns.Count; c++)
                _provider.SetCellValue(fn, sheetName, r + 2, c + 1, dt.Rows[r][c]);

        _provider.SaveWorkbook(fn);
    }

    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
}