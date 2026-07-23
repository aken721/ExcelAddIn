using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Builder;
using System.CommandLine.Parsing;
using System.IO;
using System.Linq;
using System.Text.Json;
using TableMagic.Cli.Excel;
using TableMagic.Cli.Mcp;
using TableMagic.Cli.Skills;

namespace TableMagic.Cli;

class Program
{
    static async Task<int> Main(string[] args)
    {
        var rootCommand = new RootCommand("TableMagic - Excel Skill CLI & MCP Server");

        var basePathOption = new Option<string>(
            name: "--base-path",
            description: "Excel文件基础路径",
            getDefaultValue: () => "./excel_files");

        var mcpCommand = new Command("mcp", "启动MCP服务器（stdio模式）");
        mcpCommand.AddOption(basePathOption);
        mcpCommand.SetHandler(async (string basePath) =>
        {
            await RunMcpServer(basePath);
        }, basePathOption);

        var listToolsCommand = new Command("list-tools", "列出所有可用的工具");
        listToolsCommand.AddOption(basePathOption);
        listToolsCommand.SetHandler((string basePath) =>
        {
            using var provider = new ClosedXmlExcelProvider(basePath);
            var manager = CreateSkillManager(provider);
            var tools = manager.GetAllTools();
            Console.WriteLine($"共 {tools.Count} 个工具:\n");
            foreach (var tool in tools)
            {
                Console.WriteLine($"  {tool.Name}: {tool.Description}");
                if (tool.RequiredParameters?.Count > 0)
                    Console.WriteLine($"    必需参数: {string.Join(", ", tool.RequiredParameters)}");
            }
        }, basePathOption);

        var callCommand = new Command("call", "调用指定工具");
        var toolNameArg = new Argument<string>("toolName", "工具名称");
        var argsOption = new Option<string>(
            name: "--args",
            description: "工具参数（JSON格式）",
            getDefaultValue: () => "{}");
        callCommand.AddArgument(toolNameArg);
        callCommand.AddOption(argsOption);
        callCommand.AddOption(basePathOption);
        callCommand.SetHandler(async (string toolName, string argsJson, string basePath) =>
        {
            using var provider = new ClosedXmlExcelProvider(basePath);
            var manager = CreateSkillManager(provider);
            var cleanJson = argsJson.TrimStart('\uFEFF');
            var arguments = JsonSerializer.Deserialize<Dictionary<string, object>>(cleanJson,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new();

            var result = await manager.ExecuteToolAsync(toolName, arguments);
            if (result.Success)
            {
                Console.WriteLine(result.Content);
            }
            else
            {
                Console.Error.WriteLine($"错误: {result.Error}");

                if (result.MissingRequiredParams && result.MissingParams.Count > 0)
                {
                    Console.Error.WriteLine();
                    Console.Error.WriteLine("缺少必需参数，请补充：");
                    foreach (var mp in result.MissingParams)
                    {
                        Console.Error.WriteLine($"  - {mp.Name} ({mp.Type}): {mp.PromptHint}");
                    }
                    Console.Error.WriteLine();
                    Console.Error.WriteLine($"示例: tablemagic call {toolName} --args '{{...}}'");
                }

                Environment.ExitCode = 1;
            }
        }, toolNameArg, argsOption, basePathOption);

        var batchCommand = new Command("batch", "批量执行工具调用（从JSON文件读取）");
        var inputFileArg = new Argument<string>("inputFile", "输入JSON文件路径");
        batchCommand.AddArgument(inputFileArg);
        batchCommand.AddOption(basePathOption);
        batchCommand.SetHandler(async (string inputFile, string basePath) =>
        {
            using var provider = new ClosedXmlExcelProvider(basePath);
            var manager = CreateSkillManager(provider);
            var json = await File.ReadAllTextAsync(inputFile);
            var cleanJson = json.TrimStart('\uFEFF');
            var calls = JsonSerializer.Deserialize<List<BatchCall>>(cleanJson,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new();

            foreach (var call in calls)
            {
                Console.WriteLine($"执行: {call.Tool}");
                var result = await manager.ExecuteToolAsync(call.Tool, call.Arguments ?? new());
                Console.WriteLine(result.Success ? result.Content : $"错误: {result.Error}");
                Console.WriteLine();
            }
        }, inputFileArg, basePathOption);

        rootCommand.AddCommand(mcpCommand);
        rootCommand.AddCommand(listToolsCommand);
        rootCommand.AddCommand(callCommand);
        rootCommand.AddCommand(batchCommand);

        return await new CommandLineBuilder(rootCommand)
            .UseVersionOption("--version", "-v")
            .UseHelp("--help", "-h", "-?")

            .UseParseErrorReporting()
            .UseExceptionHandler()
            .Build()
            .InvokeAsync(args);
    }

    private static SkillManager CreateSkillManager(ClosedXmlExcelProvider provider)
    {
        var manager = new SkillManager();

        manager.LoadSkill(new ExcelBaseSkill(provider, manager));
        manager.LoadSkill(new ExcelWorkbookSkill(provider));
        manager.LoadSkill(new ExcelSheetSkill(provider));
        manager.LoadSkill(new ExcelCellSkill(provider));
        manager.LoadSkill(new ExcelRangeSkill(provider));
        manager.LoadSkill(new ExcelFormatSkill(provider));
        manager.LoadSkill(new ExcelChartSkill(provider));
        manager.LoadSkill(new ExcelPivotSkill(provider));
        manager.LoadSkill(new ExcelAnalysisSkill(provider));
        manager.LoadSkill(new ExcelFinanceSkill(provider));
        manager.LoadSkill(new ExcelDataSkill(provider));
        manager.LoadSkill(new ExcelDatabaseSkill(provider));
        manager.LoadSkill(new ExcelApiSkill(provider));
        manager.LoadSkill(new ExcelChartEnhanceSkill(provider));
        manager.LoadSkill(new ExcelMailSkill());
        manager.LoadSkill(new ExcelFileSkill());
        manager.LoadSkill(new ExcelQRSkill(provider));
        manager.LoadSkill(new ExcelInvoiceSkill(provider));
        manager.LoadSkill(new ExcelRegexSkill(provider));
        manager.LoadSkill(new ExcelTocSkill(provider));
        manager.LoadSkill(new ExcelScheduleSkill());
        manager.LoadSkill(new DocumentGenerationSkill());
        manager.LoadSkill(new ExcelPdfSkill(provider));

        return manager;
    }

    private static async Task RunMcpServer(string basePath)
    {
        using var provider = new ClosedXmlExcelProvider(basePath);
        var manager = CreateSkillManager(provider);
        var server = new McpServer(manager);
        await server.RunAsync();
    }

    private class BatchCall
    {
        public string Tool { get; set; } = "";
        public Dictionary<string, object>? Arguments { get; set; }
    }
}