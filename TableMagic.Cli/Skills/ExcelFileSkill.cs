using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelFileSkill : ISkill
{
    public string Name => "ExcelFile";
    public string Description => "文件/文件夹操作：批量重命名、复制、移动、删除、列表";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "list_files", Description = "列出文件夹中的文件信息",
                Parameters = P(new[]{"folderPath"}, new[]{"pattern","includeSubfolders"}), RequiredParameters = new List<string>{"folderPath"} },
            new() { Name = "batch_rename", Description = "批量重命名文件",
                Parameters = P(new[]{"oldNameColumn","newNameColumn"}, new[]{"folderPath","fileName","sheetName"}), RequiredParameters = new List<string>{"oldNameColumn","newNameColumn"} },
            new() { Name = "batch_copy", Description = "批量复制文件到目标文件夹",
                Parameters = P(new[]{"fileNameColumn","targetFolder"}, new[]{"sourceFolder","fileName","sheetName"}), RequiredParameters = new List<string>{"fileNameColumn","targetFolder"} },
            new() { Name = "batch_move", Description = "批量移动文件到目标文件夹",
                Parameters = P(new[]{"fileNameColumn","targetFolder"}, new[]{"sourceFolder","fileName","sheetName"}), RequiredParameters = new List<string>{"fileNameColumn","targetFolder"} },
            new() { Name = "batch_delete", Description = "批量删除文件",
                Parameters = P(new[]{"fileNameColumn"}, new[]{"folderPath","fileName","sheetName"}), RequiredParameters = new List<string>{"fileNameColumn"} },
            new() { Name = "create_folder", Description = "创建文件夹",
                Parameters = P(new[]{"folderPath"}, Array.Empty<string>()), RequiredParameters = new List<string>{"folderPath"} },
            new() { Name = "get_file_info", Description = "获取文件详细信息",
                Parameters = P(new[]{"filePath"}, Array.Empty<string>()), RequiredParameters = new List<string>{"filePath"} },
            new() { Name = "open_folder", Description = "在资源管理器中打开文件夹",
                Parameters = P(new[]{"folderPath"}, Array.Empty<string>()), RequiredParameters = new List<string>{"folderPath"} }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "list_files" => ListFiles(arguments),
                "batch_rename" => BatchRename(arguments),
                "batch_copy" => BatchCopy(arguments),
                "batch_move" => BatchMove(arguments),
                "batch_delete" => BatchDelete(arguments),
                "create_folder" => CreateFolder(arguments),
                "get_file_info" => GetFileInfo(arguments),
                "open_folder" => OpenFolder(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult ListFiles(Dictionary<string, object> args)
    {
        var folder = GetStr(args, "folderPath");
        var pattern = GetStr(args, "pattern") ?? "*.*";
        var includeSub = GetBool(args, "includeSubfolders");
        if (!Directory.Exists(folder)) return new SkillResult { Success = false, Error = $"文件夹不存在: {folder}" };

        var option = includeSub ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
        var files = Directory.GetFiles(folder, pattern, option);
        var sb = new System.Text.StringBuilder();
        sb.AppendLine($"共 {files.Length} 个文件:");
        foreach (var f in files.Take(100))
        {
            var fi = new FileInfo(f);
            sb.AppendLine($"  {fi.Name} | {Math.Round(fi.Length / 1024.0, 2)}KB | {fi.LastWriteTime:yyyy-MM-dd HH:mm}");
        }
        if (files.Length > 100) sb.AppendLine($"  ... 还有 {files.Length - 100} 个文件");
        return SkillResult.Ok(sb.ToString());
    }

    private SkillResult BatchRename(Dictionary<string, object> args) => BatchFileOp(args, "重命名", (oldPath, newPath) => { if (File.Exists(oldPath)) File.Move(oldPath, newPath); else if (Directory.Exists(oldPath)) Directory.Move(oldPath, newPath); });
    private SkillResult BatchCopy(Dictionary<string, object> args) => BatchFileOp(args, "复制", (oldPath, newPath) => { if (File.Exists(oldPath)) File.Copy(oldPath, newPath, true); });
    private SkillResult BatchMove(Dictionary<string, object> args) => BatchFileOp(args, "移动", (oldPath, newPath) => { if (File.Exists(oldPath)) File.Move(oldPath, newPath); });
    private SkillResult BatchDelete(Dictionary<string, object> args) => BatchFileOp(args, "删除", (oldPath, _) => { if (File.Exists(oldPath)) File.Delete(oldPath); });

    private SkillResult BatchFileOp(Dictionary<string, object> args, string opName, Action<string, string> action)
    {
        var folder = GetStr(args, "folderPath") ?? GetStr(args, "sourceFolder");
        int success = 0, fail = 0;
        var errors = new List<string>();

        var oldCol = GetStr(args, "oldNameColumn") ?? GetStr(args, "fileNameColumn");
        var newCol = GetStr(args, "newNameColumn");
        if (oldCol == null) return new SkillResult { Success = false, Error = "未指定列名" };

        var files = new List<(string Old, string New)>();
        if (args.ContainsKey("fileList"))
        {
            var list = args["fileList"] as List<string> ?? new List<string>();
            foreach (var f in list) files.Add((f, f));
        }

        foreach (var (old, @new) in files)
        {
            try
            {
                var oldPath = folder != null ? Path.Combine(folder, old) : old;
                var newPath = newCol != null ? Path.Combine(folder ?? "", @new) : oldPath;
                action(oldPath, newPath);
                success++;
            }
            catch (Exception ex) { fail++; errors.Add($"{old}: {ex.Message}"); }
        }

        var msg = $"{opName}完成：成功 {success}，失败 {fail}";
        if (errors.Count > 0) msg += $"\n错误: {string.Join("; ", errors.Take(5))}";
        return SkillResult.Ok(msg);
    }

    private SkillResult CreateFolder(Dictionary<string, object> args)
    {
        var path = GetStr(args, "folderPath");
        Directory.CreateDirectory(path);
        return SkillResult.Ok($"文件夹已创建: {path}");
    }

    private SkillResult GetFileInfo(Dictionary<string, object> args)
    {
        var path = GetStr(args, "filePath");
        if (!File.Exists(path)) return new SkillResult { Success = false, Error = $"文件不存在: {path}" };
        var fi = new FileInfo(path);
        return SkillResult.Ok($"文件: {fi.Name}\n大小: {Math.Round(fi.Length / 1024.0, 2)} KB\n创建: {fi.CreationTime:yyyy-MM-dd HH:mm}\n修改: {fi.LastWriteTime:yyyy-MM-dd HH:mm}\n路径: {fi.FullName}");
    }

    private SkillResult OpenFolder(Dictionary<string, object> args)
    {
        var path = GetStr(args, "folderPath");
        System.Diagnostics.Process.Start("explorer.exe", path);
        return SkillResult.Ok($"已打开文件夹: {path}");
    }

    private static Dictionary<string, object> P(string[] req, string[] opt)
    {
        var p = new Dictionary<string, object>();
        foreach (var r in req) p[r] = new { type = "string", description = $"{r}（必需）" };
        foreach (var o in opt) p[o] = new { type = "string", description = $"{o}（可选）" };
        return new Dictionary<string, object> { { "type", "object" }, { "properties", p } };
    }
    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
    private static bool GetBool(Dictionary<string, object> a, string k) => a.ContainsKey(k) && bool.TryParse(a[k]?.ToString(), out var v) && v;
}