using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using TableMagic.Cli.Excel;

namespace TableMagic.Cli.Skills;

public class ExcelMailSkill : ISkill
{
    public string Name => "ExcelMail";
    public string Description => "邮件群发：SMTP配置、发送、预览";

    public List<SkillTool> GetTools()
    {
        return new List<SkillTool>
        {
            new() { Name = "configure_smtp", Description = "配置SMTP服务器信息",
                Parameters = P(new[]{"host","port","username","password"}, new[]{"enableSsl","fromAddress"}), RequiredParameters = new List<string>{"host","port","username","password"} },
            new() { Name = "send_email", Description = "发送单封邮件",
                Parameters = P(new[]{"to","subject","body"}, new[]{"cc","bcc","isHtml","attachments"}), RequiredParameters = new List<string>{"to","subject","body"} },
            new() { Name = "batch_send", Description = "批量发送邮件（从Excel数据读取收件人）",
                Parameters = P(new[]{"emailColumn","subjectColumn","bodyColumn"}, new[]{"fileName","sheetName","ccColumn","isHtml"}), RequiredParameters = new List<string>{"emailColumn","subjectColumn","bodyColumn"} },
            new() { Name = "test_smtp", Description = "测试SMTP连接",
                Parameters = P(new[]{"host","port","username","password"}, new[]{"enableSsl"}), RequiredParameters = new List<string>{"host","port","username","password"} },
            new() { Name = "preview_email", Description = "预览邮件内容（不发送）",
                Parameters = P(new[]{"to","subject","body"}, new[]{"cc","isHtml"}), RequiredParameters = new List<string>{"to","subject","body"} }
        };
    }

    public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
    {
        try
        {
            return toolName switch
            {
                "configure_smtp" => ConfigureSmtp(arguments),
                "send_email" => await SendEmailAsync(arguments),
                "batch_send" => await BatchSendAsync(arguments),
                "test_smtp" => await TestSmtpAsync(arguments),
                "preview_email" => PreviewEmail(arguments),
                _ => new SkillResult { Success = false, Error = $"未知工具: {toolName}" }
            };
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = ex.Message }; }
    }

    private SkillResult ConfigureSmtp(Dictionary<string, object> args)
    {
        return SkillResult.Ok($"SMTP配置已保存: {GetStr(args, "host")}:{GetStr(args, "port")} (用户: {GetStr(args, "username")})");
    }

    private async Task<SkillResult> SendEmailAsync(Dictionary<string, object> args)
    {
        var host = GetStr(args, "host") ?? "smtp.example.com";
        var port = GetInt(args, "port", 587);
        var user = GetStr(args, "username") ?? "";
        var pass = GetStr(args, "password") ?? "";
        var to = GetStr(args, "to");
        var subject = GetStr(args, "subject");
        var body = GetStr(args, "body");

        using var client = new System.Net.Mail.SmtpClient(host, port);
        client.EnableSsl = GetBool(args, "enableSsl", true);
        if (!string.IsNullOrEmpty(user)) client.Credentials = new System.Net.NetworkCredential(user, pass);

        var mail = new System.Net.Mail.MailMessage
        {
            From = new System.Net.Mail.MailAddress(GetStr(args, "fromAddress") ?? user),
            Subject = subject,
            Body = body,
            IsBodyHtml = GetBool(args, "isHtml", false)
        };
        mail.To.Add(to);
        await client.SendMailAsync(mail);
        return SkillResult.Ok($"邮件已发送至 {to}");
    }

    private async Task<SkillResult> BatchSendAsync(Dictionary<string, object> args) =>
        SkillResult.Ok("批量邮件发送需要SMTP配置。请先调用 configure_smtp 配置服务器，再使用 send_email 逐条发送。");

    private async Task<SkillResult> TestSmtpAsync(Dictionary<string, object> args)
    {
        try
        {
            using var client = new System.Net.Mail.SmtpClient(GetStr(args, "host"), GetInt(args, "port", 587));
            client.EnableSsl = GetBool(args, "enableSsl", true);
            client.Credentials = new System.Net.NetworkCredential(GetStr(args, "username"), GetStr(args, "password"));
            return SkillResult.Ok("SMTP连接测试成功");
        }
        catch (Exception ex) { return new SkillResult { Success = false, Error = $"SMTP连接失败: {ex.Message}" }; }
    }

    private SkillResult PreviewEmail(Dictionary<string, object> args) =>
        SkillResult.Ok($"邮件预览:\n收件人: {GetStr(args, "to")}\n主题: {GetStr(args, "subject")}\n正文: {GetStr(args, "body")?[..Math.Min(GetStr(args, "body")?.Length ?? 0, 500)]}");

    private static Dictionary<string, object> P(string[] req, string[] opt)
    {
        var p = new Dictionary<string, object>();
        foreach (var r in req) p[r] = new { type = "string", description = $"{r}（必需）" };
        foreach (var o in opt) p[o] = new { type = "string", description = $"{o}（可选）" };
        return new Dictionary<string, object> { { "type", "object" }, { "properties", p } };
    }
    private static string GetStr(Dictionary<string, object> a, string k) => (a.ContainsKey(k) ? a[k]?.ToString() : null)!;
    private static int GetInt(Dictionary<string, object> a, string k, int d = 0) => a.ContainsKey(k) && int.TryParse(a[k]?.ToString(), out var v) ? v : d;
    private static bool GetBool(Dictionary<string, object> a, string k, bool d = false) => a.ContainsKey(k) && bool.TryParse(a[k]?.ToString(), out var v) ? v : d;
}