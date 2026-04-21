using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace TableMagic.Skills
{
    public class ExcelMailSkill : ISkill
    {
        public string Name => "ExcelMail";
        public string Description => "邮件技能，支持邮件群发、邮件模板和附件发送";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "send_email",
                    Description = "发送单封邮件。当用户要求发送邮件时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "to", new { type = "string", description = "收件人邮箱地址" } },
                                { "subject", new { type = "string", description = "邮件主题" } },
                                { "body", new { type = "string", description = "邮件正文" } },
                                { "isHtml", new { type = "boolean", description = "是否HTML格式（默认false）" } },
                                { "attachments", new { type = "string", description = "附件路径列表（JSON数组格式，可选）" } },
                                { "cc", new { type = "string", description = "抄送人邮箱（可选）" } },
                                { "bcc", new { type = "string", description = "密送人邮箱（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "to", "subject", "body" }
                },
                new SkillTool
                {
                    Name = "batch_send_emails",
                    Description = "根据Excel数据批量发送邮件。Excel中需包含收件人、主题、正文列。当用户要求群发邮件、批量发送邮件时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "emailColumn", new { type = "string", description = "收件人邮箱列名" } },
                                { "subjectColumn", new { type = "string", description = "邮件主题列名" } },
                                { "bodyColumn", new { type = "string", description = "邮件正文列名" } },
                                { "attachmentColumn", new { type = "string", description = "附件路径列名（可选）" } },
                                { "subjectTemplate", new { type = "string", description = "主题模板（使用{列名}替换，可选）" } },
                                { "bodyTemplate", new { type = "string", description = "正文模板（使用{列名}替换，可选）" } },
                                { "sheetName", new { type = "string", description = "工作表名称（可选）" } },
                                { "startRow", new { type = "integer", description = "起始行（默认2）" } },
                                { "endRow", new { type = "integer", description = "结束行（可选，默认到最后一行）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "emailColumn" }
                },
                new SkillTool
                {
                    Name = "configure_smtp",
                    Description = "配置SMTP邮件服务器设置。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "host", new { type = "string", description = "SMTP服务器地址" } },
                                { "port", new { type = "integer", description = "SMTP端口（默认587）" } },
                                { "username", new { type = "string", description = "发件人邮箱" } },
                                { "password", new { type = "string", description = "邮箱密码或授权码" } },
                                { "enableSsl", new { type = "boolean", description = "是否启用SSL（默认true）" } },
                                { "displayName", new { type = "string", description = "发件人显示名称（可选）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "host", "username", "password" }
                },
                new SkillTool
                {
                    Name = "test_smtp_connection",
                    Description = "测试SMTP服务器连接是否正常。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "preview_email",
                    Description = "预览邮件内容（不发送）。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "emailColumn", new { type = "string", description = "收件人邮箱列名" } },
                                { "subjectTemplate", new { type = "string", description = "主题模板" } },
                                { "bodyTemplate", new { type = "string", description = "正文模板" } },
                                { "previewRow", new { type = "integer", description = "预览行号（默认2）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "emailColumn" }
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "send_email":
                        return await SendEmailAsync(arguments);
                    case "batch_send_emails":
                        return await BatchSendEmailsAsync(arguments);
                    case "configure_smtp":
                        return ConfigureSmtp(arguments);
                    case "test_smtp_connection":
                        return await TestSmtpConnectionAsync();
                    case "preview_email":
                        return PreviewEmail(arguments);
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelMailSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private static SmtpConfig _smtpConfig;

        private async Task<SkillResult> SendEmailAsync(Dictionary<string, object> arguments)
        {
            return await Task.Run(() =>
            {
                if (_smtpConfig == null)
                    return new SkillResult { Success = false, Error = "请先配置SMTP服务器" };

                var to = arguments["to"].ToString();
                var subject = arguments["subject"].ToString();
                var body = arguments["body"].ToString();
                var isHtml = arguments.ContainsKey("isHtml") && Convert.ToBoolean(arguments["isHtml"]);
                var cc = arguments.ContainsKey("cc") ? arguments["cc"].ToString() : null;
                var bcc = arguments.ContainsKey("bcc") ? arguments["bcc"].ToString() : null;
                var attachments = arguments.ContainsKey("attachments") 
                    ? Newtonsoft.Json.JsonConvert.DeserializeObject<List<string>>(arguments["attachments"].ToString()) 
                    : null;

                try
                {
                    using (var message = new MailMessage())
                    {
                        message.From = new MailAddress(_smtpConfig.Username, _smtpConfig.DisplayName);
                        message.To.Add(to);

                        if (!string.IsNullOrEmpty(cc))
                            message.CC.Add(cc);

                        if (!string.IsNullOrEmpty(bcc))
                            message.Bcc.Add(bcc);

                        message.Subject = subject;
                        message.Body = body;
                        message.IsBodyHtml = isHtml;

                        if (attachments != null)
                        {
                            foreach (var attachment in attachments)
                            {
                                if (File.Exists(attachment))
                                    message.Attachments.Add(new Attachment(attachment));
                            }
                        }

                        using (var client = new SmtpClient(_smtpConfig.Host, _smtpConfig.Port))
                        {
                            client.EnableSsl = _smtpConfig.EnableSsl;
                            client.Credentials = new NetworkCredential(_smtpConfig.Username, _smtpConfig.Password);
                            client.Send(message);
                        }
                    }

                    return new SkillResult { Success = true, Content = $"邮件发送成功：{to}" };
                }
                catch (Exception ex)
                {
                    return new SkillResult { Success = false, Error = $"邮件发送失败：{ex.Message}" };
                }
            });
        }

        private async Task<SkillResult> BatchSendEmailsAsync(Dictionary<string, object> arguments)
        {
            if (_smtpConfig == null)
                return new SkillResult { Success = false, Error = "请先配置SMTP服务器" };

            var emailColumn = arguments["emailColumn"].ToString();
            var subjectColumn = arguments.ContainsKey("subjectColumn") ? arguments["subjectColumn"].ToString() : null;
            var bodyColumn = arguments.ContainsKey("bodyColumn") ? arguments["bodyColumn"].ToString() : null;
            var attachmentColumn = arguments.ContainsKey("attachmentColumn") ? arguments["attachmentColumn"].ToString() : null;
            var subjectTemplate = arguments.ContainsKey("subjectTemplate") ? arguments["subjectTemplate"].ToString() : null;
            var bodyTemplate = arguments.ContainsKey("bodyTemplate") ? arguments["bodyTemplate"].ToString() : null;
            var sheetName = arguments.ContainsKey("sheetName") ? arguments["sheetName"].ToString() : null;
            var startRow = arguments.ContainsKey("startRow") ? Convert.ToInt32(arguments["startRow"]) : 2;
            var endRow = arguments.ContainsKey("endRow") ? Convert.ToInt32(arguments["endRow"]) : 0;

            return await Task.Run(() =>
            {
                var workbook = ThisAddIn.app.ActiveWorkbook;
                var sheet = string.IsNullOrEmpty(sheetName) ? workbook.ActiveSheet : workbook.Worksheets[sheetName];

                var usedRange = sheet.UsedRange;
                int lastRow = endRow > 0 ? endRow : usedRange.Rows.Count;
                int lastCol = usedRange.Columns.Count;

                var columnMap = new Dictionary<string, int>();
                for (int c = 1; c <= lastCol; c++)
                {
                    var colName = sheet.Cells[1, c].Text?.ToString();
                    if (!string.IsNullOrEmpty(colName))
                        columnMap[colName] = c;
                }

                if (!columnMap.ContainsKey(emailColumn))
                    return new SkillResult { Success = false, Error = $"未找到邮箱列: {emailColumn}" };

                int emailColIdx = columnMap[emailColumn];
                int subjectColIdx = subjectColumn != null && columnMap.ContainsKey(subjectColumn) ? columnMap[subjectColumn] : 0;
                int bodyColIdx = bodyColumn != null && columnMap.ContainsKey(bodyColumn) ? columnMap[bodyColumn] : 0;
                int attachmentColIdx = attachmentColumn != null && columnMap.ContainsKey(attachmentColumn) ? columnMap[attachmentColumn] : 0;

                int successCount = 0;
                int failCount = 0;
                var errors = new List<string>();

                for (int r = startRow; r <= lastRow; r++)
                {
                    var email = sheet.Cells[r, emailColIdx].Text?.ToString();
                    if (string.IsNullOrEmpty(email)) continue;

                    string subject = subjectTemplate ?? "";
                    string body = bodyTemplate ?? "";

                    if (!string.IsNullOrEmpty(subjectTemplate))
                    {
                        subject = ReplaceTemplate(subjectTemplate, sheet, r, columnMap);
                    }
                    else if (subjectColIdx > 0)
                    {
                        subject = sheet.Cells[r, subjectColIdx].Text?.ToString() ?? "";
                    }

                    if (!string.IsNullOrEmpty(bodyTemplate))
                    {
                        body = ReplaceTemplate(bodyTemplate, sheet, r, columnMap);
                    }
                    else if (bodyColIdx > 0)
                    {
                        body = sheet.Cells[r, bodyColIdx].Text?.ToString() ?? "";
                    }

                    var attachmentPath = attachmentColIdx > 0 ? sheet.Cells[r, attachmentColIdx].Text?.ToString() : null;

                    try
                    {
                        using (var message = new MailMessage())
                        {
                            message.From = new MailAddress(_smtpConfig.Username, _smtpConfig.DisplayName);
                            message.To.Add(email);
                            message.Subject = subject;
                            message.Body = body;
                            message.IsBodyHtml = true;

                            if (!string.IsNullOrEmpty(attachmentPath) && File.Exists(attachmentPath))
                                message.Attachments.Add(new Attachment(attachmentPath));

                            using (var client = new SmtpClient(_smtpConfig.Host, _smtpConfig.Port))
                            {
                                client.EnableSsl = _smtpConfig.EnableSsl;
                                client.Credentials = new NetworkCredential(_smtpConfig.Username, _smtpConfig.Password);
                                client.Send(message);
                            }
                        }

                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        errors.Add($"行{r} ({email}): {ex.Message}");
                    }
                }

                var result = $"批量邮件发送完成\n成功: {successCount} 封\n失败: {failCount} 封";
                if (errors.Count > 0)
                    result += $"\n失败详情:\n{string.Join("\n", errors.Take(10))}";

                return new SkillResult { Success = failCount == 0, Content = result };
            });
        }

        private SkillResult ConfigureSmtp(Dictionary<string, object> arguments)
        {
            _smtpConfig = new SmtpConfig
            {
                Host = arguments["host"].ToString(),
                Port = arguments.ContainsKey("port") ? Convert.ToInt32(arguments["port"]) : 587,
                Username = arguments["username"].ToString(),
                Password = arguments["password"].ToString(),
                EnableSsl = !arguments.ContainsKey("enableSsl") || Convert.ToBoolean(arguments["enableSsl"]),
                DisplayName = arguments.ContainsKey("displayName") ? arguments["displayName"].ToString() : ""
            };

            return new SkillResult { Success = true, Content = $"SMTP配置成功\n服务器: {_smtpConfig.Host}:{_smtpConfig.Port}\n发件人: {_smtpConfig.Username}" };
        }

        private async Task<SkillResult> TestSmtpConnectionAsync()
        {
            if (_smtpConfig == null)
                return new SkillResult { Success = false, Error = "请先配置SMTP服务器" };

            return await Task.Run(() =>
            {
                try
                {
                    using (var client = new SmtpClient(_smtpConfig.Host, _smtpConfig.Port))
                    {
                        client.EnableSsl = _smtpConfig.EnableSsl;
                        client.Credentials = new NetworkCredential(_smtpConfig.Username, _smtpConfig.Password);
                        client.Send(new MailMessage(_smtpConfig.Username, _smtpConfig.Username, "测试连接", "这是一封测试邮件"));
                    }

                    return new SkillResult { Success = true, Content = "SMTP连接测试成功" };
                }
                catch (Exception ex)
                {
                    return new SkillResult { Success = false, Error = $"SMTP连接测试失败：{ex.Message}" };
                }
            });
        }

        private SkillResult PreviewEmail(Dictionary<string, object> arguments)
        {
            var emailColumn = arguments["emailColumn"].ToString();
            var subjectTemplate = arguments.ContainsKey("subjectTemplate") ? arguments["subjectTemplate"].ToString() : "";
            var bodyTemplate = arguments.ContainsKey("bodyTemplate") ? arguments["bodyTemplate"].ToString() : "";
            var previewRow = arguments.ContainsKey("previewRow") ? Convert.ToInt32(arguments["previewRow"]) : 2;

            var workbook = ThisAddIn.app.ActiveWorkbook;
            var sheet = workbook.ActiveSheet;

            var usedRange = sheet.UsedRange;
            int lastCol = usedRange.Columns.Count;

            var columnMap = new Dictionary<string, int>();
            for (int c = 1; c <= lastCol; c++)
            {
                var colName = sheet.Cells[1, c].Text?.ToString();
                if (!string.IsNullOrEmpty(colName))
                    columnMap[colName] = c;
            }

            var email = columnMap.ContainsKey(emailColumn) 
                ? sheet.Cells[previewRow, columnMap[emailColumn]].Text?.ToString() 
                : "";

            var subject = ReplaceTemplate(subjectTemplate, sheet, previewRow, columnMap);
            var body = ReplaceTemplate(bodyTemplate, sheet, previewRow, columnMap);

            return new SkillResult 
            { 
                Success = true, 
                Content = $"邮件预览（第{previewRow}行）\n收件人: {email}\n主题: {subject}\n正文:\n{body}" 
            };
        }

        private string ReplaceTemplate(string template, Excel.Worksheet sheet, int row, Dictionary<string, int> columnMap)
        {
            var result = template;
            foreach (var kvp in columnMap)
            {
                var value = sheet.Cells[row, kvp.Value].Text?.ToString() ?? "";
                result = result.Replace($"{{{kvp.Key}}}", value);
            }
            return result;
        }

        private class SmtpConfig
        {
            public string Host { get; set; }
            public int Port { get; set; }
            public string Username { get; set; }
            public string Password { get; set; }
            public bool EnableSsl { get; set; }
            public string DisplayName { get; set; }
        }
    }
}
