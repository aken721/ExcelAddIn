using System;
using System.Collections.Generic;
using System.Drawing;
using System.Net.Mail;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }
        private Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;
        private string pictureType = "hide";
        private string myAttachment_dir;
        private List<string> myMailsto = new List<string>();
        private List<string> myAttachment = new List<string>();
        private List<string> errRecord = new List<string>();
        private Dictionary<string, List<string>> address_attachment = new Dictionary<string, List<string>>();

        private void Form2_Load(object sender, EventArgs e)
        {
            this.MaximizeBox = false;
            send_progress_label.Visible = false;
            send_progressBar.Visible = false;
            attachment_no_radioButton.Select();
            attachment_textBox.Visible = false;
            attachment_checkBox.Visible = false;
            mailpassword_textBox.UseSystemPasswordChar = true;
            pictureType = "hide";
            for (int i = 1; i <= ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count; i++)
            {
                mailto_comboBox.Items.Add(ThisAddIn.app.ActiveSheet.Cells[1, i].Text);
            }
            foreach (Control control in this.Controls)
            {
                if (control is Label)
                {
                    control.TabStop = false;
                }
            }
        }

        //发送附件radioButton选中或未选中
        private void attachment_yes_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (attachment_yes_radioButton.Checked == true)
            {
                attachment_textBox.Visible = true;
                attachment_checkBox.Visible = true;
            }
            else
            {
                attachment_textBox.Visible = false;
                attachment_checkBox.Visible = false;
            }
        }

        //附件textBox被单击时
        private void attachment_textBox_Click(object sender, EventArgs e)
        {
            if (attachment_textBox.Text == "请手工输入附件所在目录的完整路径，或双击选择目录" || attachment_textBox.Text == "请手工输入文件的完整路径，或双击选择文件")
            {
                attachment_textBox.Text = "";
                attachment_textBox.ForeColor = Color.Black;
            }
        }


        //进度条更新函数
        private void UpdateProgressBar(ProgressBar progressBar, int currentSheet, int totalSheets, Label progressBar_result_label, string progressBar_result)
        {
            // 计算进度百分比
            int progressPercentage = (int)((double)currentSheet / totalSheets * 100);
            // 更新进度条控件
            progressBar.Value = progressPercentage;
            progressBar.Update();
            // 显示百分比数字
            progressBar_result_label.Text = progressBar_result + progressPercentage.ToString() + "%";
        }

        //发送按钮
        private void send_button_Click(object sender, EventArgs e)
        {
            if (mailfrom_textBox.Text != "" && mailfrom_comboBox.Text != "" && smtp_textBox.Text != "" && port_textBox.Text != "" && mailpassword_textBox.Text != "")
            {
                string myMail = mailfrom_textBox.Text + "@" + mailfrom_comboBox.Text;
                string myPassword = mailpassword_textBox.Text;
                string mySmtp = smtp_textBox.Text;
                string myPort = port_textBox.Text;
                string mySubject = subject_textBox.Text;
                string myAttachpath = attachment_textBox.Text;

                string myBody = body_richTextBox.Text;
                int address_column = 0;
                int attachment_column = 0;

                if (attachment_textBox.Text == "请手工输入附件所在目录的完整路径，或双击选择目录" || attachment_textBox.Text == "请手工输入文件的完整路径，或双击选择文件")
                {
                    attachment_textBox.Text = "";
                    attachment_textBox.ForeColor = Color.Black;
                }

                //读取收件人地址并写入myMailsto列表
                if (mailto_textBox.Text != "" || mailto_comboBox.Text != "")
                {
                    myMailsto.Clear();
                    //读取手工输入收件人地址
                    if (mailto_textBox.Text != "")
                    {
                        foreach (string mail in mailto_textBox.Text.Split(",".ToCharArray()))
                        {
                            myMailsto.Add(mail);
                        }
                    }

                    //读取excel表中收件人地址列内容
                    if (mailto_comboBox.Text != "")
                    {
                        foreach (Excel.Range fields_range in ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells[1, 1], ThisAddIn.app.ActiveSheet.Cells[1, ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count]])
                        {
                            if (fields_range.Value == mailto_comboBox.Text)
                            {
                                address_column = fields_range.Column;
                                break;
                            }
                        }
                        foreach (Excel.Range records_range in ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells[2, address_column], ThisAddIn.app.ActiveSheet.Cells[ThisAddIn.app.ActiveSheet.UsedRange.Rows.Count, address_column]])
                        {
                            if (!string.IsNullOrEmpty(records_range.Value) && !myMailsto.Contains(records_range.Value))
                            {
                                myMailsto.Add(records_range.Value);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("收件人邮箱地址不能为空，请核对后再次运行");
                    return;
                }

                //读取附件
                if (attachment_yes_radioButton.Checked)
                {
                    if (attachment_checkBox.Checked)
                    {
                        foreach (Excel.Range fields_range in ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells[1, 1], ThisAddIn.app.ActiveSheet.Cells[1, ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count]])
                        {
                            if (fields_range.Value == "附件")
                            {
                                attachment_column = fields_range.Column;
                                break;
                            }
                        }
                        address_attachment.Clear();
                        for (int i = 2; i <= ThisAddIn.app.ActiveSheet.UsedRange.Rows.Count; i++)
                        {
                            if (!address_attachment.ContainsKey(ThisAddIn.app.ActiveSheet.Cells[i, address_column].Value))
                            {
                                string dic_key = ThisAddIn.app.ActiveSheet.Cells[i, address_column].Value;
                                List<string> dic_value = new List<string>();
                                string attach_path = ThisAddIn.app.ActiveSheet.Cells[i, attachment_column].Value;
                                if (attach_path.Contains(";"))
                                {
                                    string[] attachmentPathValues = attach_path.Split(';');
                                    foreach (string attachmentPathValue in attachmentPathValues)
                                    {
                                        dic_value.Add(myAttachpath + "\\" + attachmentPathValue);
                                    }
                                }
                                else if (attach_path.Contains("；"))
                                {
                                    string[] attachmentPathValues = attach_path.Split('；');
                                    foreach (string attachmentPathValue in attachmentPathValues)
                                    {
                                        dic_value.Add(myAttachpath + "\\" + attachmentPathValue);
                                    }
                                }
                                else
                                {
                                    dic_value.Add(myAttachpath + "\\" + attach_path);
                                }
                                address_attachment.Add(dic_key, dic_value);
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                    else
                    {
                        address_attachment.Clear();
                        foreach (string address_key in myMailsto)
                        {
                            MessageBox.Show(address_key);
                            address_attachment.Add(address_key, myAttachment);
                        }
                    }
                }
                else
                {
                    address_attachment.Clear();
                    foreach (string address_key in myMailsto)
                    {
                        MessageBox.Show(address_key);
                        address_attachment.Add(address_key, myAttachment);
                    }
                }
                int current_maito = 1;
                int total_mailto = myMailsto.Count;

                //遍历收件人地址，调用发邮件函数发送邮件
                foreach (string myMailto in myMailsto)
                {
                    //更新进度条
                    string result_text = "正在发送" + current_maito.ToString() + "个，共" + total_mailto.ToString() + "个，已完成";
                    UpdateProgressBar(send_progressBar, current_maito, total_mailto, send_progress_label, result_text);
                    send_progress_label.Visible = true;
                    send_progressBar.Visible = true;
                    string result;
                    if (ssl_checkBox.Checked)
                    {
                        result = SendMail(myMailto, myMail, myPassword, mySmtp, myPort, mySubject, myBody, address_attachment[myMailto], true);
                    }
                    else
                    {
                        result = SendMail(myMailto, myMail, myPassword, mySmtp, myPort, mySubject, myBody, address_attachment[myMailto]);
                    }
                    current_maito++;
                    if (result != "finished")
                    {
                        errRecord.Add(myMailto + ":" + result);
                    }
                }
                if (errRecord.Count == 0)
                {
                    MessageBox.Show("共" + total_mailto.ToString() + "封邮件，全部发送成功");
                }
                else
                {
                    MessageBox.Show("共" + total_mailto.ToString() + "封邮件，其中有" + errRecord.Count.ToString() + "未发送成功,原因是：" + string.Join("\n", errRecord));
                }
            }
            else
            {
                MessageBox.Show("发件人邮箱和密码不能为空");
            }
        }


        //邮件发送
        private string SendMail(string mailTo, string mailFrom, string password, string mailSmtp, string smtPort, string mailSubject, string mailBody, List<string> mailAttachPaths = null, bool ssl = false)
        {
            try
            {
                //设置SMTP服务器和发送者信息
                SmtpClient smtpServer = new SmtpClient(mailSmtp);
                MailMessage mail = new MailMessage()
                {
                    From = new MailAddress(mailFrom)
                };
                smtpServer.Credentials = new System.Net.NetworkCredential(mailFrom, password);
                if (ssl)
                {
                    smtpServer.EnableSsl = true;
                }
                else { smtpServer.EnableSsl = false; }

                //设置收件人和邮件内容
                mail.To.Add(mailTo);
                mail.Subject = mailSubject;
                mail.Body = mailBody;
                mail.IsBodyHtml = true;

                //添加附件
                if (mailAttachPaths != null)
                {
                    foreach (string mailAttachPath in mailAttachPaths)
                    {
                        Attachment attachment = new Attachment(mailAttachPath);
                        mail.Attachments.Add(attachment);
                    }
                }

                //发送邮件
                smtpServer.Send(mail);

                //清空收件人和附件列表
                mail.To.Clear();
                mail.Attachments.Clear();

                //关闭SMTP连接
                smtpServer.Dispose();
                return "finished";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }



        //清空按钮
        private void clear_button_Click(object sender, EventArgs e)
        {
            subject_textBox.Text = "";
            body_richTextBox.Text = "";
            attachment_textBox.Text = "";
            mailto_textBox.Text = "";
            mailto_comboBox.Text = "";
            mailfrom_textBox.Text = "";
            mailfrom_comboBox.Text = "";
            mailpassword_textBox.Text = "";
            smtp_textBox.Text = "";
            port_textBox.Text = "";
        }

        //退出按钮
        private void quit_button_Click(object sender, EventArgs e)
        {
            Close();
        }

        //双击文本框选择附件名称和路径，并写入全局列表中
        private void attachment_textBox_DoubleClick(object sender, EventArgs e)
        {
            ThisAddIn.app.DisplayAlerts = false;
            ThisAddIn.app.ScreenUpdating = false;
            if (attachment_checkBox.Checked == true)
            {
                attachment_folderBrowserDialog.Description = "请选择附件所在文件夹";
                attachment_folderBrowserDialog.ShowDialog();
                if (attachment_folderBrowserDialog.SelectedPath.Length > 0)
                {
                    attachment_textBox.Text = attachment_folderBrowserDialog.SelectedPath;
                    attachment_textBox.ForeColor = Color.Black;
                    myAttachment_dir = attachment_textBox.Text;
                }
                else
                {
                    return;
                }
            }
            else
            {
                attachment_openFileDialog.Title = "选择一个或多个文件";
                attachment_openFileDialog.Filter = "所有文件|*.*";
                attachment_openFileDialog.Multiselect = true;
                attachment_openFileDialog.ShowDialog();
                if (attachment_openFileDialog.FileNames.Length > 0)
                {
                    attachment_textBox.Text = string.Join(";", attachment_openFileDialog.SafeFileNames);
                    attachment_textBox.ForeColor = Color.Black;
                    myAttachment.Clear();
                    foreach (string file_name in attachment_openFileDialog.FileNames)
                    {
                        myAttachment.Add(file_name);
                    }
                }
                else
                {
                    return;
                }
            }
        }

        //邮箱smtp服务器和端口号自动匹配
        private void mailfrom_comboBox_TextChanged(object sender, EventArgs e)
        {
            switch (mailfrom_comboBox.Text)
            {
                case "163.com":
                    smtp_textBox.Text = "smtp.163.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "qq.com":
                    smtp_textBox.Text = "smtp.qq.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "sina.com":
                    smtp_textBox.Text = "smtp.sina.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "sina.cn":
                    smtp_textBox.Text = "smtp.sina.cn";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "126.com":
                    smtp_textBox.Text = "smtp.126.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "sohu.com":
                    smtp_textBox.Text = "smtp.sohu.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "yeah.net":
                    smtp_textBox.Text = "smtp.yeah.net";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "139.com":
                    smtp_textBox.Text = "smtp.139.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "189.cn":
                    smtp_textBox.Text = "smtp.189.cn";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "gmail.com":
                    smtp_textBox.Text = "smtp.gmail.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "outlook.com":
                    smtp_textBox.Text = "smtp-mail.outlook.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "hotmail.com":
                    smtp_textBox.Text = "smtp-mail.outlook.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "aliyun.com":
                    smtp_textBox.Text = "smtp.aliyun.com";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                case "wo.cn":
                    smtp_textBox.Text = "smtp.wo.cn";
                    smtp_textBox.ReadOnly = true;
                    if (ssl_checkBox.Checked == true)
                    {
                        port_textBox.Text = "465";
                        port_textBox.ReadOnly = false;
                    }
                    else
                    {
                        port_textBox.Text = "25";
                        port_textBox.ReadOnly = true;
                    }
                    break;
                default:
                    smtp_textBox.Text = "";
                    smtp_textBox.ReadOnly = false;
                    port_textBox.Text = "";
                    port_textBox.ReadOnly = false;
                    break;
            }
        }

        //ssl加密选项被选中
        private void ssl_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (ssl_checkBox.Checked == true)
            {
                if (port_textBox.Text != "")
                {
                    port_textBox.Text = "465";
                    port_textBox.ReadOnly = false;
                }
                else
                {
                    port_textBox.Text = "";
                    port_textBox.ReadOnly = false;
                }
            }
            else
            {
                if (smtp_textBox.Text != "")
                {
                    port_textBox.Text = "25";
                    port_textBox.ReadOnly = true;
                }
                else
                {
                    port_textBox.Text = "";
                    port_textBox.ReadOnly = false;
                }
            }
        }


        //发送不同附件checkBox选中时
        private void attachment_checkBox_CheckedChanged(object sender, EventArgs e)
        {
            if (attachment_checkBox.Checked == true)
            {
                attachment_textBox.ForeColor = Color.LightGray;
                toolTip1.SetToolTip(attachment_textBox, "请手工输入附件所在目录的完整路径，或双击选择目录");
                attachment_textBox.Text = "请手工输入附件所在目录的完整路径，或双击选择目录";
                mailto_textBox.Enabled = false;
            }
            else
            {
                attachment_textBox.ForeColor = Color.LightGray;
                toolTip1.SetToolTip(attachment_textBox, "请手工输入文件的完整路径，或双击选择文件");
                attachment_textBox.Text = "请手工输入文件的完整路径，或双击选择文件";
                mailto_textBox.Enabled = true;
            }
        }

        //不同附件选项改变时
        private void attachment_checkBox_CheckStateChanged(object sender, EventArgs e)
        {
            if (attachment_checkBox.Checked == true)
            {
                List<string> messages = new List<string>() {
                        "1.发送不同附件只能使用表结构维护,需要在当前excel表中存在‘附件’列。",
                        "2.请注意所有附件均应放入同一文件夹，并可使用功能包中的‘批读文件名’读入文件名。",
                        "3.将批量读入文件名复制入‘附件’列，并做好与电子邮箱所在列的对应关系。",
                        "4.如某一发件人需接收多个附件，请将文件名维护在同一单元格内，并用;隔开。",
                        "请务必按以上说明操作，否则发送附件将会出错!"
                    };
                string message = string.Join(Environment.NewLine, messages);
                string caption = "重要提示";
                DialogResult dr = MessageBox.Show(message, caption, MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (dr != DialogResult.OK)
                {
                    attachment_checkBox.Checked = false;
                }
            }
        }


        //密码可见图片单击时
        private void mailpassword_pictureBox_Click(object sender, EventArgs e)
        {
            switch (pictureType)
            {
                case "hide":
                    mailpassword_pictureBox.Image = ExcelAddIn.Properties.Resources.eye_open;
                    mailpassword_textBox.UseSystemPasswordChar = false;
                    pictureType = "open";
                    break;
                case "open":
                    mailpassword_pictureBox.Image = ExcelAddIn.Properties.Resources.eye_hide;
                    mailpassword_textBox.UseSystemPasswordChar = true;
                    pictureType = "hide";
                    break;
            }
        }

        //窗体最小化时
        private void Form2_VisibleChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                ThisAddIn.app.Visible = true;
                ThisAddIn.app.WindowState = Excel.XlWindowState.xlMaximized;
            }
        }


        private void mailfrom_comboBox_GotFocus(object sender, EventArgs e)
        {
            mailfrom_comboBox.Text = mailfrom_comboBox.Items[0].ToString();
        }
    }
}
