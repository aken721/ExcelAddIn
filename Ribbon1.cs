using Microsoft.Office.Tools.Ribbon;
using System.Threading.Tasks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using FlaUI.Core;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Conditions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using Microsoft.Office.Interop.Excel;
using System.Windows.Controls.Primitives;
using System.Reflection;
using System.IO;
using System.Windows.Shapes;
using static System.Net.WebRequestMethods;

namespace ExcelAddIn
{


    public partial class Ribbon1
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            select_f_or_d.Checked = false;
            select_f_or_d.Label = "改文件名";
            select_f_or_d.ShowLabel = false;
            switch_FD_label.Label = "文件名";
        }


        //表操作按钮
        private void excel_extend_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 form1 = new Form1();
            form1.ShowDialog();
        }

        //邮件群发按钮
        private void send_mail_Click(object sender, RibbonControlEventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }


        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [STAThread]
        private static void send_wx(string target, string sendMsg)
        {
            // 设置控制台程序窗口位置和大小
            IntPtr ptr = GetConsoleWindow();
            MoveWindow(ptr, 0, 0, 200, 100, true);


            Process[] processes = Process.GetProcessesByName("WeChat");
            if (processes.Count() != 1)
            {
               MessageBox.Show("微信未启动或启动多个微信");
            }
            else
            {
                //1.附加到微信进程
                using (var app = FlaUI.Core.Application.Attach(processes.First().Id))
                {
                    using (var automation = new UIA3Automation())
                    {

                        //2.获取主界面
                        var mainWindow = app.GetMainWindow(automation);
                        MessageBox.Show("获取主界面");
                        //3.切换到通讯录
                        var elements = mainWindow.FindAll(FlaUI.Core.Definitions.TreeScope.Subtree, TrueCondition.Default);
                        var addressBook = mainWindow.FindFirstDescendant(cf => cf.ByName("通讯录"));
                        addressBook.DrawHighlight(System.Drawing.Color.Red);
                        MessageBox.Show("点击通讯录");
                        addressBook.Click();

                        // 4.搜索
                        var searchTextBox = mainWindow.FindFirstDescendant(cf => cf.ByName("搜索")).AsTextBox();
                        searchTextBox.Click();

                        Keyboard.Type(target);
                        Keyboard.Type(VirtualKeyShort.RETURN);
                        MessageBox.Show("搜索目标对象");

                        //5.切换到对话框
                        Thread.Sleep(500);

                        var searchList = mainWindow.FindFirstDescendant(cf => cf.ByName("搜索结果"));
                        if (searchList != null)
                        {
                            // cf.Name.Contains(target) 模糊查询，也可以 cf.Name==target 精确查询
                            var searchItem = searchList.FindAllDescendants().FirstOrDefault(cf => cf.Name.Contains(target) && cf.ControlType == FlaUI.Core.Definitions.ControlType.ListItem);
                            searchItem?.DrawHighlight(System.Drawing.Color.Red);
                            searchItem?.AsListBoxItem().Click();
                        }
                        else
                        {
                            Console.WriteLine("没有搜索到内容");
                            return;
                        }
                        Thread.Sleep(500);


                        //6.输入文本
                        var msgInput = mainWindow.FindFirstDescendant(cf => cf.ByName("输入")).AsTextBox();

                        if (msgInput == null) return;

                        msgInput?.Click();
                        System.Windows.Forms.Clipboard.SetText(sendMsg);
                        Keyboard.TypeSimultaneously(new[] { VirtualKeyShort.CONTROL, VirtualKeyShort.KEY_V });

                        Thread.Sleep(500);

                        //按下回车
                        Keyboard.Press(VirtualKeyShort.ENTER);

                        //点击发送按钮
                        //var sendBtn = mainWindow.FindFirstDescendant(cf => cf.ByName("发送(S)"));
                        //Console.WriteLine(sendBtn.Name);
                        //sendBtn?.DrawHighlight(System.Drawing.Color.Red);
                        //sendBtn?.Click();

                    }
                }
            }
        }

        //微信群发按钮
        private void send_message_Click(object sender, RibbonControlEventArgs e)
        {
            int target_contacts_column = getUsedRangeColumn("昵称");
            int target_message_column = getUsedRangeColumn("信息");
            for(int i = 2;i<=ThisAddIn.app.ActiveSheet.UsedRange.Rows.Count;i++) 
            {
                send_wx(ThisAddIn.app.ActiveSheet.Cells[i,target_contacts_column].Value, ThisAddIn.app.ActiveSheet.Cells[i, target_message_column].Value);
            }
        }

        
        //指定字段名所处的列
        private  int getUsedRangeColumn(string targetColumn)
        {
            for(int n=1;n<=ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count;n++)
            {
                string targetValue= ThisAddIn.app.ActiveSheet.Cells[1,n].Value.ToString();
                if( targetValue== targetColumn) return n;
            }
            return 0;
        }


        //批读文件名和批改文件名选择路径
        string get_directory_path;

        //批读文件名
        private void files_read_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;

            
            ThisAddIn.app.DisplayAlerts = false;
            ThisAddIn.app.ScreenUpdating = false;


            folderBrowserDialog1.Description = "请选择文件所在文件夹";
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                get_directory_path = folderBrowserDialog1.SelectedPath;
            }
            else
            {
                return;
            }

            if (!string.IsNullOrEmpty(get_directory_path))
            {
                string bat_name = get_directory_path + "\\run.bat";
                FileInfo run_file = new FileInfo(bat_name);
                if (run_file.Exists)
                {
                    run_file.Delete();
                }
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == "rename")
                    {
                        sheet.Name = "rename_备份";
                    }
                }
                Excel.Worksheet worksheet = workbook.Worksheets.Add();
                worksheet.Name = "rename";
                worksheet.Activate();
                switch (select_f_or_d.Checked)
                {
                    case false:  
                        worksheet.Cells[1, 1] = "路径";
                        worksheet.Cells[1, 2] = "旧文件名";
                        worksheet.Cells[1, 3] = "新文件名";
                        string[] files = Directory.GetFiles(get_directory_path, "*.*", SearchOption.AllDirectories);
                        if (files.Length > 0)
                        {
                            for (int i = 1; i <= files.Length; i++)
                            {
                                string file_name = System.IO.Path.GetFileName(files[i - 1]);
                                string file_path = System.IO.Path.GetDirectoryName(files[i - 1]);
                                workbook.ActiveSheet.Cells[i + 1, 1] = file_path;
                                workbook.ActiveSheet.Cells[i + 1, 2] = file_name;
                                workbook.ActiveSheet.Cells[i + 1, 3] = file_name;
                            }
                        }
                        break;
                    case true:
                        worksheet.Cells[1, 1] = "文件夹路径";
                        worksheet.Cells[1, 2] = "旧文件夹名";
                        worksheet.Cells[1, 3] = "新文件夹名";
                        string[] directorys = Directory.GetDirectories(get_directory_path, "*", SearchOption.AllDirectories);
                        if (directorys.Length > 0)
                        {
                            for (int i = 1; i <= directorys.Length; i++)
                            {
                                string[] directory = directorys[i - 1].Split('\\');
                                string directory_name = directory[directory.Length - 1];
                                Array.Resize(ref directory, directory.Length - 1);
                                string directory_path = string.Join("\\", directory);
                                workbook.ActiveSheet.Cells[i + 1, 1] = directory_path;
                                workbook.ActiveSheet.Cells[i + 1, 2] = directory_name;
                                workbook.ActiveSheet.Cells[i + 1, 3] = directory_name;
                            }
                        }
                        break;
                }
            }
            else
            {
                MessageBox.Show("未选择文件夹");
            }
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }

        //批量重命名
        private void file_rename_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;
            if (!string.IsNullOrEmpty(get_directory_path))
            {
                string bat_name = get_directory_path + "\\run.bat";
                FileInfo run_file = new FileInfo(bat_name);
                if (run_file.Exists)
                {
                    run_file.Delete();
                }

                using (StreamWriter bat_file = new StreamWriter(bat_name,false,Encoding.GetEncoding("gb2312")))
                {
                    for (int i = 2; i <= workbook.ActiveSheet.UsedRange.Rows.Count; i++)
                    {
                        string cell1 = workbook.ActiveSheet.Cells[i, 1].Value;
                        string cell2 = workbook.ActiveSheet.Cells[i, 2].Value;
                        string cell3 = workbook.ActiveSheet.Cells[i, 3].Value;
                        string full_path;
                        string old_name;
                        string new_name;
                        string[] cell_arr = cell1.Split('\\');
                        for (int n = 0; n < cell_arr.Length; n++)
                        {
                            if (cell_arr[n].Contains(" "))
                            {
                                cell_arr[n] = "\"" + cell_arr[n] + "\"";
                            }
                        }
                        full_path = string.Join("\\", cell_arr);
                        if (cell2.Contains(" "))
                        {
                            old_name = "\"" + cell2 + "\"";
                        }
                        else
                        {
                            old_name = cell2;
                        }
                        if (cell3.Contains(" "))
                        {
                            new_name = "\"" + cell3 + "\"";
                        }
                        else
                        {
                            new_name = cell3;
                        }
                        string code = "ren " + full_path + "\\" + old_name + " " + new_name;
                        bat_file.WriteLine(code, Encoding.GetEncoding("gb2312"));
                    }
                    bat_file.Close();
                }

                workbook.Worksheets["rename"].Delete();
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == "rename_备份")
                    {
                        workbook.Worksheets["rename_备份"].Name = "rename";
                    }
                }
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;

                //调用批处理文件改文件名，并且不显示cmd窗口
                ProcessStartInfo startInfo = new ProcessStartInfo()
                {
                    FileName =bat_name,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                Process proc = Process.Start(startInfo);
                proc.WaitForExit();
                proc.Close();
                System.IO.File.Delete(bat_name);
                MessageBox.Show("文件名修改完毕");
                Process.Start(get_directory_path); 
            }
            else
            {
                MessageBox.Show("没有选择文件夹，请先使用批读文件名功能后再使用该功能");
            }
           
        }

        //文件目录选项
        private void select_f_or_d_Click(object sender, RibbonControlEventArgs e)
        {
            if(select_f_or_d.Checked==true)
            {
                select_f_or_d.Image = ExcelAddIn.Properties.Resources.Radio_Button_on;
                select_f_or_d.Label = "改文件夹名";
                select_f_or_d.ShowLabel = false;
                switch_FD_label.Label = "目录名";
            }
            else
            {
                select_f_or_d.Image = ExcelAddIn.Properties.Resources.Radio_Button_off;
                select_f_or_d.Label = "改文件名";
                select_f_or_d.ShowLabel = false;
                switch_FD_label.Label = "文件名";
            }
        }

        private void rename_mp3_Click(object sender, RibbonControlEventArgs e)
        {
            Form3 form3 = new Form3();
            form3.ShowDialog();
        }
    }
}
