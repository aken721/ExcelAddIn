using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class Ribbon1
    {
        private int readFile;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            select_f_or_d.Checked = false;
            select_f_or_d.Label = "改文件名";
            select_f_or_d.ShowLabel = false;
            switch_FD_label.Label = "文件名";
            readFile = 0;
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

        //指定字段名所处的列
        private int getUsedRangeColumn(string targetColumn)
        {
            for (int n = 1; n <= ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count; n++)
            {
                string targetValue = ThisAddIn.app.ActiveSheet.Cells[1, n].Value.ToString();
                if (targetValue == targetColumn) return n;
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
                    if (sheet.Name == "_rename")
                    {
                        sheet.Name = "_rename_备份";
                    }
                }
                Excel.Worksheet worksheet = workbook.Worksheets.Add();
                worksheet.Name = "_rename";
                worksheet.Activate();
                switch (select_f_or_d.Checked)
                {
                    case false:
                        worksheet.Cells[1, 1] = "路径";
                        worksheet.Cells[1, 2] = "旧文件名";
                        worksheet.Cells[1, 3] = "新文件名";
                        List<string> files = new List<string>(Directory.GetFiles(get_directory_path, "*.*", SearchOption.AllDirectories));
                        files.RemoveAll(file => (File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden);
                        if (files.Count > 0)
                        {
                            for (int i = 1; i <= files.Count; i++)
                            {
                                string file_name = Path.GetFileName(files[i - 1]);
                                string file_path = Path.GetDirectoryName(files[i - 1]);
                                workbook.ActiveSheet.Cells[i + 1, 1] = file_path;
                                workbook.ActiveSheet.Cells[i + 1, 2] = file_name;
                                workbook.ActiveSheet.Cells[i + 1, 3] = file_name;
                            }
                        }
                        worksheet.Range["C2"].Select();
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
                            worksheet.Range["C2"].Select();
                        }
                        break;
                }
            }
            else
            {
                MessageBox.Show("未选择文件夹");
            }
            readFile = 1;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }

        //批量重命名
        private void file_rename_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            if (!string.IsNullOrEmpty(get_directory_path) && readFile == 1 && IsSheetExist(workbook, "_rename"))
            {
                //调用file.move或direction.move修改名
                for (int i = 2; i <= workbook.ActiveSheet.UsedRange.Rows.Count; i++)
                {
                    string cell1 = workbook.Worksheets["_rename"].Cells[i, 1].Value;
                    string cell2 = workbook.Worksheets["_rename"].Cells[i, 2].Value;
                    string cell3 = workbook.Worksheets["_rename"].Cells[i, 3].Value;
                    string full_path = cell1;
                    string old_name = Path.Combine(cell1, cell2);
                    string new_name = Path.Combine(cell1, cell3);
                    switch (select_f_or_d.Checked == true)
                    {
                        case false:
                            int exist_file = 0;
                            if (old_name != new_name)
                            {
                                while (File.Exists(new_name))
                                {
                                    exist_file++;
                                    new_name = Path.Combine(cell1, Path.GetFileNameWithoutExtension(new_name) + "(" + exist_file.ToString() + ")" + Path.GetExtension(new_name));
                                }
                                File.Move(old_name, new_name);
                            }
                            break;
                        case true:
                            int exist_fold = 0;
                            if (old_name != new_name)
                            {
                                while (Directory.Exists(new_name))
                                {
                                    exist_fold++;
                                    new_name = Path.Combine(cell1, cell3 + "(" + exist_fold.ToString() + ")");
                                }
                                Directory.Move(old_name, new_name);
                            }
                            break;
                    }
                }

                //删除_rename表，并显示完成结果
                workbook.Worksheets["_rename"].Delete();
                if (IsSheetExist(workbook, "_rename_备份"))
                {
                    workbook.Worksheets["_rename_备份"].Name = "_rename";
                }
                MessageBox.Show("文件名修改完毕");
                Process.Start(get_directory_path);
                readFile = 0;
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;

            }
            else
            {
                MessageBox.Show("没有选择文件夹，请先使用批读文件名功能后再使用该功能");
                readFile = 0;
            }
        }

        //文件目录选项
        private void select_f_or_d_Click(object sender, RibbonControlEventArgs e)
        {
            if (select_f_or_d.Checked == true)
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

        //判断指定工作簿中指定工作表名是否存在
        public static bool IsSheetExist(Excel.Workbook workbook, string sheetName)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Name == sheetName)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
