using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;



namespace ExcelAddIn
{
    public partial class Form1 : Form
    {

        private Excel.Workbook workbook;
        private string excelFilePath;
        private Int32 used_time_count = 0;
        private bool res = false;
        private Thread thread;
        private List<string> active_sheet_names = new List<string>();
        private List<string> new_sheet_names = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        //窗体初始化
        private void Form1_Load(object sender, EventArgs e)
        {
            this.MaximizeBox = false;
            //初始化tabcontrol控件
            tabControl1.SelectTab(0);
            workbook = ThisAddIn.app.ActiveWorkbook;
            excelFilePath = workbook.FullName;
            sheet_name_combobox.Items.Clear();
            field_name_combobox.Items.Clear();
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                sheet_name_combobox.Items.Add(worksheet.Name);
            }
            sheet_name_combobox.Refresh();
            field_name_combobox.Refresh();
            split_sheet_result_label.Visible = false;
            split_sheet_result_label.Text = "";
            split_sheet_progressBar.Visible = false;
            splitProgressBar_label.Visible = false;

            if (active_sheet_names.Count > 0)
            {
                active_sheet_names.Clear();
            }
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                active_sheet_names.Add(sheet.Name);
            }
            if (new_sheet_names.Count > 0)
            {
                new_sheet_names.Clear();
            }

            //初始化功能四右窗口
            which_field_label.Visible = false;
            which_field_combobox.Visible = false;
            what_type_label.Visible = false;
            what_type_combobox.Visible = false;
            regex_rule_label.Visible = false;
            regex_rule_textbox.Visible = false;
            run_result_label.Visible = false;
            regex_run_button.Visible = false;
            regex_clear_button.Visible = false;
            function_title_label.Text = "请选择所需使用的功能";
        }

        //重绘选项页布局
        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            //调整选项卡文字方向
            SolidBrush _Brush = new SolidBrush(Color.Black);//单色画刷
            RectangleF _TabTextArea = (RectangleF)tabControl1.GetTabRect(e.Index);//绘制区域
            StringFormat _sf = new StringFormat();//封装文本布局格式信息
            _sf.LineAlignment = StringAlignment.Center;
            _sf.Alignment = StringAlignment.Center;
            e.Graphics.DrawString(tabControl1.Controls[e.Index].Text, SystemInformation.MenuFont, _Brush, _TabTextArea, _sf);
        }

        //选项页初始化
        private async void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    sheet_name_combobox.Items.Clear();
                    field_name_combobox.Items.Clear();
                    foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                    {
                        sheet_name_combobox.Items.Add(worksheet.Name);
                    }
                    sheet_name_combobox.Refresh();
                    field_name_combobox.Refresh();
                    split_sheet_result_label.Visible = false;
                    split_sheet_result_label.Text = "";
                    split_sheet_progressBar.Visible = false;
                    splitProgressBar_label.Visible = false;
                    break;
                case 1:
                    merge_sheet_result_label.Visible = false;
                    merge_sheet_result_label.Text = "";
                    mergeProgressBar_label.Visible = false;
                    mergeProgressBar_label.Text = "";
                    merge_sheet_progressBar.Visible = false;
                    break;
                case 2:
                    sheet_listbox.Items.Clear();
                    await Task.Run(() =>
                    {
                        foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                        {
                            string worksheet_name = worksheet.Name;
                            sheet_listbox.Invoke((MethodInvoker)(() =>
                            {
                                sheet_listbox.Items.Add(worksheet_name);
                            }));
                        }
                    });
                    sheet_listbox.Refresh();
                    break;
                case 3:
                    break;
                case 4:
                    break;
                case 5:
                    this.Dispose();
                    break;
            }
        }

        //更新分表功能中选中表的字段选项
        private void sheet_name_combobox_TextChanged(object sender, EventArgs e)
        {
            split_sheet_result_label.Text = "";
            split_sheet_result_label.Visible = false;
            string selectworksheet = sheet_name_combobox.Text;
            if (string.IsNullOrEmpty(selectworksheet))
            {
                field_name_combobox.Items.Clear();
            }
            else
            {
                Excel.Worksheet worksheet = ThisAddIn.app.ActiveWorkbook.Worksheets[selectworksheet];
                field_name_combobox.Items.Clear();
                foreach (Excel.Range range in worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, worksheet.UsedRange.Columns.Count]])
                {
                    string range_value = range.Value2;
                    if (!string.IsNullOrEmpty(range_value))
                    {
                        field_name_combobox.Items.Add(range_value);
                    }
                }
                if (field_name_combobox.Items.Count > 0)
                {
                    field_name_combobox.Text = Convert.ToString(field_name_combobox.Items[0]);
                }
                else
                {
                    field_name_combobox.Text = "";
                }
            }
            field_name_combobox.Refresh();
        }

        //分表功能中清空combobox内容
        private void clear_button_Click(object sender, EventArgs e)
        {
            sheet_name_combobox.Text = "";
            field_name_combobox.Text = "";
        }

        //分表（UI主线程）
        private void split_button_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(sheet_name_combobox.Text) && string.IsNullOrEmpty(field_name_combobox.Text))
            {
                ShowLabel(split_sheet_result_label, true, "表和字段均不能为空！");
                StartTimer();
                return;
            }
            int field_column = 0;
            foreach (Excel.Range range in workbook.Worksheets[sheet_name_combobox.Text].Range[workbook.Worksheets[sheet_name_combobox.Text].Cells[1, 1], workbook.Worksheets[sheet_name_combobox.Text].Cells[1, workbook.Worksheets[sheet_name_combobox.Text].UsedRange.Columns.Count]])
            {
                if (range.Value == field_name_combobox.Text)
                {
                    field_column = range.Column;
                    break;
                }
            }
            string select_field = sheet_name_combobox.Text;
            thread = new Thread(() => SplitTask(select_field, field_column));
            thread.Start();
            split_sheet_result_label.Visible = true;
            split_sheet_timer.Interval = 1000;
            split_sheet_timer.Enabled = true;
            splitProgressBar_label.Visible = true;
            split_sheet_progressBar.Visible = true;
        }

        //分表（程序执行线程）
        private void SplitTask(string sheetName, int selectFieldsColumn)
        {
            res = false;
            tabControl1.Enabled = false;
            split_button.Enabled = false;
            splitsheet_export_button.Enabled = false;
            splitsheet_delete_button.Enabled = false;
            clear_button.Enabled = false;
            sheet_name_combobox.Enabled = false;
            field_name_combobox.Enabled = false;
            this.ControlBox = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            try
            {
                //声明范围列数、范围行数、分表依据列数、筛选结果第一列数
                List<string> records = new List<string>();
                int record_row = workbook.Worksheets[sheetName].UsedRange.Rows.Count;
                int current_record = 1;
                int total_record = 0;

                //将去重后的表名加入数组
                foreach (Excel.Range range in workbook.Worksheets[sheetName].Range[workbook.Worksheets[sheetName].Cells[2, selectFieldsColumn], workbook.Worksheets[sheetName].Cells[record_row, selectFieldsColumn]])
                {
                    if (records.Contains(range.Value) || string.IsNullOrEmpty(range.Value))
                    {
                        continue;
                    }
                    else
                    {
                        records.Add(range.Value);
                    }
                }

                total_record = records.Count;

                //动态更新一个分表工作簿中所有表的名称
                List<string> dynamic_sheet_name = new List<string>();


                //新建分表，并通过关键字段筛选，筛出结果复制到相应分表中
                foreach (string record in records)
                {
                    //更新进度条
                    UpdateProgressBar(split_sheet_progressBar, current_record, total_record, splitProgressBar_label, "分表进度");

                    //分表
                    if (dynamic_sheet_name.Count > 0)
                    {
                        dynamic_sheet_name.Clear();
                    }
                    foreach (Excel.Worksheet ws in workbook.Worksheets)
                    {
                        dynamic_sheet_name.Add(ws.Name);
                    }

                    Excel.Worksheet add_sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                    if (!dynamic_sheet_name.Contains(record))
                    {
                        add_sheet.Name = record;
                    }
                    else
                    {
                        int i = 1;
                        do { i++; } while (dynamic_sheet_name.Contains(record + i.ToString()));
                        add_sheet.Name = record + i.ToString();
                    }
                    workbook.Worksheets[sheetName].select();
                    ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells[1, 1], ThisAddIn.app.ActiveSheet.Cells[1, ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count]].Select();
                    ThisAddIn.app.Selection.AutoFilter(selectFieldsColumn, record);
                    ThisAddIn.app.ActiveSheet.Rows[1].Select();
                    ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.Selection, ThisAddIn.app.Selection.End(Excel.XlDirection.xlDown)].Select();
                    ThisAddIn.app.Selection.Copy(ThisAddIn.app.ActiveWorkbook.Worksheets[record].Range["A1"]);
                    current_record++;
                }
                ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells[1, 1], ThisAddIn.app.ActiveSheet.Cells[1, ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count]].AutoFilter();
                ThisAddIn.app.ActiveSheet.Range["A1"].Select();

                //对有序号列的表中序号重排序
                foreach (Excel.Worksheet worksheet in ThisAddIn.app.ActiveWorkbook.Worksheets)
                {
                    worksheet.Activate();
                    foreach (Excel.Range rng in ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells(1, 1), ThisAddIn.app.ActiveSheet.Cells(1, 1).End(Excel.XlDirection.xlToRight)])
                    {
                        if (rng.Value == "序号")
                        {
                            int tt = rng.Column;
                            for (int number = 1; number < ThisAddIn.app.ActiveSheet.UsedRange.Rows.count; number++)
                            {
                                ThisAddIn.app.ActiveSheet.Cells[number + 1, tt].Value = number;
                            }
                            break;
                        }
                    }
                }
                workbook.Worksheets[sheetName].Activate();
                ThisAddIn.app.ActiveSheet.Range("A1").Select();
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.CutCopyMode = Excel.XlCutCopyMode.xlCopy;
            }
            catch (Exception ex)
            {
                MessageBox.Show("选择的表或字段不正确，请核对后再试。错误问题：" + ex.Message);
            }
            finally
            {
                if (new_sheet_names.Count > 0)
                {
                    new_sheet_names.Clear();
                }
                foreach (Excel.Worksheet newsheet in workbook.Worksheets)
                {
                    if (!active_sheet_names.Contains(newsheet.Name))
                    {
                        new_sheet_names.Add(newsheet.Name);
                    }
                }

                tabControl1.Enabled = true;
                split_button.Enabled = true;
                splitsheet_export_button.Enabled = true;
                splitsheet_delete_button.Enabled = true;
                clear_button.Enabled = true;
                sheet_name_combobox.Enabled = true;
                field_name_combobox.Enabled = true;
                this.ControlBox = true;
                this.TopMost = false;
                res = true;
            }
        }

        //分表导出（UI主线程）
        private void splitsheet_export_button_Click(object sender, EventArgs e)
        {
            if (new_sheet_names.Count > 0)
            {
                thread = new Thread(() => splitsheetExportTask());
                thread.Start();
                split_sheet_result_label.Visible = true;
                split_sheet_timer.Interval = 1000;
                split_sheet_timer.Enabled = true;
                splitProgressBar_label.Visible = true;
                split_sheet_progressBar.Visible = true;
            }
            else
            {
                ShowLabel(split_sheet_result_label, true, "未找到本次分出的表，分表导出不成功。如确需导出表，请使用“批量导删”功能");
                StartTimer();
            }

        }

        //分表导出（程序执行线程）
        private void splitsheetExportTask()
        {
            res = false;
            tabControl1.Enabled = false;
            split_button.Enabled = false;
            splitsheet_export_button.Enabled = false;
            splitsheet_delete_button.Enabled = false;
            clear_button.Enabled = false;
            sheet_name_combobox.Enabled = false;
            field_name_combobox.Enabled = false;
            this.ControlBox = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            int current_sheet = 0;
            int total_sheet = new_sheet_names.Count;

            try
            {
                foreach (Excel.Worksheet exportsheet in workbook.Worksheets)
                {
                    //更新进度条
                    UpdateProgressBar(split_sheet_progressBar, current_sheet, total_sheet, splitProgressBar_label, "分表导出进度");

                    int i = 0;
                    string path = workbook.Path;
                    string create_dir = path + "\\分表导出文件";
                    if (new_sheet_names.Contains(exportsheet.Name))
                    {
                        string save_as1 = create_dir + "\\" + exportsheet.Name + ".xlsx";
                        string save_as2 = create_dir + "\\" + exportsheet.Name + i.ToString() + ".xlsx";
                        if (!Directory.Exists(create_dir))
                        {
                            Directory.CreateDirectory(create_dir);
                        }
                        if (!File.Exists(save_as1) && exportsheet.Name != ThisAddIn.app.ActiveWorkbook.Name.Split('.')[0])
                        {
                            Excel.Workbook exportworkbook = ThisAddIn.app.Workbooks.Add();
                            exportsheet.Copy(exportworkbook.Sheets[1]);
                            exportworkbook.Sheets[1].Name = exportsheet.Name;
                            exportworkbook.SaveAs(save_as1);
                            exportworkbook.Close();
                        }
                        else
                        {
                            do
                            {
                                i++;
                                save_as2 = create_dir + "\\" + exportsheet.Name + i.ToString() + ".xlsx";
                            } while (File.Exists(save_as2).ToString() == "true" || exportsheet.Name + i.ToString() == ThisAddIn.app.ActiveWorkbook.Name.Split('.')[0]);
                            Excel.Workbook exportworkbook = ThisAddIn.app.Workbooks.Add();
                            exportsheet.Copy(exportworkbook.Sheets[1]);
                            exportworkbook.Sheets[1].Name = exportsheet.Name;
                            exportworkbook.SaveAs(save_as2);
                            exportworkbook.Close();
                        }
                    }
                    current_sheet++;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("导出错误，原因是：" + ex.Message);
            }
            finally
            {
                tabControl1.Enabled = true;
                split_button.Enabled = true;
                splitsheet_export_button.Enabled = true;
                splitsheet_delete_button.Enabled = true;
                clear_button.Enabled = true;
                sheet_name_combobox.Enabled = true;
                field_name_combobox.Enabled = true;
                this.ControlBox = true;
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;
                res = true;
            }
        }

        //分表删除
        private void splitsheet_delete_button_Click(object sender, EventArgs e)
        {
            tabControl1.Enabled = false;
            split_button.Enabled = false;
            splitsheet_export_button.Enabled = false;
            splitsheet_delete_button.Enabled = false;
            clear_button.Enabled = false;
            sheet_name_combobox.Enabled = false;
            field_name_combobox.Enabled = false;
            this.ControlBox = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            splitProgressBar_label.Visible = false;
            split_sheet_progressBar.Visible = false;

            if (new_sheet_names.Count > 0)
            {
                foreach (Excel.Worksheet deletesheet in workbook.Worksheets)
                {
                    if (new_sheet_names.Contains(deletesheet.Name))
                    {
                        deletesheet.Delete();
                    }
                }
                ShowLabel(split_sheet_result_label, true, "分表删除完成");
                StartTimer();
            }
            else
            {
                ShowLabel(split_sheet_result_label, true, "未找到本次分出的表，分表删除不成功。如确需删除表，请使用“批量导删”功能");
                StartTimer();
            }

            tabControl1.Enabled = true;
            split_button.Enabled = true;
            splitsheet_export_button.Enabled = true;
            splitsheet_delete_button.Enabled = true;
            clear_button.Enabled = true;
            sheet_name_combobox.Enabled = true;
            field_name_combobox.Enabled = true;
            this.ControlBox = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }


        //并表功能中的选择文件夹按钮
        private void dir_select_button_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "请选择导出到文件夹";
            folderBrowserDialog1.ShowDialog();
            string select_export_path = folderBrowserDialog1.SelectedPath;
            if (!string.IsNullOrEmpty(select_export_path))
            {
                dir_select_textbox.Text = select_export_path;
            }
            else
            {
                MessageBox.Show("未选择需合并文件所在文件夹");
            }
        }


        //同一工作簿并表（UI主线程）
        private void single_merge_button_Click(object sender, EventArgs e)
        {
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;
            foreach (Excel.Worksheet source_sheet in workbook.Worksheets)
            {
                if (source_sheet.Name == "并表汇总")
                {
                    MessageBox.Show("已存在名称为并表汇总的表，请将该表改名后再试", "注意！");
                    return;
                }
            }
            Int32 data_start_row = Convert.ToInt32(ThisAddIn.app.InputBox(@"请输入数据起始行的行数（所输数字应大于等于2,若不输入或输入小于2数字则默认数据起始行为第2行）", "输入数据起始行"));
            if (data_start_row < 2)
            {
                data_start_row = 2;
            }
            //在List<string>中保存当前所有表名
            if (active_sheet_names.Count > 0)
            {
                active_sheet_names.Clear();
            }
            foreach (Excel.Worksheet active_sheet_name in workbook.Worksheets)
            {
                active_sheet_names.Add(active_sheet_name.Name);
            }

            //启动并表线程
            thread = new Thread(() => mergeTask(data_start_row, active_sheet_names));
            thread.Start();
            merge_sheet_timer.Interval = 1000;
            merge_sheet_timer.Enabled = true;
            merge_sheet_result_label.Visible = true;
            mergeProgressBar_label.Visible = true;
            merge_sheet_progressBar.Visible = true;
        }



        //同一工作簿并表（程序执行线程）
        private void mergeTask(Int32 titleRow, List<string> unMergeSheets, bool exist_bool = true)
        {
            try
            {
                tabControl1.Enabled = false;
                single_merge_button.Enabled = false;
                dir_select_textbox.Enabled = false;
                dir_select_button.Enabled = false;
                multi_merge_button.Enabled = false;
                multi_merge_sheet_checkBox.Enabled = false;
                ThisAddIn.app.ScreenUpdating = false;
                ThisAddIn.app.DisplayAlerts = false;

                Excel.Worksheet destination_sheet = workbook.Worksheets.Add(Before: workbook.Sheets[1]);
                destination_sheet.Name = "并表汇总";
                destination_sheet.Activate();
                destination_sheet.Range["A1"].Select();

                //在合并表中粘贴标题行
                workbook.Sheets[workbook.Worksheets.Count].Rows["1:" + Convert.ToString(titleRow - 1)].Copy(destination_sheet.Cells[1, 1]);

                int current_sheet = 1;
                int total_sheet = workbook.Worksheets.Count;

                //合并各表中数据行
                switch (exist_bool)
                {
                    case true:
                        foreach (Excel.Worksheet source_sheet in workbook.Worksheets)
                        {
                            if (source_sheet.Name != "并表汇总")
                            {
                                //更新进度条
                                UpdateProgressBar(merge_sheet_progressBar, current_sheet, total_sheet - 1, mergeProgressBar_label, "并表进度");

                                //destination_range为汇总表A列有数据行下一格
                                Excel.Range destination_range = destination_sheet.Range["A" + destination_sheet.UsedRange.Rows.Count.ToString()].Offset[1, 0];

                                //source_range为分表的要复制区域
                                Excel.Range source_range = source_sheet.Range[source_sheet.Cells[titleRow, 1], source_sheet.Cells[source_sheet.UsedRange.Rows.Count, source_sheet.UsedRange.Columns.Count]];
                                source_range.Copy(destination_range);
                                current_sheet++;
                                if (unMergeSheets.Contains(source_sheet.Name) == false)
                                {
                                    source_sheet.Delete();
                                }
                            }
                        }
                        //重写合并表中的序号列
                        destination_sheet.Activate();
                        foreach (Excel.Range destination_title_range in destination_sheet.Range[destination_sheet.Cells[titleRow - 1, 1], destination_sheet.Cells[titleRow - 1, destination_sheet.UsedRange.Columns.Count]])
                        {
                            if (destination_title_range.Value == "序号")
                            {
                                for (int i = 1; i <= destination_sheet.UsedRange.Columns.Count - titleRow + 1; i++)
                                {
                                    destination_sheet.Cells[titleRow - 1 + i, destination_title_range.Column].Value = i;
                                }
                            }
                        }
                        break;
                    case false:
                        foreach (Excel.Worksheet source_sheet in workbook.Worksheets)
                        {
                            if (source_sheet.Name != "并表汇总" && unMergeSheets.Contains(source_sheet.Name) == false)
                            {
                                //更新进度条
                                UpdateProgressBar(merge_sheet_progressBar, current_sheet, total_sheet - unMergeSheets.Count - 1, mergeProgressBar_label, "并表进度");

                                //destination_range为汇总表A列有数据行下一格
                                Excel.Range destination_range = destination_sheet.Range["A" + destination_sheet.UsedRange.Rows.Count.ToString()].Offset[1, 0];

                                //source_range为分表的要复制区域
                                Excel.Range source_range = source_sheet.Range[source_sheet.Cells[titleRow, 1], source_sheet.Cells[source_sheet.UsedRange.Rows.Count, source_sheet.UsedRange.Columns.Count]];
                                source_range.Copy(destination_range);
                                current_sheet++;
                                source_sheet.Delete();
                            }
                        }
                        //重写合并表中的序号列
                        destination_sheet.Activate();
                        foreach (Excel.Range destination_title_range in destination_sheet.Range[destination_sheet.Cells[titleRow - 1, 1], destination_sheet.Cells[titleRow - 1, destination_sheet.UsedRange.Columns.Count]])
                        {
                            if (destination_title_range.Value == "序号")
                            {
                                for (int i = 1; i <= destination_sheet.UsedRange.Columns.Count - titleRow + 1; i++)
                                {
                                    destination_sheet.Cells[titleRow - 1 + i, destination_title_range.Column].Value = i;
                                }
                            }
                        }
                        break;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("合并工作表未成功，错误原因是：" + ex.Message);
            }
            finally
            {
                tabControl1.Enabled = true;
                single_merge_button.Enabled = true;
                dir_select_textbox.Enabled = true;
                dir_select_button.Enabled = true;
                multi_merge_button.Enabled = true;
                multi_merge_sheet_checkBox.Enabled = true;
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;
                res = true;
            }
        }

        //不同工作簿并表（UI主线程）
        private void multi_merge_button_Click(object sender, EventArgs e)
        {
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;
            foreach (Excel.Worksheet source_sheet in workbook.Worksheets)
            {
                if (source_sheet.Name == "并表汇总")
                {
                    MessageBox.Show("已存在名称为并表汇总的表，请将该表改名后再试", "注意！");
                    return;
                }
            }
            Int32 data_start_row = Convert.ToInt32(ThisAddIn.app.InputBox(@"请输入数据起始行的行数（所输数字应大于等于2,若不输入或输入小于2数字则默认数据起始行为第2行）", "输入数据起始行"));
            if (data_start_row < 2)
            {
                data_start_row = 2;
            }

            //在List<string>中保存当前所有表名
            if (active_sheet_names.Count > 0)
            {
                active_sheet_names.Clear();
            }
            foreach (Excel.Worksheet active_sheet_name in workbook.Worksheets)
            {
                active_sheet_names.Add(active_sheet_name.Name);
            }

            //启动并表线程
            thread = new Thread(() => multiMergeTask(data_start_row, active_sheet_names));
            thread.Start();
            merge_sheet_timer.Interval = 1000;
            merge_sheet_timer.Enabled = true;
            mergeProgressBar_label.Visible = true;
            merge_sheet_result_label.Visible = true;
            merge_sheet_progressBar.Visible = true;
        }

        //不同工作簿并表（程序执行线程）
        private void multiMergeTask(Int32 titleRow, List<string> activeSheetNames)
        {
            tabControl1.Enabled = false;
            single_merge_button.Enabled = false;
            dir_select_textbox.Enabled = false;
            dir_select_button.Enabled = false;
            multi_merge_button.Enabled = false;
            multi_merge_sheet_checkBox.Enabled = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            //将指定文件夹内所有工作簿的所有表转至当前工作簿内
            Excel.Workbook destination_workbook = ThisAddIn.app.ActiveWorkbook;
            string destination_workbook_name = destination_workbook.Name;
            DirectoryInfo folder = new DirectoryInfo(dir_select_textbox.Text);
            string source_workbook_name = null;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;


            //获取当前打开excel文件名称
            int totalFile_count = get_File_Count(dir_select_textbox.Text, "*.xls*");
            int currentFile_count = 0;
            foreach (FileInfo file in folder.GetFiles("*.xls*", SearchOption.AllDirectories).Where(file => !file.Attributes.HasFlag(FileAttributes.Hidden)))
            {
                if (file.Name != ThisAddIn.app.ActiveWorkbook.Name)
                {
                    //更新进度条
                    UpdateProgressBar(merge_sheet_progressBar, currentFile_count, totalFile_count, mergeProgressBar_label, "转移表进度");

                    Excel.Workbook source_excel_workbook = ThisAddIn.app.Workbooks.Open(file.FullName);
                    for (int i = 1; i <= source_excel_workbook.Worksheets.Count; i++)
                    {
                        if (source_excel_workbook.Worksheets[i].UsedRange.cells.count == 1 && string.IsNullOrEmpty(Convert.ToString(source_excel_workbook.Worksheets[i].UsedRange.cells[1, 1].Value)))
                        {
                            continue;
                        }
                        else
                        {
                            source_workbook_name = source_excel_workbook.Name.Split('.')[0];
                            string source_sheet_name = source_excel_workbook.Worksheets[i].Name;
                            string destination_sheet_name = source_workbook_name + "_" + source_sheet_name;
                            source_excel_workbook.Worksheets[i].Activate();
                            ThisAddIn.app.ActiveSheet.UsedRange.Copy();
                            ThisAddIn.app.Workbooks[destination_workbook_name].Activate();
                            Excel.Worksheet add_sheet = ThisAddIn.app.ActiveWorkbook.Worksheets.Add(After: ThisAddIn.app.ActiveWorkbook.Worksheets[ThisAddIn.app.ActiveWorkbook.Worksheets.Count]);
                            add_sheet.Name = destination_sheet_name;
                            ThisAddIn.app.ActiveWorkbook.Worksheets[destination_sheet_name].Activate();
                            ThisAddIn.app.ActiveSheet.Range["A1"].Select();
                            ThisAddIn.app.Selection.PasteSpecial();
                            source_excel_workbook.Activate();
                            ThisAddIn.app.CutCopyMode = Excel.XlCutCopyMode.xlCopy;
                        }
                    }
                    source_excel_workbook.Close(false);
                    currentFile_count++;
                }
            }
            ThisAddIn.app.Workbooks[destination_workbook_name].Activate();
            ThisAddIn.app.ActiveWorkbook.Sheets[1].Select();
            ThisAddIn.app.ActiveSheet.Range["A1"].Select();
            ThisAddIn.app.ActiveWorkbook.RefreshAll();
            ThisAddIn.app.ActiveWorkbook.Save();
            Excel.Application excelApp = ThisAddIn.app;
            foreach (Excel.Workbook opened_workbook in excelApp.Workbooks)
            {
                if (opened_workbook.Name.Split('.')[0] == source_workbook_name)
                {
                    opened_workbook.Close(false);
                }
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;
            }
            if (multi_merge_sheet_checkBox.Checked)
            {
                mergeTask(titleRow, activeSheetNames, true);
            }
            else
            {
                mergeTask(titleRow, activeSheetNames, false);
            }

            tabControl1.Enabled = true;
            single_merge_button.Enabled = true;
            dir_select_textbox.Enabled = true;
            dir_select_button.Enabled = true;
            multi_merge_button.Enabled = true;
            multi_merge_sheet_checkBox.Enabled = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }



        //进度条更新函数
        private void UpdateProgressBar(System.Windows.Forms.ProgressBar progressBar, int currentSheet, int totalSheets, System.Windows.Forms.Label progressBar_result_label, string progressBar_result)
        {
            // 计算进度百分比
            int progressPercentage = (int)((double)currentSheet / totalSheets * 100);
            // 更新进度条控件
            progressBar.Value = progressPercentage;
            progressBar.Update();
            // 显示百分比数字
            progressBar_result_label.Text = progressBar_result + progressPercentage.ToString() + "%";
        }



        //获取指定文件夹符合要求文件的数量（包含子文件夹）
        private int get_File_Count(string dir_path, string ext)
        {
            List<string> files = new List<string>();
            DirectoryInfo folder = new DirectoryInfo(dir_path);
            foreach (FileInfo file in folder.GetFiles(ext, SearchOption.AllDirectories).Where(file => !file.Attributes.HasFlag(FileAttributes.Hidden)))
            {
                files.Add(file.Name);
            }
            return files.Count;
        }



        //批量导删中的checkbox被按下时
        private void all_select_checkbox_Click(object sender, EventArgs e)
        {
            if (all_select_checkbox.Checked == true)
            {
                all_select_checkbox.Text = "全部取消";
                for (int i = 0; i <= sheet_listbox.Items.Count - 1; i++)
                {
                    sheet_listbox.SetSelected(i, true);
                }
            }
            else
            {
                all_select_checkbox.Text = "全部选中";
                for (var i = 0; i <= sheet_listbox.Items.Count - 1; i++)
                {
                    sheet_listbox.SetSelected(i, false);
                }
            }
        }

        //批量导删中的listbox和checkbox联动
        private void sheet_listbox_SelectedValueChanged(object sender, EventArgs e)
        {
            if (sheet_listbox.Items.Count != sheet_listbox.SelectedItems.Count)
            {
                if (all_select_checkbox.Text == "全部取消")
                {
                    all_select_checkbox.Text = "全部选中";
                    all_select_checkbox.Checked = false;
                }
            }
            else
            {
                if (all_select_checkbox.Text == "全部选中")
                {
                    all_select_checkbox.Text = "全部取消";
                    all_select_checkbox.Checked = true;
                }
            }
        }


        //批量导出当前工作簿中的表
        private void batch_export_button_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "请选择导出到文件夹";
            folderBrowserDialog1.ShowDialog();
            string select_export_path = folderBrowserDialog1.SelectedPath;
            if (!string.IsNullOrEmpty(select_export_path))
            {
                tabControl1.Enabled = false;
                sheet_listbox.Enabled = false;
                all_select_checkbox.Enabled = false;
                batch_export_button.Enabled = false;
                batch_delete_button.Enabled = false;
                ThisAddIn.app.ScreenUpdating = false;
                ThisAddIn.app.DisplayAlerts = false;

                foreach (string item in sheet_listbox.SelectedItems)
                {
                    Excel.Workbook export_workbook = ThisAddIn.app.Workbooks.Add();
                    workbook.Worksheets[item].Activate();
                    workbook.Worksheets[item].Copy(export_workbook.Sheets[1]);
                    export_workbook.Sheets[1].Name = item;
                    export_workbook.SaveAs(select_export_path + "\\" + item + ".xlsx");
                    export_workbook.Close();
                }
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
                MessageBox.Show("所选分表已导出到指定文件夹");
            }
            else
            {
                MessageBox.Show("未选择导出文件夹");
            }

            tabControl1.Enabled = true;
            sheet_listbox.Enabled = true;
            all_select_checkbox.Enabled = true;
            batch_export_button.Enabled = true;
            batch_delete_button.Enabled = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }



        //批量删除当前工作簿中的表
        private void batch_delete_button_Click(object sender, EventArgs e)
        {
            tabControl1.Enabled = false;
            sheet_listbox.Enabled = false;
            all_select_checkbox.Enabled = false;
            batch_export_button.Enabled = false;
            batch_delete_button.Enabled = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            int selected_sheet_count = sheet_listbox.SelectedItems.Count;
            if (selected_sheet_count == sheet_listbox.Items.Count)
            {
                MessageBox.Show("批量删除时不能一次性删除所有表，需至少保留一张表");
            }
            else
            {
                foreach (string item in sheet_listbox.SelectedItems)
                {
                    workbook.Worksheets[item].Delete();

                }
                sheet_listbox.Items.Clear();
                foreach (Excel.Worksheet left_sheet in workbook.Worksheets)
                {
                    sheet_listbox.Items.Add(left_sheet.Name);
                }
                sheet_listbox.Refresh();
            }

            tabControl1.Enabled = true;
            sheet_listbox.Enabled = true;
            all_select_checkbox.Enabled = true;
            batch_export_button.Enabled = true;
            batch_delete_button.Enabled = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }



        //目录下多工作簿的表转同一个工作簿内（UI主线程）
        private void move_sheet_button_Click(object sender, EventArgs e)
        {
            //右侧功能区初始化
            which_field_label.Visible = false;
            which_field_combobox.Visible = false;
            what_type_label.Visible = false;
            what_type_combobox.Visible = false;
            regex_rule_label.Visible = false;
            regex_rule_textbox.Visible = false;
            run_result_label.Visible = false;
            regex_run_button.Visible = false;
            regex_clear_button.Visible = false;
            function_title_label.Text = "不同工作簿中的表全部复制到本工作簿";
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;
            //this.TopMost = true;

            //左侧按钮状态改变
            move_sheet_button.Enabled = false;
            add_sheet_button.Enabled = false;
            transposition_button.Enabled = false;
            regex_button.Enabled = false;
            payslip_button.Enabled = false;
            contents_button.Enabled = false;

            folderBrowserDialog1.Description = "请选择工作簿所在文件夹";
            folderBrowserDialog1.ShowDialog();
            string select_fold_path = folderBrowserDialog1.SelectedPath;

            //转移文件夹，启动多线程并等待执行结果
            if (!string.IsNullOrEmpty(select_fold_path))
            {
                string result = null;
                Task.Run(() =>
                {
                    // 启动多线程执行长时间任务
                    result = movesheetTask(select_fold_path);

                }).ContinueWith((task) =>
                {
                    // 长时间任务完成后启动定时器
                    if (result == "finished")
                    {
                        ShowLabel(run_result_label, true, "转移工作表成功完成");
                        StartTimer();

                    }
                    else
                    {
                        ShowLabel(run_result_label, true, "转移工作表错误，原因是" + result);
                        StartTimer();
                    }
                }, TaskScheduler.FromCurrentSynchronizationContext());
            }
            else
            {
                ShowLabel(run_result_label, true, "未正确选择文件夹");
                StartTimer();
            }

            //左侧按钮状态改变
            move_sheet_button.Enabled = true;
            add_sheet_button.Enabled = true;
            transposition_button.Enabled = true;
            regex_button.Enabled = true;
            payslip_button.Enabled = true;
            contents_button.Enabled = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
            //this.TopMost = false;
        }

        //目录下多工作簿表转同一工作簿内（程序执行线程）
        private string movesheetTask(string get_fold_path)
        {
            Excel.Workbook destination_workbook = ThisAddIn.app.ActiveWorkbook;
            string destination_workbook_name = destination_workbook.Name;
            DirectoryInfo folder = new DirectoryInfo(get_fold_path);
            string source_workbook_name = null;

            try
            {
                //获取当前打开excel文件名称

                foreach (FileInfo file in folder.GetFiles("*.xls*", SearchOption.AllDirectories).Where(file => !file.Attributes.HasFlag(FileAttributes.Hidden)))
                {
                    if (file.Name != ThisAddIn.app.ActiveWorkbook.Name)
                    {
                        Excel.Workbook source_excel_workbook = ThisAddIn.app.Workbooks.Open(file.FullName);
                        for (int i = 1; i <= source_excel_workbook.Worksheets.Count; i++)
                        {
                            if (source_excel_workbook.Worksheets[i].UsedRange.cells.count == 1 && string.IsNullOrEmpty(Convert.ToString(source_excel_workbook.Worksheets[i].UsedRange.cells[1, 1].Value)))
                            {
                                continue;
                            }
                            else
                            {
                                source_workbook_name = source_excel_workbook.Name.Split('.')[0];
                                string source_sheet_name = source_excel_workbook.Worksheets[i].Name;
                                string destination_sheet_name = source_workbook_name + "_" + source_sheet_name;
                                source_excel_workbook.Worksheets[i].Activate();
                                ThisAddIn.app.ActiveSheet.UsedRange.Copy();
                                ThisAddIn.app.Workbooks[destination_workbook_name].Activate();
                                Excel.Worksheet add_sheet = ThisAddIn.app.ActiveWorkbook.Worksheets.Add(After: ThisAddIn.app.ActiveWorkbook.Worksheets[ThisAddIn.app.ActiveWorkbook.Worksheets.Count]);
                                add_sheet.Name = destination_sheet_name;
                                ThisAddIn.app.ActiveWorkbook.Worksheets[destination_sheet_name].Activate();
                                ThisAddIn.app.ActiveSheet.Range["A1"].Select();
                                ThisAddIn.app.Selection.PasteSpecial();
                                source_excel_workbook.Activate();
                                ThisAddIn.app.CutCopyMode = Excel.XlCutCopyMode.xlCopy;
                            }
                        }
                        source_excel_workbook.Close(false);
                    }
                }
                ThisAddIn.app.Workbooks[destination_workbook_name].Activate();
                ThisAddIn.app.ActiveWorkbook.Sheets[1].Select();
                ThisAddIn.app.ActiveSheet.Range["A1"].Select();
                ThisAddIn.app.ActiveWorkbook.RefreshAll();
                ThisAddIn.app.ActiveWorkbook.Save();
                return "finished";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                Excel.Application excelApp = ThisAddIn.app;
                foreach (Excel.Workbook opened_workbook in excelApp.Workbooks)
                {
                    if (opened_workbook.Name.Split('.')[0] == source_workbook_name)
                    {
                        opened_workbook.Close(false);
                    }
                }
            }
        }



        //一键建立多个工作表
        private void add_sheet_button_Click(object sender, EventArgs e)
        {
            //右侧功能区初始化
            function_title_label.Text = "建立指定名称和数量的新工作表";
            which_field_label.Visible = false;
            which_field_combobox.Visible = false;
            what_type_label.Visible = false;
            what_type_combobox.Visible = false;
            regex_rule_label.Visible = false;
            regex_rule_textbox.Visible = false;
            run_result_label.Visible = false;
            regex_run_button.Visible = false;
            regex_clear_button.Visible = false;

            //左侧按钮状态改变
            tabControl1.Enabled = false;
            move_sheet_button.Enabled = false;
            add_sheet_button.Enabled = false;
            transposition_button.Enabled = false;
            regex_button.Enabled = false;
            payslip_button.Enabled = false;
            contents_button.Enabled = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            string shtname = "";
            int i = 0;
            ThisAddIn.app.DisplayAlerts = false;
            ThisAddIn.app.ScreenUpdating = false;
            string activated_sheet_name = ThisAddIn.app.ActiveSheet.Name;
            int n = Convert.ToInt32(ThisAddIn.app.InputBox("请输入需要新建空表数量：", "输入建表数量"));

            if (n > 0)
            {
                shtname = Convert.ToString(ThisAddIn.app.InputBox("请输入表统一名称,未输入则缺省命名为‘新建表’：", "输入表名称"));
                string pattern = @"[、/?？*\[\]]";
                if (ContainsSpecialChars(shtname, pattern))
                {
                    MessageBox.Show("表名输入不合法，将按照缺省名称建表");
                    shtname = "新建表";
                }
                else if (string.IsNullOrEmpty(shtname))
                {
                    shtname = "新建表";
                }
                for (i = 1; i <= n; i++)
                {
                    Excel.Worksheet totelsheet = ThisAddIn.app.ActiveWorkbook.Worksheets.Add(After: ThisAddIn.app.ActiveWorkbook.Worksheets[ThisAddIn.app.ActiveWorkbook.Worksheets.Count]);
                    totelsheet.Name = shtname + Convert.ToString(i);
                }
                Excel.Worksheet originalWorksheet = (Excel.Worksheet)ThisAddIn.app.ActiveWorkbook.Sheets[activated_sheet_name];
                originalWorksheet.Activate();
                Excel.Range selectrange = ThisAddIn.app.ActiveSheet.Range["A1"];
                selectrange.Select();
                ThisAddIn.app.ActiveWorkbook.Save();
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;
                ShowLabel(run_result_label, true, "新建表完成");
                StartTimer();
            }
            else
            {
                ShowLabel(run_result_label, true, "未正确输入新建表数量");
                StartTimer();
            }

            //左侧按钮状态改变
            tabControl1.Enabled = true;
            move_sheet_button.Enabled = true;
            add_sheet_button.Enabled = true;
            transposition_button.Enabled = true;
            regex_button.Enabled = true;
            payslip_button.Enabled = true;
            contents_button.Enabled = true;
            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;
        }



        //转置工作表(UI主线程）
        private void transposition_button_Click(object sender, EventArgs e)
        {
            //右侧功能区初始化
            function_title_label.Text = "将列名称转置为字段内数据";
            which_field_label.Visible = false;
            which_field_combobox.Visible = false;
            what_type_label.Visible = false;
            what_type_combobox.Visible = false;
            regex_rule_label.Visible = false;
            regex_rule_textbox.Visible = false;
            run_result_label.Visible = false;
            regex_run_button.Visible = false;
            regex_clear_button.Visible = false;

            //左侧按钮状态改变
            tabControl1.Enabled = false;
            move_sheet_button.Enabled = false;
            add_sheet_button.Enabled = false;
            transposition_button.Enabled = false;
            regex_button.Enabled = false;
            payslip_button.Enabled = false;
            contents_button.Enabled = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            int result = 0;
            Task.Run(() =>
            {
                // 启动多线程执行长时间任务
                result = transTask();

            }).ContinueWith((task) =>
            {
                // 长时间任务完成后启动定时器
                switch (result)
                {
                    case 0:
                        ShowLabel(run_result_label, true, "转置完成");
                        StartTimer();
                        break;
                    case -1:
                        ShowLabel(run_result_label, true, "转置开始列数字输入错误");
                        StartTimer();
                        break;
                    case 1:
                        ShowLabel(run_result_label, true, "新建字段的名称输入错误，不能为空或“False”关键字");
                        StartTimer();
                        break;

                }
            }, TaskScheduler.FromCurrentSynchronizationContext());

            //左侧按钮状态改变
            tabControl1.Enabled = true;
            move_sheet_button.Enabled = true;
            add_sheet_button.Enabled = true;
            transposition_button.Enabled = true;
            regex_button.Enabled = true;
            payslip_button.Enabled = true;
            contents_button.Enabled = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }

        //转置工作表（程序执行线程）
        private int transTask()
        {

            Excel.Worksheet worksheet = ThisAddIn.app.ActiveSheet;
            //获取当前表名称
            string active_sheet_name = worksheet.Name;
            //获取当前表全部行数
            long row_count = worksheet.Rows.Count;
            //获取当前表全部列数
            long column_count = worksheet.Columns.Count;
            //获取最后数据行数
            long used_row_count = worksheet.UsedRange.Rows.Count;
            //获取最后数据列数
            long used_column_count = worksheet.UsedRange.Columns.Count;
            string trans_sheet_name = active_sheet_name + "转置表";

            //设置转置开始列
            string translation_start_column = Convert.ToString(ThisAddIn.app.InputBox("请输入从第几列（不小于2的数字）开始转置：", "注意"));
            string pat = @"^[1-9]\d*$";
            if (ContainsSpecialChars(translation_start_column, pat) == false || Convert.ToUInt32(translation_start_column) > used_column_count || translation_start_column == "False")
            {
                return -1;
            }

            //设置转置后字段名
            string field_name = Convert.ToString(ThisAddIn.app.InputBox("请输入转置列的字段名称：", "注意"));
            if (string.IsNullOrEmpty(field_name) || field_name == "False")
            {
                return 1;
            }

            ThisAddIn.app.DisplayAlerts = false;
            ThisAddIn.app.ScreenUpdating = false;

            //将表中空值补为0
            foreach (Excel.Range range in ThisAddIn.app.ActiveSheet.Range(ThisAddIn.app.ActiveSheet.Cells(1, 1), ThisAddIn.app.ActiveSheet.Cells(used_row_count, used_column_count)))
            {
                object value = range.Value;
                if (string.IsNullOrEmpty(value.ToString()))
                {
                    range.Value = 0;
                }
            }

            int translation_start_column1 = Convert.ToInt32(translation_start_column); //将转置起始列转为数值
            //新建“转置表”
            Excel.Worksheet trans_sheet = ThisAddIn.app.ActiveWorkbook.Worksheets.Add(Before: worksheet);
            trans_sheet.Name = trans_sheet_name;
            for (int s = 1; s < translation_start_column1; s++)
            {
                trans_sheet.Cells[1, s] = ThisAddIn.app.ActiveWorkbook.Worksheets[active_sheet_name].Cells[1, s];
            }
            trans_sheet.Cells[1, Convert.ToInt32(translation_start_column)] = "数值";
            string pattern = @"^[-+]?[0-9]*\.?[0-9]+$";
            if (ContainsSpecialChars(ThisAddIn.app.ActiveWorkbook.Worksheets[active_sheet_name].Cells[1, translation_start_column1].value, pattern))
            {
                trans_sheet.Columns[translation_start_column1].NumberFormatLocal = "#,##0.00";
            }
            trans_sheet.Cells[1, translation_start_column1 + 1] = field_name;
            if (field_name == "日期")
            {
                trans_sheet.Columns[Convert.ToInt32(translation_start_column) + 1].NumberFormatLocal = "yyyy-m-d";
            }
            trans_sheet.Activate();


            //复制粘贴转置内容
            for (int n = translation_start_column1; n <= used_column_count; n++) //循环重复数据列次
            {
                //复制粘贴固定字段
                worksheet.Activate();
                worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[used_row_count, translation_start_column1 - 1]].Select();
                ThisAddIn.app.Selection.Copy();
                trans_sheet.Activate();
                ThisAddIn.app.ActiveSheet.Cells[row_count, 1].End(Excel.XlDirection.xlUp).Offset(1, 0).Select();
                ThisAddIn.app.ActiveSheet.PasteSpecial();
                ThisAddIn.app.Application.CutCopyMode = Excel.XlCutCopyMode.xlCopy;

                ////复制粘贴转置字段
                worksheet.Activate();
                worksheet.Range[worksheet.Cells[2, n], worksheet.Cells[used_row_count, n]].Select();
                ThisAddIn.app.Selection.Copy();
                trans_sheet.Activate();
                ThisAddIn.app.ActiveSheet.Cells[row_count, translation_start_column1].End(Excel.XlDirection.xlUp).Offset(1, 0).Select();
                ThisAddIn.app.Selection.PasteSpecial();
                ThisAddIn.app.Application.CutCopyMode = Excel.XlCutCopyMode.xlCopy;

                //复制粘贴新建字段
                worksheet.Activate();
                worksheet.Cells[1, n].Copy();
                trans_sheet.Activate();
                ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells[row_count, translation_start_column1 + 1].End(Excel.XlDirection.xlUp).Offset(1, 0), ThisAddIn.app.ActiveSheet.Cells[ThisAddIn.app.ActiveSheet.UsedRange.Rows.count, translation_start_column1 + 1]].Select();
                ThisAddIn.app.Selection.PasteSpecial();
                ThisAddIn.app.Application.CutCopyMode = Excel.XlCutCopyMode.xlCopy;
            }
            worksheet.Select();
            worksheet.Range["A1"].Select();
            trans_sheet.Activate();
            trans_sheet.Range["A1"].Select();
            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;
            return 0;
        }



        //正则表达式功能激活
        private void regex_button_Click(object sender, EventArgs e)
        {
            function_title_label.Text = "正则表达式提取指定内容";
            which_field_label.Visible = true;
            which_field_combobox.Visible = true;
            what_type_label.Visible = true;
            what_type_combobox.Visible = true;
            regex_run_button.Visible = true;
            regex_clear_button.Visible = true;
            run_result_label.Visible = false;

        }

        //正则表达式功能区中各控件可见时
        private async void which_field_combobox_VisibleChanged(object sender, EventArgs e)
        {
            if (Visible)
            {
                which_field_combobox.Items.Clear();
                await Task.Run(() =>
                {
                    long usedrange_columns = ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count;
                    foreach (Excel.Range cell in ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells[1, 1], ThisAddIn.app.ActiveSheet.Cells[1, usedrange_columns]])
                    {
                        string cellValue = cell.Value?.ToString();
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            Invoke(new Action(() =>
                            {
                                which_field_combobox.Items.Add($"列{cell.Column}空");
                            }));
                        }
                        else
                        {
                            Invoke(new Action(() =>
                            {
                                which_field_combobox.Items.Add(cell.Value);
                            }));
                        }
                    }
                });
            }
            else
            {
                which_field_combobox.Items.Clear();
            }
        }

        //正则表达式功能区提取内容如果选择自定义时，过滤规则标签和文本框显示，否则不显示
        private void what_type_combobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (what_type_combobox.SelectedItem is null)
            {
                regex_rule_label.Visible = false;
                regex_rule_textbox.Visible = false;
                regex_rule_textbox.Text = "";
            }
            else
            {
                string selectvalue = what_type_combobox.SelectedItem.ToString();
                if (selectvalue == "自定义")
                {
                    regex_rule_label.Visible = true;
                    regex_rule_textbox.Visible = true;
                }
                else
                {
                    regex_rule_label.Visible = false;
                    regex_rule_textbox.Visible = false;
                    regex_rule_textbox.Text = "";
                }
            }
        }

        //正则表达式提取内容
        private void regex_run_button_Click(object sender, EventArgs e)
        {
            tabControl1.Enabled = false;
            move_sheet_button.Enabled = false;
            add_sheet_button.Enabled = false;
            transposition_button.Enabled = false;
            regex_button.Enabled = false;
            payslip_button.Enabled = false;
            contents_button.Enabled = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            Excel.Worksheet ws = ThisAddIn.app.ActiveSheet;

            //定义已有数据范围的行数变量
            long rown = ws.UsedRange.Rows.Count;
            //定义已有数据范围的列数变量
            long coln = ws.UsedRange.Columns.Count;

            //定义所选择列变量
            long col = 0;

            //窗体内选择需过滤的数据列和过滤规则
            if (what_type_combobox.SelectedItem == null)
            {
                col = 0;
            }
            else
            {
                foreach (Excel.Range cell in ws.Range[ws.Cells[1, 1], ws.Cells[1, coln]])
                {
                    string type_selected = which_field_combobox.Text;
                    string currentCellValue = cell.Value?.ToString();
                    if ( currentCellValue == type_selected)
                    {
                        col = cell.Column;
                    }
                    else if (type_selected.Length==3 && type_selected.Substring(0,1)=="列" && type_selected.Substring(2,1)=="空")
                    {
                        col = int.Parse(type_selected.Substring(1, 1));
                    }
                }
            }
            string regex_type = what_type_combobox.Text;
            string pat = null;

            //选择已定义的正则表达式过滤条件，或自行写入过滤规则
            switch (regex_type)
            {
                case "数字":
                    pat = "\\d+\\.?\\d*";
                    ws.Range[ws.Cells[1, coln + 1], ws.Cells[rown, coln + 1]].NumberFormatLocal = "@";
                    break;
                case "英文":
                    pat = "[A-Za-z]+";
                    ws.Range[ws.Cells[1, coln + 1], ws.Cells[rown, coln + 1]].NumberFormatLocal = "@";
                    break;
                case "中文":
                    pat = "[^\\x00-\\xff]+";
                    ws.Range[ws.Cells[1, coln + 1], ws.Cells[rown, coln + 1]].NumberFormatLocal = "@";
                    break;
                case "网址":
                    pat = "((http|https):\\/\\/)?[\\w-]+(\\.[\\w-]+)+([\\w.,@?^=%&amp;:/~+#-]*[\\w@?^=%&amp;/~+#-])?";
                    ws.Range[ws.Cells[1, coln + 1], ws.Cells[rown, coln + 1]].NumberFormatLocal = "@";
                    break;
                case "身份证号":
                    pat = "\\d{15}$|\\d{17}([0-9]|X|x)";
                    ws.Range[ws.Cells[1, coln + 1], ws.Cells[rown, coln + 1]].NumberFormatLocal = "@";
                    break;
                case "电子邮箱":
                    pat = "\\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Z|a-z]{2,}\\b";
                    ws.Range[ws.Cells[1, coln + 1], ws.Cells[rown, coln + 1]].NumberFormatLocal = "@";
                    break;
                case "电话号码":
                    pat = "(?:(?:\\+|00)86)?1[3-9]\\d{9}|(?:0[1-9]\\d{1,2}-)?\\d{7,8}";
                    ws.Range[ws.Cells[1, coln + 1], ws.Cells[rown, coln + 1]].NumberFormatLocal = "@";
                    break;
                case "IP地址":
                    pat = "\\b\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\.\\d{1,3}\\b";
                    ws.Range[ws.Cells[1, coln + 1], ws.Cells[rown, coln + 1]].NumberFormatLocal = "@";
                    break;
                case "自定义":
                    if (string.IsNullOrEmpty(regex_rule_textbox.Text))
                    {
                        MessageBox.Show("请输入正则表达式过滤规则");
                        return;
                    }
                    else
                    {
                        pat = regex_rule_textbox.Text;
                    }
                    break;
            }

            if (col < coln + 1 && col > 0)
            {
                ws.Range[ws.Cells[1, col], ws.Cells[rown, col]].Select();
                Regex rgx = new Regex(pat);
                List<string> matchValue = new List<string>();
                foreach (Excel.Range tempLoopVar_rng in ThisAddIn.app.Selection)
                {
                    string tempLoopVar_rngValue=tempLoopVar_rng.Value?.ToString();
                    if (!string.IsNullOrEmpty(tempLoopVar_rngValue))
                    {
                        matchValue.Clear();
                        foreach (Match match in rgx.Matches(System.Convert.ToString(tempLoopVar_rng.Value)))
                        {
                            matchValue.Add(match.Value);
                        }
                        string result = string.Join("|", matchValue);
                        ThisAddIn.app.Cells[tempLoopVar_rng.Row, coln + 1] = result;
                    }
                }
                ShowLabel(run_result_label, true, "提取完毕");
                StartTimer();
            }
            else
            {
                MessageBox.Show("您输入的列数有误，请确认");
            }

            tabControl1.Enabled = true;
            move_sheet_button.Enabled = true;
            add_sheet_button.Enabled = true;
            transposition_button.Enabled = true;
            regex_button.Enabled = true;
            payslip_button.Enabled = true;
            contents_button.Enabled = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }

        //正则表达式清空选项
        private void regex_clear_button_Click(object sender, EventArgs e)
        {
            which_field_combobox.Text = null;
            what_type_combobox.Text = null;
            regex_rule_textbox.Text = null;
        }



        //一键生成工资条
        private void payslip_button_Click(object sender, EventArgs e)
        {
            //右侧功能区初始化
            function_title_label.Text = "工资表转换为工资条格式";
            which_field_label.Visible = false;
            which_field_combobox.Visible = false;
            what_type_label.Visible = false;
            what_type_combobox.Visible = false;
            regex_rule_label.Visible = false;
            regex_rule_textbox.Visible = false;
            run_result_label.Visible = false;
            regex_run_button.Visible = false;
            regex_clear_button.Visible = false;

            //左侧按钮状态改变
            tabControl1.Enabled = false;
            move_sheet_button.Enabled = false;
            add_sheet_button.Enabled = false;
            transposition_button.Enabled = false;
            regex_button.Enabled = false;
            payslip_button.Enabled = false;
            contents_button.Enabled = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            long used_range_row = ThisAddIn.app.ActiveSheet.UsedRange.rows.count;
            Excel.Range range = workbook.ActiveSheet.UsedRange;
            range.Select();
            ThisAddIn.app.Selection.Copy();
            Excel.Worksheet worksheet = workbook.Worksheets.Add(Before: workbook.ActiveSheet);
            worksheet.Name = "工资条";
            Excel.Worksheet new_worksheet = workbook.Worksheets["工资条"];
            new_worksheet.Activate();
            new_worksheet.Range["A1"].PasteSpecial(Excel.XlPasteType.xlPasteAll);
            for (long n = used_range_row; n >= 3; n--)
            {
                ThisAddIn.app.ActiveSheet.Rows(1).Copy();
                ThisAddIn.app.ActiveSheet.Rows(n).Select();
                ThisAddIn.app.Selection.Insert(Excel.XlDirection.xlDown);
                //Globals.ThisAddIn.Application.CutCopyMode = false;
                ThisAddIn.app.Selection.Insert(Excel.XlDirection.xlDown);
            }
            workbook.Worksheets["工资条"].Activate();
            workbook.ActiveSheet.Range["A1"].Select();
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;

            ShowLabel(run_result_label, true, "工资条转换完毕");
            StartTimer();

            //左侧按钮状态改变
            tabControl1.Enabled = true;
            move_sheet_button.Enabled = true;
            add_sheet_button.Enabled = true;
            transposition_button.Enabled = true;
            regex_button.Enabled = true;
            payslip_button.Enabled = true;
            contents_button.Enabled = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }



        //一键根据目录页建立新表
        private void contents_button_Click(object sender, EventArgs e)
        {
            //右侧功能区初始化
            function_title_label.Text = "根据目录页新建空白表";
            which_field_label.Visible = false;
            which_field_combobox.Visible = false;
            what_type_label.Visible = false;
            what_type_combobox.Visible = false;
            regex_rule_label.Visible = false;
            regex_rule_textbox.Visible = false;
            run_result_label.Visible = false;
            regex_run_button.Visible = false;
            regex_clear_button.Visible = false;

            //左侧按钮状态改变
            tabControl1.Enabled = false;
            move_sheet_button.Enabled = false;
            add_sheet_button.Enabled = false;
            transposition_button.Enabled = false;
            regex_button.Enabled = false;
            payslip_button.Enabled = false;
            contents_button.Enabled = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            //Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            MessageBox.Show("该选项是将只包含一个命名为‘目录’Sheet的excel文件自动生成各页空白表的功能,请确定该文件中只包含一个Sheet且已改名为‘目录’，同时各目录项从第二行开始");
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                string sheet_name = worksheet.Name;
                if (sheet_name == "目录")
                {
                    long row_count = worksheet.Rows.Count;
                    long used_row_count = worksheet.UsedRange.Rows.Count;
                    for (var i = 2; i <= used_row_count; i++)
                    {
                        worksheet.Activate();
                        string add_sheet_name = System.Convert.ToString(workbook.Worksheets["目录"].Cells(i, 1).Value);
                        Excel.Worksheet add_sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                        add_sheet.Name = add_sheet_name;
                        worksheet.Activate();
                        worksheet.Hyperlinks.Add(worksheet.Cells[i, 1], "", Convert.ToString(worksheet.Cells[i, 1].value) + "!A1", Convert.ToString(worksheet.Cells[i, 1].value));
                        worksheet.Cells[i, 1].Font.Name = "微软雅黑";
                        worksheet.Cells[i, 1].Font.Size = 12;
                        worksheet.Cells[i, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                        worksheet.Cells[i, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    }
                    workbook.Worksheets["目录"].Activate();
                    ThisAddIn.app.ActiveSheet.Range["A1"].Font.Name = "微软雅黑";
                    ThisAddIn.app.ActiveSheet.Range["A1"].Font.Size = 12;
                    ThisAddIn.app.ActiveSheet.Range["A1"].Font.Bold = true;
                    ThisAddIn.app.ActiveSheet.Range["A1"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    ThisAddIn.app.ActiveSheet.Range["A1"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                    ThisAddIn.app.DisplayAlerts = true;
                    ThisAddIn.app.ScreenUpdating = true;

                    ShowLabel(run_result_label, true, "根据目录页建新表完成");
                    StartTimer();
                    return;
                }
            }


            //左侧按钮状态改变
            tabControl1.Enabled = true;
            move_sheet_button.Enabled = true;
            add_sheet_button.Enabled = true;
            transposition_button.Enabled = true;
            regex_button.Enabled = true;
            payslip_button.Enabled = true;
            contents_button.Enabled = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
            MessageBox.Show("未包含命名为'目录'的表格");
        }


        //分表时间控件，显示分表运行时间
        private void split_sheet_timer_Tick(object sender, EventArgs e)
        {
            int hour = used_time_count / 3600;
            int min = (used_time_count % 3600) / 60;
            int sec = (used_time_count % 3600) % 60;
            if (res == true)
            {
                if (split_sheet_result_label.Visible == false)
                {
                    split_sheet_result_label.Visible = true;
                }
                split_sheet_timer.Enabled = false;
                if (hour > 0)
                {
                    split_sheet_result_label.Text = "分表完成，共用" + Convert.ToString(hour) + "时" + Convert.ToString(min) + "分" + Convert.ToString(sec) + "秒";
                }
                else if (min > 0)
                {
                    split_sheet_result_label.Text = "分表完成，共用" + Convert.ToString(min) + "分" + Convert.ToString(sec) + "秒";
                }
                else
                {
                    split_sheet_result_label.Text = "分表完成，共用" + Convert.ToString(sec) + "秒";
                }
                if (thread.IsAlive == true)
                {
                    thread.Abort();
                }
                used_time_count = 0;
                this.TopMost = false;
            }
            else
            {
                if (split_sheet_result_label.Visible == false)
                {
                    split_sheet_result_label.Visible = true;
                }
                used_time_count++;
                if (hour > 0)
                {
                    split_sheet_result_label.Text = "请勿退出工具，分表中......,已用时" + Convert.ToString(hour) + "小时" + Convert.ToString(min) + "分" + Convert.ToString(sec) + "秒";
                }
                else if (min > 0)
                {
                    split_sheet_result_label.Text = "请勿退出工具，分表中......,已用时" + Convert.ToString(min) + "分" + System.Convert.ToString(sec) + "秒";
                }
                else
                {
                    split_sheet_result_label.Text = "请勿退出工具，分表中......,已用时" + System.Convert.ToString(sec) + "秒";
                }
            }
        }


        //并表时间控件，显示并表运行时间
        private void merge_sheet_timer_Tick(object sender, EventArgs e)
        {
            int hour = used_time_count / 3600;
            int min = (used_time_count % 3600) / 60;
            int sec = (used_time_count % 3600) % 60;
            if (res == true)
            {
                if (merge_sheet_result_label.Visible == false)
                {
                    merge_sheet_result_label.Visible = true;
                }
                merge_sheet_timer.Enabled = false;
                if (hour > 0)
                {
                    merge_sheet_result_label.Text = "并表完成，共用" + Convert.ToString(hour) + "时" + Convert.ToString(min) + "分" + Convert.ToString(sec) + "秒";
                }
                else if (min > 0)
                {
                    merge_sheet_result_label.Text = "并表完成，共用" + Convert.ToString(min) + "分" + Convert.ToString(sec) + "秒";
                }
                else
                {
                    merge_sheet_result_label.Text = "并表完成，共用" + Convert.ToString(sec) + "秒";
                }
                if (thread.IsAlive == true)
                {
                    thread.Abort();
                }
                used_time_count = 0;
                this.TopMost = false;
            }
            else
            {
                if (merge_sheet_result_label.Visible == false)
                {
                    merge_sheet_result_label.Visible = true;
                }
                used_time_count++;
                if (hour > 0)
                {
                    merge_sheet_result_label.Text = "请勿退出工具，分表合并中......,已用时" + Convert.ToString(hour) + "小时" + Convert.ToString(min) + "分" + Convert.ToString(sec) + "秒";
                }
                else if (min > 0)
                {
                    merge_sheet_result_label.Text = "请勿退出工具，分表合并中......,已用时" + Convert.ToString(min) + "分" + System.Convert.ToString(sec) + "秒";
                }
                else
                {
                    merge_sheet_result_label.Text = "请勿退出工具，分表合并中......,已用时" + System.Convert.ToString(sec) + "秒";
                }
            }
        }




        //时间控件，控制完成提示标签显示5秒后消失
        private System.Timers.Timer aTimer = new System.Timers.Timer();
        private delegate void SafeCallDelegate(Label label, bool Visible, string Text);

        private void ShowLabel(Label label, bool Visible, string Text)
        {
            if (label.InvokeRequired)
            {
                SafeCallDelegate d = new SafeCallDelegate(ShowLabel);
                label.Invoke(d, new object[] { label, Visible, Text });
            }
            else
            {
                label.Visible = Visible;
                label.Text = Text;
            }
        }

        private void HideLabel(Label label, bool Visible, string Text)
        {
            if (label.InvokeRequired)
            {
                SafeCallDelegate d = new SafeCallDelegate(HideLabel);
                label.Invoke(d, new object[] { label, Visible, Text });
            }
            else
            {
                label.Visible = Visible;
                label.Text = Text;
            }
        }

        private void StartTimer()
        {
            aTimer.Interval = 5000; //5 seconds
            aTimer.Elapsed += OnTimedEvent;
            aTimer.Enabled = true;
        }

        private void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            HideLabel(run_result_label, false, "");
        }

        //自建函数

        //正则表达式函数，判断输入字符是否合规，如有不合规字符，返回true，否则返回false
        public static bool ContainsSpecialChars(string str, string reg_rule)
        {
            Regex reg1 = new Regex(reg_rule);
            return reg1.IsMatch(str);
        }

        private void multi_merge_sheet_checkBox_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void function_title_label_Click(object sender, EventArgs e)
        {

        }
    }
}
