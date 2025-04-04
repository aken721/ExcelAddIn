﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using MySql.Data.MySqlClient;
using Npgsql;
using System.Data.SQLite;
using Oracle.ManagedDataAccess.Client;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;
using System.Configuration;
using ZXing;
using ZXing.QrCode;
using Excel = Microsoft.Office.Interop.Excel;
using ZXing.Rendering;
using System.Collections;
using Microsoft.Office.Core;
using System.Diagnostics;
using Sdcb.WordClouds;
using SkiaSharp;
using System.Drawing.Text;
using System.Net.Http;



namespace ExcelAddIn
{
    public partial class Form1 : Form
    {
        private Excel.Workbook workbook;
        private string sheetindex;
        private string excelFilePath;
        private Int32 used_time_count = 0;
        private bool res = false;
        private Thread thread;
        private List<string> activeWorkBook_sheet_names = new List<string>();     //当前工作簿中所有工作表名称，初始化时首次写入
        private List<string> new_sheet_names = new List<string>();       //打开后新建表所有表名

        public Form1()
        {
            InitializeComponent();
            this.FormClosing += (s, e) => CleanupTempFiles();
        }

        //窗体初始化
        private void Form1_Load(object sender, EventArgs e)
        {
            this.MaximizeBox = false;
            //初始化tabcontrol控件
            tabControl1.SelectTab(0);
            workbook = ThisAddIn.app.ActiveWorkbook;      //变量workbook指定为当前打开的工作簿
            sheetindex = ThisAddIn.app.ActiveSheet.Name;
            excelFilePath = workbook.FullName;
            sheet_name_combobox.Items.Clear();
            field_name_combobox.Items.Clear();
            data_start_combobox.SelectedIndex = 1;
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
            version_label.Text = ConfigurationManager.AppSettings["version"].ToString();

            if (activeWorkBook_sheet_names.Count > 0)
            {
                activeWorkBook_sheet_names.Clear();
            }
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                activeWorkBook_sheet_names.Add(sheet.Name);
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
            contents_to_sheet_radioButton.Visible = false;
            sheet_to_contents_radioButton.Visible = false;
            QR_label.Visible = false;
            QR_listBox.Visible = false;
            QR_radioButton.Visible = false;
            QR_radioButton.Checked = false;
            foreColor_select_button.Visible = false;
            foreColor_label.Visible = false;
            backColor_select_button.Visible = false;
            backColor_label.Visible = false;
            QR_logo_label.Visible = false;
            QR_logo_pictureBox.Visible = false;
            BC_radioButton.Visible = false;
            BC_radioButton.Checked = false;

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
                //分表
                case 0:
                    sheet_name_combobox.Items.Clear();
                    field_name_combobox.Items.Clear();
                    data_start_combobox.SelectedIndex = 1;
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

                //并表
                case 1:
                    merge_sheet_result_label.Visible = false;
                    merge_sheet_result_label.Text = "";
                    mergeProgressBar_label.Visible = false;
                    mergeProgressBar_label.Text = "";
                    merge_sheet_progressBar.Visible = false;
                    break;

                //批量导、删表
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

                //实用功能汇总
                case 3:
                    which_field_label.Visible = false;
                    which_field_combobox.Visible = false;
                    what_type_label.Visible = false;
                    what_type_combobox.Visible = false;
                    regex_rule_label.Visible = false;
                    regex_rule_textbox.Visible = false;
                    run_result_label.Visible = false;
                    regex_run_button.Visible = false;
                    regex_clear_button.Visible = false;
                    contents_to_sheet_radioButton.Visible = false;
                    sheet_to_contents_radioButton.Visible = false;
                    QR_label.Visible = false;
                    QR_listBox.Visible = false;
                    QR_radioButton.Visible = false;
                    QR_radioButton.Checked = false;
                    foreColor_select_button.Visible = false;
                    foreColor_label.Visible = false;
                    backColor_select_button.Visible = false;
                    backColor_label.Visible = false;
                    QR_logo_label.Visible = false;
                    QR_logo_pictureBox.Visible = false;
                    QR_logo_pictureBox.Image = ExcelAddIn.Properties.Resources.pic_logo;
                    BC_radioButton.Visible = false;
                    BC_radioButton.Checked = false;

                    function_title_label.Text = "请选择所需使用的功能";
                    break;

                //数据库表提取
                case 4:
                    database_result_label.Text = string.Empty;
                    dbsheet_comboBox.Items.Clear();
                    dbexport_result_label.Text = string.Empty;
                    dbexport_result_label.Text = string.Empty;
                    find_keywordclear_pictureBox.Visible = false;
                    break;

                //图表增强
                case 5:
                    this.TopMost = false;
                    chart_select_comboBox.SelectedIndex = 0;
                    LoadFontsToComboBox(comboBoxFonts);
                    comboBoxTextDirection.SelectedIndex = 0;
                    //目前测试词云图中字体设定功能可能对中文不起作用，暂时隐藏
                    labelFonts.Visible = false;                  
                    comboBoxFonts.Visible = false;
                    comboBoxFonts.Text = "微软雅黑";             
                    break;

                //帮助
                case 6:
                    break;

                //退出
                case 7:
                    this.Dispose();
                    break;
            }
        }

        // 定义用于存储字体信息的类
        public class FontInfo
        {
            public string DisplayName { get; set; }
            public string EnglishName { get; set; }
        }

        //在一个文本选择框中添加本机字体库中字体名称
        private void LoadFontsToComboBox(ComboBox comboBox)
        {
            // 获取系统安装的字体
            InstalledFontCollection fonts = new InstalledFontCollection();
            foreach (FontFamily font in fonts.Families)
            {
                // 获取英文名称（LCID 1033对应英文）
                string englishName = font.GetName(1033);
                // 如果英文名为空，则使用字体家族名
                if (string.IsNullOrEmpty(englishName))
                    englishName = font.Name;

                comboBoxFonts.Items.Add(new FontInfo
                {
                    DisplayName = font.Name,
                    EnglishName = englishName
                });
            }

            // 设置显示属性为DisplayName
            comboBoxFonts.DisplayMember = "DisplayName";
            comboBoxFonts.Sorted = true;
        }

        //限制指定数据起始行的ComboBox中只能输入数字
        private void data_start_combobox_KeyPress(object sender, KeyPressEventArgs e)
        {
            // 如果输入的字符不是数字，也不是控制字符（如退格键），则阻止输入
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void data_start_combobox_TextChanged(object sender, EventArgs e)
        {
            this.sheet_name_combobox_TextChanged(sender, e);
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
                int title_last_row = int.Parse(data_start_combobox.Text) - 1;
                if (title_last_row == 0)
                {
                    ShowLabel(split_sheet_result_label, true, "该表没有标题行！");
                    StartTimer();
                    for (int i = 1; i <= worksheet.UsedRange.Columns.Count; i++)
                    {
                        field_name_combobox.Items.Add(Convert.ToString(i));
                    }
                }
                else
                {
                    foreach (Excel.Range range in worksheet.Range[worksheet.Cells[title_last_row, 1], worksheet.Cells[title_last_row, worksheet.UsedRange.Columns.Count]])
                    {
                        string range_value = range.Value;
                        if (string.IsNullOrEmpty(range_value))
                        {
                            if (range.MergeCells)
                            {
                                Excel.Range merge_range = range.MergeArea;
                                string merge_range_value = ThisAddIn.app.ActiveSheet.Range[merge_range.Address.Split(':')[0]].Value;
                                if (string.IsNullOrEmpty(merge_range_value))
                                {
                                    field_name_combobox.Items.Add(range.Column.ToString());
                                }
                                else
                                {
                                    field_name_combobox.Items.Add(range.Column.ToString() + "." + merge_range_value);
                                }

                            }
                            else
                            {
                                field_name_combobox.Items.Add(range.Column.ToString());
                            }
                        }
                        else
                        {
                            field_name_combobox.Items.Add(range.Column.ToString() + "." + range.Value);
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
            }
            field_name_combobox.Refresh();
        }



        //分表功能中清空combobox内容
        private void clear_button_Click(object sender, EventArgs e)
        {
            sheet_name_combobox.Text = "";
            field_name_combobox.Text = "";
            data_start_combobox.Text = "2";
        }

        //分表（UI主线程）
        private void split_button_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(sheet_name_combobox.Text) && string.IsNullOrEmpty(field_name_combobox.Text) && string.IsNullOrEmpty(data_start_combobox.Text))
            {
                ShowLabel(split_sheet_result_label, true, "表、字段和数据起始行均不能为空！");
                StartTimer();
                return;
            }
            int field_column = 0;
            string field_name_selected = field_name_combobox.Text;

            if (field_name_selected.Contains("."))
            {
                field_column = int.Parse(field_name_selected.Split('.')[0]);
            }
            else
            {
                field_column = int.Parse(field_name_selected);
            }
            string select_sheet = sheet_name_combobox.Text;
            int dataStartRow = int.Parse(data_start_combobox.Text);
            thread = new Thread(() => SplitTask(select_sheet, field_column, dataStartRow));
            thread.Start();
            split_sheet_result_label.Visible = true;
            split_sheet_timer.Interval = 1000;
            split_sheet_timer.Enabled = true;
            splitProgressBar_label.Visible = true;
            split_sheet_progressBar.Visible = true;
        }

        //分表（程序执行线程）
        private void SplitTask(string sheetName, int selectFieldsColumn, int selectDataStartRow)
        {
            res = false;
            this.Invoke(new Action(() =>
            {
                tabControl1.Enabled = false;
                split_button.Enabled = false;
                splitsheet_export_button.Enabled = false;
                splitsheet_delete_button.Enabled = false;
                clear_button.Enabled = false;
                sheet_name_combobox.Enabled = false;
                field_name_combobox.Enabled = false;
                this.ControlBox = false;
            }));
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            try
            {
                //声明范围列数、范围行数、分表依据列数、筛选结果第一列数
                List<string> records = new List<string>();
                long record_row = workbook.Worksheets[sheetName].Cells[workbook.Worksheets[sheetName].Rows.Count, selectFieldsColumn].End(Excel.XlDirection.xlUp).Row;  //待分表中的数据行数
                int current_record = 1;

                //将去重后的表名加入数组
                foreach (Excel.Range range in workbook.Worksheets[sheetName].Range[workbook.Worksheets[sheetName].Cells[selectDataStartRow, selectFieldsColumn], workbook.Worksheets[sheetName].Cells[record_row, selectFieldsColumn]])
                {
                    if (records.Contains(range.Value) || string.IsNullOrEmpty(range.Value))
                    {
                        continue;
                    }
                    else
                    {
                        var rangeValue = range.Value;
                        records.Add(Convert.ToString(rangeValue));
                    }
                }
                int total_record = records.Count;

                //动态更新一个分表工作簿中所有表的名称
                List<string> dynamic_sheet_name = new List<string>();


                //新建分表，并通过关键字段筛选，筛出结果复制到相应分表中
                foreach (string record in records)
                {
                    //更新进度条
                    UpdateProgressBar(split_sheet_progressBar, current_record, total_record, splitProgressBar_label, "分表进度");

                    //未分表前工作簿已有表登记
                    if (dynamic_sheet_name.Count > 0)
                    {
                        dynamic_sheet_name.Clear();
                    }
                    foreach (Excel.Worksheet ws in workbook.Worksheets)
                    {
                        dynamic_sheet_name.Add(ws.Name);
                    }

                    //新建名为record的分表
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
                    workbook.Worksheets[sheetName].Select();

                    //定义筛选范围，前边已定义record_row(标题+数据总行数)、selectDataStartRow(数据起始行)、selectFieldsColumn(筛选关键字所在列)
                    int record_column = workbook.Worksheets[sheetName].UsedRange.Columns.Count;
                    int sheet_allRows = workbook.Worksheets[sheetName].Rows.Count;
                    Excel.Worksheet activeSheet = ThisAddIn.app.ActiveSheet;

                    activeSheet.Rows[selectDataStartRow].Insert(Shift: Excel.XlInsertShiftDirection.xlShiftDown, CopyOrigin: Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                    activeSheet.Range[activeSheet.Cells[selectDataStartRow, 1], activeSheet.Cells[record_row, record_column]].Select();
                    ThisAddIn.app.Selection.AutoFilter(selectFieldsColumn, record);
                    int autofilter_row = activeSheet.Cells[sheet_allRows, selectFieldsColumn].End(Excel.XlDirection.xlUp).Row;
                    activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[autofilter_row, record_column]].Select();
                    ThisAddIn.app.Selection.Copy(ThisAddIn.app.ActiveWorkbook.Worksheets[record].Range["A1"]);
                    activeSheet.Rows[selectDataStartRow].AutoFilter();
                    activeSheet.Rows[selectDataStartRow].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);



                    //对有序号列的表数据重新排序
                    add_sheet.Select();
                    add_sheet.Rows[selectDataStartRow].Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                    foreach (Excel.Range rng in ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells(1, 1), ThisAddIn.app.ActiveSheet.Cells(selectDataStartRow - 1, ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count)])
                    {
                        if (rng.Value == "序号")
                        {
                            int t_column = rng.Column;                                                                             //“序号”所在的列
                            int data_row = add_sheet.Cells[add_sheet.Rows.Count, record_column].End(Excel.XlDirection.xlUp).Row;   //分后的表中最后一条数据所在行
                            for (int number = 0; number <= data_row - selectDataStartRow; number++)
                            {
                                ThisAddIn.app.ActiveSheet.Cells[selectDataStartRow + number, t_column].Value = number + 1;
                            }
                            break;
                        }
                    }
                    add_sheet.Range[add_sheet.Cells[1, 1], add_sheet.Cells[1, record_column]].EntireColumn.AutoFit();
                    add_sheet.Range["A1"].Select();
                    current_record++;
                }
                workbook.Worksheets[sheetName].Activate();
                ThisAddIn.app.ActiveSheet.Range("A1").Select();
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
                    if (!activeWorkBook_sheet_names.Contains(newsheet.Name))
                    {
                        new_sheet_names.Add(newsheet.Name);
                    }
                }
                this.Invoke(new Action(() =>
                {
                    tabControl1.Enabled = true;
                    split_button.Enabled = true;
                    splitsheet_export_button.Enabled = true;
                    splitsheet_delete_button.Enabled = true;
                    clear_button.Enabled = true;
                    sheet_name_combobox.Enabled = true;
                    field_name_combobox.Enabled = true;
                    this.ControlBox = true;
                    this.TopMost = false;
                }));
                res = true;
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
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
            if (activeWorkBook_sheet_names.Count > 0)
            {
                activeWorkBook_sheet_names.Clear();
            }
            foreach (Excel.Worksheet active_sheet_name in workbook.Worksheets)
            {
                activeWorkBook_sheet_names.Add(active_sheet_name.Name);
            }

            //启动并表线程
            thread = new Thread(() => mergeTask(data_start_row, activeWorkBook_sheet_names));
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
            if (activeWorkBook_sheet_names.Count > 0)
            {
                activeWorkBook_sheet_names.Clear();
            }
            foreach (Excel.Worksheet active_sheet_name in workbook.Worksheets)
            {
                activeWorkBook_sheet_names.Add(active_sheet_name.Name);
            }

            //启动并表线程
            thread = new Thread(() => multiMergeTask(data_start_row, activeWorkBook_sheet_names));
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
            this.Invoke(new Action(() =>
            {
                progressBar.Value = progressPercentage;
                progressBar.Update();
                // 显示百分比数字
                progressBar_result_label.Text = progressBar_result + progressPercentage.ToString() + "%";
            }));
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
            contents_to_sheet_radioButton.Visible = false;
            sheet_to_contents_radioButton.Visible = false;
            function_title_label.Text = "不同工作簿中的表全部复制到本工作簿";
            QR_label.Visible = false;
            QR_listBox.Visible = false;
            QR_radioButton.Visible = false;
            QR_radioButton.Checked = false;
            foreColor_select_button.Visible = false;
            foreColor_label.Visible = false;
            backColor_select_button.Visible = false;
            backColor_label.Visible = false;
            QR_logo_label.Visible = false;
            QR_logo_pictureBox.Visible = false;
            BC_radioButton.Visible = false;
            BC_radioButton.Checked = false;
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
            QR_button.Enabled = false;

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
            QR_button.Enabled = true;
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
            QR_label.Visible = false;
            QR_listBox.Visible = false;
            QR_radioButton.Visible = false;
            QR_radioButton.Checked = false;
            foreColor_select_button.Visible = false;
            foreColor_label.Visible = false;
            backColor_select_button.Visible = false;
            backColor_label.Visible = false;
            QR_logo_label.Visible = false;
            QR_logo_pictureBox.Visible = false;
            BC_radioButton.Visible = false;
            BC_radioButton.Checked = false;

            contents_to_sheet_radioButton.Visible = false;
            sheet_to_contents_radioButton.Visible = false;

            //左侧按钮状态改变
            tabControl1.Enabled = false;
            move_sheet_button.Enabled = false;
            add_sheet_button.Enabled = false;
            transposition_button.Enabled = false;
            regex_button.Enabled = false;
            payslip_button.Enabled = false;
            contents_button.Enabled = false;
            QR_button.Enabled = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;
            try
            {
                string shtname = "";
                int i = 0;
                ThisAddIn.app.DisplayAlerts = false;
                ThisAddIn.app.ScreenUpdating = false;
                string activated_sheet_name = ThisAddIn.app.ActiveSheet.Name;
                int n = Convert.ToInt32(ThisAddIn.app.InputBox("请输入需要新建空表数量（最多15张）：", "输入建表数量"));
                if (n > 0 && n <= 15)
                {
                    shtname = Convert.ToString(ThisAddIn.app.InputBox("请输入表统一名称,未输入则缺省命名为‘新建表’：", "输入表名称"));
                    if (shtname == "False")
                    {
                        ShowLabel(run_result_label, true, "取消表名输入");
                        StartTimer();
                    }
                    else
                    {
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

                }
                else
                {
                    ShowLabel(run_result_label, true, "未正确输入新建表数量");
                    StartTimer();
                }
            }
            catch (Exception ex)
            {
                ShowLabel(run_result_label, true, ex.Message);
            }
            finally
            {
                //左侧按钮状态改变
                tabControl1.Enabled = true;
                move_sheet_button.Enabled = true;
                add_sheet_button.Enabled = true;
                transposition_button.Enabled = true;
                regex_button.Enabled = true;
                payslip_button.Enabled = true;
                contents_button.Enabled = true;
                QR_button.Enabled = true;
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
            }
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
            QR_label.Visible = false;
            QR_listBox.Visible = false;
            QR_radioButton.Visible = false;
            QR_radioButton.Checked = false;
            foreColor_select_button.Visible = false;
            foreColor_label.Visible = false;
            backColor_select_button.Visible = false;
            backColor_label.Visible = false;
            QR_logo_label.Visible = false;
            QR_logo_pictureBox.Visible = false;
            BC_radioButton.Visible = false;
            BC_radioButton.Checked = false;
            contents_to_sheet_radioButton.Visible = false;
            sheet_to_contents_radioButton.Visible = false;

            //左侧按钮状态改变
            tabControl1.Enabled = false;
            move_sheet_button.Enabled = false;
            add_sheet_button.Enabled = false;
            transposition_button.Enabled = false;
            regex_button.Enabled = false;
            payslip_button.Enabled = false;
            contents_button.Enabled = false;
            QR_button.Enabled = false;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            try
            {

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
                        case -2:
                            ShowLabel(run_result_label, true, "程序运行出现错误，可能转置表已存在，出现表名冲突");
                            StartTimer();
                            break;
                    }
                }, TaskScheduler.FromCurrentSynchronizationContext());
            }
            catch (Exception ex)
            {
                ShowLabel(run_result_label, true, ex.Message);
                StartTimer();
            }
            finally
            {
                //左侧按钮状态改变
                tabControl1.Enabled = true;
                move_sheet_button.Enabled = true;
                add_sheet_button.Enabled = true;
                transposition_button.Enabled = true;
                regex_button.Enabled = true;
                payslip_button.Enabled = true;
                contents_button.Enabled = true;
                QR_button.Enabled = true;
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;
            }
        }

        //转置工作表（程序执行线程）
        private int transTask()
        {
            try
            {
                Excel.Worksheet worksheet = ThisAddIn.app.ActiveSheet;
                //获取当前表名称
                string active_sheet_name = worksheet.Name;
                //获取当前表全部行数
                long row_count = worksheet.Rows.Count;
                //获取当前表全部列数
                //long column_count = worksheet.Columns.Count;
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
                string field_name = Convert.ToString(ThisAddIn.app.InputBox("请输入转置列的字段名称（若是日期可命名为“日期”或“date”，会自动格式化转置数据）：", "注意"));
                if (string.IsNullOrEmpty(field_name) || field_name == "False")
                {
                    return 1;
                }

                //开始转置
                if (run_result_label.InvokeRequired)
                {
                    run_result_label.Invoke(new Action(() =>
                        {
                            run_result_label.Visible = true;
                            run_result_label.Text = "正在转置......";
                        }));
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
                trans_sheet.Cells[1, Convert.ToInt32(translation_start_column)] = "value";
                trans_sheet.Cells[1, translation_start_column1 + 1] = field_name;


                //日期数据格式化
                if (field_name == "日期" || field_name == "date")
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

                //判断value列数据是否为数字，如果是数字，将转置表的该列统一格式化为“千位分隔符+小数点后保留2位”
                string pattern = @"^[-+]?[0-9]*\.?[0-9]+$";
                string cellValue = ThisAddIn.app.ActiveWorkbook.Worksheets[active_sheet_name].Cells[2, translation_start_column1].Value.ToString();
                if (ContainsSpecialChars(cellValue, pattern))
                {
                    trans_sheet.Columns[translation_start_column1].NumberFormatLocal = "#,##0.00";
                }


                worksheet.Select();
                worksheet.Range["A1"].Select();
                trans_sheet.Activate();
                trans_sheet.Range["A1"].Select();
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
                return 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("转置出错：" + ex.Message);
                return -2;
            }
            finally
            {
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
            }
        }

        //正则表达式功能激活
        private void regex_button_Click(object sender, EventArgs e)
        {

            selectfunction = 1;
            if (which_field_label.Visible == true && which_field_combobox.Visible == true)
            {
                which_field_combobox.Text = "";
                which_field_combobox.Items.Clear();
                which_field_label.Visible = false;
                which_field_combobox.Visible = false;
            }
            contents_to_sheet_radioButton.Checked = false;
            sheet_to_contents_radioButton.Checked = false;
            //workbook.Worksheets[sheetindex].Activate();
            workbook.RefreshAll();
            function_title_label.Text = "正则表达式提取指定内容";
            which_field_label.Text = "提取哪列";
            which_field_combobox.Text = "";
            which_field_label.Visible = true;
            which_field_combobox.Visible = true;
            what_type_label.Visible = true;
            what_type_combobox.Visible = true;
            regex_run_button.Visible = true;
            regex_clear_button.Visible = true;
            run_result_label.Visible = false;
            contents_to_sheet_radioButton.Visible = false;
            sheet_to_contents_radioButton.Visible = false;
        }

        //which_field_combobox控件可见时
        private async void which_field_combobox_VisibleChanged(object sender, EventArgs e)
        {
            if (which_field_combobox.Visible == true)
            {
                switch (selectfunction)
                {
                    case 1:
                        which_field_combobox.Items.Clear();
                        which_field_combobox.Text = "";
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
                        break;

                    case 2:
                        which_field_combobox.Items.Clear();
                        await Task.Run(() =>
                        {
                            if (SheetExist("目录"))
                            {
                                Excel.Worksheet contentSheet = workbook.Worksheets["目录"];
                                long usedrange_columns = contentSheet.UsedRange.Columns.Count;
                                foreach (Excel.Range cell in contentSheet.Range[contentSheet.Cells[1, 1], contentSheet.Cells[1, usedrange_columns]])
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
                            }
                        });
                        break;
                }
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

        private int selectfunction = 0;

        //右侧功能区运行按钮，包括正则表达式和目录页功能
        private void regex_run_button_Click(object sender, EventArgs e)
        {
            try
            {
                switch (selectfunction)
                {
                    case 0:
                        break;

                    /*该部分为按正则表达式功能模块
                    * 可按既定规则提取相应内容
                    * 也可自定义提取规则提取相应内容
                    */
                    case 1:
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
                                if (currentCellValue == type_selected)
                                {
                                    col = cell.Column;
                                }
                                else if (type_selected.Length == 3 && type_selected.Substring(0, 1) == "列" && type_selected.Substring(2, 1) == "空")
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
                                string tempLoopVar_rngValue = tempLoopVar_rng.Value?.ToString();
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
                        break;

                    /*该部分为按目录页创建（或链接）工作簿中表的功能
                     * 目录页只需表名设置为“目录”
                     * 目录页各字段可个性化设定，可按字段选择需链接列
                     */
                    case 2:

                        //左侧按钮状态改变
                        tabControl1.Enabled = false;
                        move_sheet_button.Enabled = false;
                        add_sheet_button.Enabled = false;
                        transposition_button.Enabled = false;
                        regex_button.Enabled = false;
                        payslip_button.Enabled = false;
                        contents_button.Enabled = false;

                        List<string> sheetsName = new List<string>();
                        foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                        {
                            sheetsName.Add(worksheet.Name);
                        }

                        if (sheetsName.Contains("目录"))
                        {
                            Excel.Worksheet contentsSheet = workbook.Worksheets["目录"];
                            int targetColumn = TargetField(contentsSheet, which_field_combobox.Text);
                            if (which_field_combobox.Text != "")
                            {
                                Task.Run(() =>
                                {
                                    ThisAddIn.app.ScreenUpdating = false;
                                    ThisAddIn.app.DisplayAlerts = false;
                                    long row_count = contentsSheet.Rows.Count;
                                    long used_row_count = contentsSheet.UsedRange.Rows.Count;
                                    for (var i = 2; i <= used_row_count; i++)
                                    {
                                        contentsSheet.Activate();
                                        string add_sheet_name = System.Convert.ToString(contentsSheet.Cells[i, targetColumn].Value);
                                        if (!sheetsName.Contains(add_sheet_name))
                                        {
                                            Excel.Worksheet add_sheet = workbook.Worksheets.Add(After: workbook.Worksheets[workbook.Worksheets.Count]);
                                            add_sheet.Name = add_sheet_name;
                                            contentsSheet.Activate();
                                        }
                                        contentsSheet.Hyperlinks.Add(contentsSheet.Cells[i, targetColumn], "", Convert.ToString(contentsSheet.Cells[i, targetColumn].value) + "!A1", Convert.ToString(contentsSheet.Cells[i, targetColumn].value));
                                        contentsSheet.Cells[i, targetColumn].Font.Name = "微软雅黑";
                                        contentsSheet.Cells[i, targetColumn].Font.Size = 12;
                                        contentsSheet.Cells[i, targetColumn].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                        contentsSheet.Cells[i, targetColumn].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                    }
                                    contentsSheet.Activate();
                                    ThisAddIn.app.ActiveSheet.UsedRange.Font.Name = "微软雅黑";
                                    ThisAddIn.app.ActiveSheet.UsedRange.Font.Size = 12;
                                    ThisAddIn.app.ActiveSheet.Range[ThisAddIn.app.ActiveSheet.Cells[1, 1], ThisAddIn.app.ActiveSheet.Cells[1, ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count]].Font.Bold = true;
                                    ThisAddIn.app.ActiveSheet.UsedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                    ThisAddIn.app.ActiveSheet.UsedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                                    ThisAddIn.app.DisplayAlerts = true;
                                    ThisAddIn.app.ScreenUpdating = true;

                                    // 更新界面
                                    this.Invoke((MethodInvoker)delegate
                                    {
                                        ShowLabel(run_result_label, true, "根据目录页建新表完成");
                                        StartTimer();
                                        tabControl1.Enabled = true;
                                        move_sheet_button.Enabled = true;
                                        add_sheet_button.Enabled = true;
                                        transposition_button.Enabled = true;
                                        regex_button.Enabled = true;
                                        payslip_button.Enabled = true;
                                        contents_button.Enabled = true;
                                        ThisAddIn.app.DisplayAlerts = true;
                                        ThisAddIn.app.ScreenUpdating = true;
                                    });
                                });
                            }
                        }
                        else
                        {
                            ShowLabel(run_result_label, true, "未包含命名为'目录'的表格");
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
                        ThisAddIn.app.DisplayAlerts = true;
                        ThisAddIn.app.ScreenUpdating = true;
                        break;

                    /*该功能为根据工作簿已有工作表建立带链接的目录页
                     *新建目录表命名为“_目录”，链接表内容位于表内“表目录”字段下
                     *建成后可对目录页加工添加其他内容，只要不破坏“表目录”字段内容即可
                     *如原工作簿已有“_目录”表，会自动更名为“_目录+当前日期时间字符串”的表
                     */
                    case 3:

                        //左侧按钮状态改变
                        tabControl1.Enabled = false;
                        move_sheet_button.Enabled = false;
                        add_sheet_button.Enabled = false;
                        transposition_button.Enabled = false;
                        regex_button.Enabled = false;
                        payslip_button.Enabled = false;
                        contents_button.Enabled = false;

                        List<string> shtsName = new List<string>();
                        foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                        {
                            shtsName.Add(worksheet.Name);
                        }
                        if (shtsName.Contains("_目录"))
                        {
                            Excel.Worksheet repeatingSheet = workbook.Worksheets["_目录"];
                            string dt = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString("00") + DateTime.Now.Day.ToString("00") + DateTime.Now.Hour.ToString("00") + DateTime.Now.Minute.ToString("00") + DateTime.Now.Second.ToString("00");
                            repeatingSheet.Name = "_目录" + dt;
                            int index = shtsName.IndexOf("_目录");
                            shtsName[index] = repeatingSheet.Name;
                        }
                        Excel.Worksheet addSheet = workbook.Worksheets.Add(Before: workbook.Worksheets[1]);
                        addSheet.Name = "_目录";
                        addSheet.Range["A1"].Value = "表目录";
                        Task.Run(() =>
                        {
                            ThisAddIn.app.ScreenUpdating = false;
                            ThisAddIn.app.DisplayAlerts = false;
                            for (int i = 0; i < shtsName.Count; i++)
                            {
                                addSheet.Cells[i + 2, 1].value = shtsName[i];
                                addSheet.Hyperlinks.Add(addSheet.Cells[i + 2, 1], "", Convert.ToString(addSheet.Cells[i + 2, 1].value) + "!A1", Convert.ToString(addSheet.Cells[i + 2, 1].value));
                                addSheet.Cells[i + 2, 1].Font.Name = "微软雅黑";
                                addSheet.Cells[i + 2, 1].Font.Size = 12;
                                addSheet.Cells[i + 2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                                addSheet.Cells[i + 2, 1].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            }
                            addSheet.Range["A1"].Font.Name = "微软雅黑";
                            addSheet.Range["A1"].Font.Bold = true;
                            addSheet.UsedRange.Font.Size = 12;
                            addSheet.UsedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                            addSheet.UsedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                            ThisAddIn.app.DisplayAlerts = true;
                            ThisAddIn.app.ScreenUpdating = true;

                            // 更新界面
                            this.Invoke((MethodInvoker)delegate
                            {
                                ShowLabel(run_result_label, true, "创建目录页完成");
                                StartTimer();
                                tabControl1.Enabled = true;
                                move_sheet_button.Enabled = true;
                                add_sheet_button.Enabled = true;
                                transposition_button.Enabled = true;
                                regex_button.Enabled = true;
                                payslip_button.Enabled = true;
                                contents_button.Enabled = true;
                                ThisAddIn.app.DisplayAlerts = true;
                                ThisAddIn.app.ScreenUpdating = true;
                            });
                        });
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
                        break;


                    /*该功能为将指定列内容的数据分别生成二维码或条形码
                     *指定列可以为一列，也可为多列
                     *可选择生成二维码或条形码
                     */
                    case 4:
                        // 计算灰度值
                        int foreGray = GetGrayScale(qrForeColor);
                        int backGray = GetGrayScale(qrBackColor);
                        // 计算对比值
                        int contrast = Math.Abs(foreGray - backGray);
                        if (contrast < 50)
                        {
                            MessageBox.Show("前景色与背景色差值过小，无法生成二维码！请重新选择前景与背景对比度差值较大的颜色");
                            return;
                        }

                        Excel.Worksheet sheet = workbook.ActiveSheet;

                        int usedColumn = sheet.UsedRange.Columns.Count;
                        Dictionary<string, int> items = new Dictionary<string, int>();
                        foreach (string selectitem in QR_listBox.SelectedItems)
                        {
                            items.Add(selectitem, TargetField(sheet, selectitem));
                        }

                        //生成二维码
                        if (QR_radioButton.Checked == true)
                        {
                            for (int i = 2; i <= sheet.UsedRange.Rows.Count; i++)
                            {
                                string data = "";

                                //如果二维码有多列，则将多列数据合并，以key:value的形式返回，用分号分割
                                if (items.Count > 1)
                                {
                                    foreach (var item in items)
                                    {
                                        string key = item.Key;
                                        int colindex = item.Value;
                                        string value = sheet.Cells[i, colindex].Text;
                                        data += $"{key}:{value};";
                                    }
                                }
                                //如果二维码只有一列，则直接取值
                                else
                                {
                                    string key = items.Keys.First();
                                    int colindex = items.Values.First();
                                    string value = sheet.Cells[i, colindex].Text;
                                    data = value;
                                }

                                // 创建新的 BarcodeWriter 实例
                                BarcodeWriter writer = new BarcodeWriter
                                {
                                    Format = BarcodeFormat.QR_CODE,
                                    Options = new QrCodeEncodingOptions
                                    {
                                        Height = 100,
                                        Width = 100,
                                        CharacterSet = "UTF-8",
                                        ErrorCorrection = ZXing.QrCode.Internal.ErrorCorrectionLevel.H,  // 设置纠错等级为H
                                        Margin = 0
                                    },
                                    Renderer = new BitmapRenderer
                                    {
                                        Foreground = qrForeColor, // 前景色
                                        Background = qrBackColor // 背景色
                                    }
                                };

                                byte[] utf8Bytes = Encoding.UTF8.GetBytes(data);
                                Bitmap qrCode = writer.Write(Encoding.UTF8.GetString(utf8Bytes));

                                // 如果提供了Logo图片路径，则在二维码中间添加Logo
                                if (!string.IsNullOrEmpty(qr_logo_path) && QR_logo_pictureBox.Image != ExcelAddIn.Properties.Resources.pic_logo)
                                {
                                    using (Bitmap logo = new Bitmap(qr_logo_path))
                                    {
                                        // 调整Logo大小为二维码的1/5
                                        int adjustedLogoWidth = (int)(qrCode.Width / 5);
                                        int adjustedLogoHeight = (int)(qrCode.Height / 5);
                                        using (Bitmap resizedLogo = new Bitmap(logo, new Size(adjustedLogoWidth, adjustedLogoHeight)))
                                        {
                                            using (Graphics g = Graphics.FromImage(qrCode))
                                            {
                                                // 计算Logo的位置
                                                float x = (qrCode.Width - adjustedLogoWidth) / 2;
                                                float y = (qrCode.Height - adjustedLogoHeight) / 2;

                                                // 绘制Logo
                                                g.DrawImage(resizedLogo, new RectangleF(x, y, adjustedLogoWidth, adjustedLogoHeight));
                                            }
                                        }
                                    }
                                }

                                string tempImagePath = System.IO.Path.GetTempFileName() + ".png";
                                qrCode.Save(tempImagePath, ImageFormat.Png);

                                Excel.Range cellForImage = sheet.Cells[i, usedColumn + 1];
                                cellForImage.Rows.RowHeight = 100; // 或根据实际需要调整行高
                                cellForImage.Columns.ColumnWidth = 15; // 或根据实际需要调整列宽

                                // 插入图片到单元格
                                Excel.Shape shape = sheet.Shapes.AddPicture(tempImagePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                                    cellForImage.Left, cellForImage.Top, -1, -1);
                                shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

                                File.Delete(tempImagePath); // 删除临时文件
                            }
                            ShowLabel(run_result_label, true, "所有二维码已生成");
                            StartTimer();
                        }
                        else if (BC_radioButton.Checked == true)                   //生成格式为Code28的条形码，其他格式暂不支持，有需要可自行修改代码支持
                        {
                            if (items.Count == 1)
                            {
                                BarcodeWriter writer = new BarcodeWriter
                                {
                                    Format = BarcodeFormat.CODE_128,
                                    Options = new QrCodeEncodingOptions
                                    {
                                        Height = 50,
                                        Width = 150,
                                        Margin = 1,
                                        PureBarcode = true,
                                        CharacterSet = "UTF-8"
                                    },
                                    Renderer = new BitmapRenderer
                                    {
                                        Foreground = qrForeColor,
                                        Background = Color.White
                                    }
                                };

                                int count = 0;
                                for (int i = 2; i <= sheet.UsedRange.Rows.Count; i++)    //默认第一行为标题行，第二行开始为数据行
                                {
                                    string value = sheet.Cells[i, items.Values.First()].Text;
                                    string asciiPattern = @"^[\x00-\x7F]*$";
                                    if (Regex.IsMatch(value, asciiPattern))  //判断生成二维码的内容是否符合Code128条形码编码规则
                                    {
                                        Bitmap qrCode = writer.Write(value);
                                        string tempImagePath = System.IO.Path.GetTempFileName() + ".png";
                                        qrCode.Save(tempImagePath, ImageFormat.Png);
                                        Excel.Range cellForImage = sheet.Cells[i, usedColumn + 1];
                                        cellForImage.Rows.RowHeight = 50; // 或根据实际需要调整行高
                                        cellForImage.Columns.ColumnWidth = 50; // 或根据实际需要调整列宽

                                        // 插入图片到单元格
                                        Excel.Shape shape = sheet.Shapes.AddPicture(tempImagePath, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoTrue,
                                            cellForImage.Left, cellForImage.Top, -1, -1);
                                        shape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

                                        File.Delete(tempImagePath); // 删除临时文件
                                        count += 1;
                                    }
                                    else
                                    {
                                        continue;
                                    }

                                }
                                string result = $"生成条形码功能已完成，共有{count.ToString()}个条形码生成";
                                ShowLabel(run_result_label, true, result);
                                StartTimer();

                            }
                        }
                        else
                        {
                            ShowLabel(run_result_label, true, "未选择生成条码格式");
                            StartTimer();
                        }

                        break;
                }
            }
            catch (Exception ex)
            {
                ShowLabel(run_result_label, true, ex.Message);
            }
            finally
            {
                //左侧按钮状态改变
                tabControl1.Enabled = true;
                move_sheet_button.Enabled = true;
                add_sheet_button.Enabled = true;
                transposition_button.Enabled = true;
                regex_button.Enabled = true;
                payslip_button.Enabled = true;
                contents_button.Enabled = true;
                QR_button.Enabled = true;
                ThisAddIn.app.ScreenUpdating = true;
                ThisAddIn.app.DisplayAlerts = true;
            }

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
            contents_to_sheet_radioButton.Visible = false;
            sheet_to_contents_radioButton.Visible = false;
            QR_label.Visible = false;
            QR_listBox.Visible = false;
            QR_radioButton.Visible = false;
            QR_radioButton.Checked = false;
            foreColor_select_button.Visible = false;
            foreColor_label.Visible = false;
            backColor_select_button.Visible = false;
            backColor_label.Visible = false;
            QR_logo_label.Visible = false;
            QR_logo_pictureBox.Visible = false;
            BC_radioButton.Visible = false;
            BC_radioButton.Checked = false;

            //左侧按钮状态改变
            tabControl1.Enabled = false;
            move_sheet_button.Enabled = false;
            add_sheet_button.Enabled = false;
            transposition_button.Enabled = false;
            regex_button.Enabled = false;
            payslip_button.Enabled = false;
            contents_button.Enabled = false;
            QR_button.Enabled = false;
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
            QR_button.Enabled = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }

        //目录页功能启动
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
            regex_clear_button.Visible = false;
            contents_to_sheet_radioButton.Visible = true;
            sheet_to_contents_radioButton.Visible = true;
            regex_run_button.Visible = true;
            contents_to_sheet_radioButton.Checked = false;
            sheet_to_contents_radioButton.Checked = false;
            QR_label.Visible = false;
            QR_listBox.Visible = false;
            QR_radioButton.Visible = false;
            QR_radioButton.Checked = false;
            foreColor_select_button.Visible = false;
            foreColor_label.Visible = false;
            backColor_select_button.Visible = false;
            backColor_label.Visible = false;
            QR_logo_label.Visible = false;
            QR_logo_pictureBox.Visible = false;
            BC_radioButton.Visible = false;
            BC_radioButton.Checked = false;
            selectfunction = 0;
        }

        //目录页单选模式改变时        
        private void contents_to_sheet_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (contents_to_sheet_radioButton.Checked == true)
            {
                selectfunction = 2;
                which_field_label.Text = "链接哪列";
                which_field_label.Visible = true;
                which_field_combobox.Visible = true;
                MessageBox.Show("1.该选项是将包含一个命名为‘目录’的表时，自动将目录各行生成链接空白表。\n\n2.各目录项从第二行开始。\n\n3.仅对“目录”表中存在的表进行生成和链接。\n\n4.“目录”表需自行手工建立");
            }
        }

        private void sheet_to_contents_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (sheet_to_contents_radioButton.Checked == true)
            {
                selectfunction = 3;
                which_field_label.Visible = false;
                which_field_label.Text = "提取哪列";
                which_field_combobox.Visible = false;
                which_field_combobox.Items.Clear();
                which_field_combobox.Text = "";
            }
        }

        //二维码功能启动
        private void QR_button_Click(object sender, EventArgs e)
        {
            //右侧功能区初始化
            function_title_label.Text = "生成二维码/条形码";
            which_field_label.Visible = false;
            which_field_combobox.Visible = false;
            what_type_label.Visible = false;
            what_type_combobox.Visible = false;
            regex_rule_label.Visible = false;
            regex_rule_textbox.Visible = false;
            run_result_label.Visible = false;
            regex_clear_button.Visible = false;
            contents_to_sheet_radioButton.Visible = false;
            sheet_to_contents_radioButton.Visible = false;
            regex_run_button.Visible = true;
            contents_to_sheet_radioButton.Checked = false;
            sheet_to_contents_radioButton.Checked = false;
            QR_label.Visible = true;
            QR_listBox.Visible = true;
            QR_radioButton.Visible = true;
            QR_radioButton.Checked = true;
            foreColor_select_button.Visible = true;
            foreColor_label.Visible = true;
            backColor_select_button.Visible = true;
            backColor_label.Visible = true;
            QR_logo_label.Visible = true;
            QR_logo_pictureBox.Visible = true;
            QR_logo_pictureBox.Image = ExcelAddIn.Properties.Resources.pic_logo;
            qr_logo_path = "";
            BC_radioButton.Visible = true;
            BC_radioButton.Checked = false;
            selectfunction = 4;
        }

        //二维码图标选择
        private string qr_logo_path = "";
        private void QR_logo_pictureBox_DoubleClick(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "图片文件|*.jpg;*.png;*.bmp;*.gif|All files (*.*)|*.*";
            openFileDialog1.Title = "请选择要添加的二维码图标图片";
            openFileDialog1.AddExtension = true;
            openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                QR_logo_pictureBox.Image = Image.FromFile(openFileDialog1.FileName);
                qr_logo_path = openFileDialog1.FileName;
            }
        }

        //二维码数据列选择框可见时
        private void QR_listBox_VisibleChanged(object sender, EventArgs e)
        {
            if (QR_listBox.Visible == true)
            {
                Excel.Worksheet sheet = ThisAddIn.app.ActiveWorkbook.ActiveSheet;
                QR_listBox.Items.Clear();

                if (sheet.UsedRange.Rows.Count > 1 && !string.IsNullOrEmpty(sheet.Cells[1, 1].Value))
                {
                    foreach (Excel.Range cell in sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, sheet.UsedRange.Columns.Count]])
                    {
                        QR_listBox.Items.Add(cell.Value);
                    }
                }
                else QR_listBox.Items.Add("空表");
            }
        }

        //二维码单选模式改变时
        private void BC_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (BC_radioButton.Checked == true)
            {
                QR_listBox.SelectionMode = SelectionMode.One;
                QR_logo_label.Visible = false;
                QR_logo_pictureBox.Visible = false;
                foreColor_select_button.Visible = false;
                foreColor_label.Visible = false;
                backColor_select_button.Visible = false;
                backColor_label.Visible = false;
            }
        }

        private void QR_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (QR_radioButton.Checked == true)
            {
                QR_listBox.SelectionMode = SelectionMode.MultiExtended;
                QR_logo_label.Visible = true;
                QR_logo_pictureBox.Visible = true;
                foreColor_select_button.Visible = true;
                foreColor_label.Visible = true;
                backColor_select_button.Visible = true;
                backColor_label.Visible = true;
            }
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

        //时间控件，控制完成提示标签显示3秒后消失
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
            aTimer.Interval = 3000;          //3 seconds
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

        //判断当前工作簿中是否存在指定表，如存在，返回true，否则返回false
        public static bool SheetExist(string sheet_name)
        {
            foreach (Excel.Worksheet sheet in ThisAddIn.app.ActiveWorkbook.Worksheets)
            {
                if (sheet.Name == sheet_name)
                {
                    return true;
                }
            }
            return false;
        }

        //判断指定字段位于指定表中的哪一列（数字）
        public static int TargetField(Excel.Worksheet targetSheet, string targetText)
        {

            foreach (Excel.Range targetCell in targetSheet.Range[targetSheet.Cells[1, 1], targetSheet.Cells[1, targetSheet.UsedRange.Columns.Count]])
            {
                if (targetCell.Value == targetText)
                {
                    return targetCell.Column;
                }
            }
            return 0;
        }

        List<string> tableNames = new List<string>();

        //DataBase模块 运行按钮点击事件（读取数据库表，并在DataGridView中显示）
        private void dbrun_button_Click(object sender, EventArgs e)
        {
            dbsheet_comboBox.Items.Clear();
            tableNames.Clear();
            if (string.IsNullOrEmpty(dbaddress_textBox.Text))
            {
                MessageBox.Show("数据库地址不能为空！", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                switch (dbtype_comboBox.SelectedIndex)
                {
                    //MySQL
                    case 0:
                        string connectionString0 = $"server={dbaddress_textBox.Text};user={dbuser_textBox.Text};database={dbname_textBox.Text};port={dbport_textBox.Text};password={dbpwd_textBox.Text}";

                        // 连接到数据库并获取表名列表
                        tableNames = MysqlDB.GetTableNames(connectionString0);

                        if (tableNames.Count == 1 && tableNames[0].Contains(":"))
                        {
                            // 处理错误情况
                            database_result_label.Text = "数据库连接失败";
                        }
                        else
                        {
                            database_result_label.Text = "数据库连接成功，数据库中包含" + tableNames.Count + "张表";

                            // 将表名添加到 ComboBox
                            dbsheet_comboBox.DataSource = tableNames;
                            //foreach (var tableName in tableNames)
                            //{
                            //    dbsheet_comboBox.Items.Add(tableName);
                            //}

                            // 默认选择第一个表
                            if (dbsheet_comboBox.Items.Count > 0)
                            {
                                dbsheet_comboBox.SelectedIndex = 0;
                                // 更新 DataGridView
                                UpdateDataGridView(dbsheet_comboBox.SelectedItem.ToString());
                            }
                            else
                            {
                                MessageBox.Show("数据库中没有表");
                                dbsheet_dataGridView.DataSource = null;
                            }
                        }
                        break;

                    //SQL Server
                    case 1:
                        string connectionString1 = $"Data Source={dbaddress_textBox.Text};Initial Catalog={dbname_textBox.Text};User ID={dbuser_textBox.Text};Password={dbpwd_textBox.Text}";

                        // 连接到数据库并获取表名列表
                        tableNames = SQLServerDB.GetTableNames(connectionString1);

                        if (tableNames.Count == 1 && tableNames[0].Contains(":"))
                        {
                            // 处理错误情况
                            database_result_label.Text = "数据库连接失败";

                        }
                        else
                        {
                            database_result_label.Text = "数据库连接成功，数据库中包含" + tableNames.Count + "张表";


                            // 将表名添加到 ComboBox
                            dbsheet_comboBox.DataSource = tableNames;
                            //foreach (var tableName in tableNames)
                            //{
                            //    dbsheet_comboBox.Items.Add(tableName);
                            //}

                            // 默认选择第一个表
                            if (dbsheet_comboBox.Items.Count > 0)
                            {
                                dbsheet_comboBox.SelectedIndex = 0;
                                // 更新 DataGridView
                                UpdateDataGridView(dbsheet_comboBox.SelectedItem.ToString());
                            }
                            else
                            {
                                MessageBox.Show("数据库中没有表");
                                dbsheet_dataGridView.DataSource = null;
                            }
                        }
                        break;

                    //Access
                    case 2:
                        string connectionString2 = null;
                        if (string.IsNullOrEmpty(dbpwd_textBox.Text))
                        {
                            connectionString2 = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbaddress_textBox.Text};Persist Security Info=False;";
                        }
                        else
                        {
                            connectionString2 = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbaddress_textBox.Text};Persist Security Info=False;Jet OLEDB:Database Password={dbpwd_textBox.Text};";
                        }
                        tableNames = AccessDB.GetTableNames(connectionString2);
                        if (tableNames.Count == 1 && tableNames[0].Contains(":"))
                        {
                            // 处理错误情况
                            database_result_label.Text = "数据库连接失败";
                        }
                        else
                        {
                            database_result_label.Text = "数据库连接成功，数据库中包含" + tableNames.Count + "张表";
                            dbname_textBox.Text = System.IO.Path.GetFileNameWithoutExtension(dbaddress_textBox.Text);

                            // 将表名添加到 ComboBox
                            dbsheet_comboBox.DataSource = tableNames;
                            //foreach (var tableName in tableNames)
                            //{
                            //    dbsheet_comboBox.Items.Add(tableName);
                            //}

                            // 默认选择第一个表
                            if (dbsheet_comboBox.Items.Count > 0)
                            {
                                dbsheet_comboBox.SelectedIndex = 0;
                                // 更新 DataGridView
                                UpdateDataGridView(dbsheet_comboBox.SelectedItem.ToString());
                            }
                            else
                            {
                                MessageBox.Show("数据库中没有表");
                                dbsheet_dataGridView.DataSource = null;
                            }
                        }
                        break;

                    //SQLite
                    case 3:
                        string connectionString3 = null;
                        if (string.IsNullOrEmpty(dbpwd_textBox.Text))
                        {
                            connectionString3 = $"Data Source={dbaddress_textBox.Text};Version=3;";
                        }
                        else
                        {
                            connectionString3 = $"Data Source={dbaddress_textBox.Text};Version=3;Password={dbpwd_textBox.Text};";
                        }
                        tableNames = SqliteDB.GetTableNames(connectionString3);
                        if (tableNames.Count == 1 && tableNames[0].Contains(":"))
                        {
                            // 处理错误情况
                            database_result_label.Text = "数据库连接失败";
                        }
                        else
                        {
                            database_result_label.Text = "数据库连接成功，数据库中包含" + tableNames.Count + "张表";
                            dbname_textBox.Text = System.IO.Path.GetFileNameWithoutExtension(dbaddress_textBox.Text);

                            // 将表名添加到 ComboBox
                            dbsheet_comboBox.DataSource = tableNames;
                            //foreach (var tableName in tableNames)
                            //{
                            //    dbsheet_comboBox.Items.Add(tableName);
                            //}

                            // 默认选择第一个表
                            if (dbsheet_comboBox.Items.Count > 0)
                            {
                                dbsheet_comboBox.SelectedIndex = 0;
                                // 更新 DataGridView
                                UpdateDataGridView(dbsheet_comboBox.SelectedItem.ToString());
                            }
                            else
                            {
                                MessageBox.Show("数据库中没有表");
                                dbsheet_dataGridView.DataSource = null;
                            }
                        }
                        break;

                    //PostgreSQL
                    case 4:
                        string connectionString4 = $"Host={dbaddress_textBox.Text};Port={dbport_textBox.Text};Username={dbuser_textBox.Text};Password={dbpwd_textBox.Text};Database={dbname_textBox.Text}";
                        // 连接到数据库并获取表名列表
                        tableNames = PostgreSqlDB.GetTableNames(connectionString4);

                        if (tableNames.Count == 1 && tableNames[0].Contains(":"))
                        {
                            // 处理错误情况
                            database_result_label.Text = "数据库连接失败";
                        }
                        else
                        {
                            database_result_label.Text = "数据库连接成功，数据库中包含" + tableNames.Count + "张表";

                            // 将表名添加到 ComboBox
                            dbsheet_comboBox.DataSource = tableNames;
                            //foreach (var tableName in tableNames)
                            //{
                            //    dbsheet_comboBox.Items.Add(tableName);
                            //}

                            // 默认选择第一个表
                            if (dbsheet_comboBox.Items.Count > 0)
                            {
                                dbsheet_comboBox.SelectedIndex = 0;
                                // 更新 DataGridView
                                UpdateDataGridView(dbsheet_comboBox.SelectedItem.ToString());
                            }
                            else
                            {
                                MessageBox.Show("数据库中没有表");
                                dbsheet_dataGridView.DataSource = null;
                            }
                        }
                        break;

                    //Oracle
                    case 5:
                        string connectionString5 = $"Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={dbaddress_textBox.Text})(PORT={dbport_textBox.Text}))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME={dbname_textBox.Text})));User Id={dbuser_textBox.Text};Password={dbpwd_textBox.Text};";

                        // 连接到数据库并获取表名列表
                        tableNames = OracleDB.GetTableNames(connectionString5);

                        if (tableNames.Count == 1 && tableNames[0].Contains(":"))
                        {
                            // 处理错误情况
                            database_result_label.Text = "数据库连接失败";
                        }
                        else
                        {
                            database_result_label.Text = "数据库连接成功，数据库中包含" + tableNames.Count + "张表";

                            // 将表名添加到 ComboBox
                            dbsheet_comboBox.DataSource = tableNames;
                            //foreach (var tableName in tableNames)
                            //{
                            //    dbsheet_comboBox.Items.Add(tableName);
                            //}

                            // 默认选择第一个表
                            if (dbsheet_comboBox.Items.Count > 0)
                            {
                                dbsheet_comboBox.SelectedIndex = 0;
                                // 更新 DataGridView
                                UpdateDataGridView(dbsheet_comboBox.SelectedItem.ToString());
                            }
                            else
                            {
                                MessageBox.Show("数据库中没有表");
                                dbsheet_dataGridView.DataSource = null;
                            }
                        }
                        break;

                    ////DB2
                    //case 6:
                    //    string connectionString6 = $"Server={dbaddress_textBox.Text}:{dbport_textBox.Text};Database={dbname_textBox.Text};UID={dbuser_textBox.Text};PWD={dbpwd_textBox.Text};";


                    //    // 连接到数据库并获取表名列表
                    //    tableNames = Db2DB.GetTableNames(connectionString6);

                    //    if (tableNames.Count == 1 && tableNames[0].Contains(":"))
                    //    {
                    //        // 处理错误情况
                    //        database_result_label.Text = "数据库连接失败";
                    //    }
                    //    else
                    //    {
                    //        database_result_label.Text = "数据库连接成功，数据库中包含" + tableNames.Count + "张表";

                    //        // 将表名添加到 ComboBox
                    //        dbsheet_comboBox.DataSource = tableNames;
                    //        //    foreach (var tableName in tableNames)
                    //        //    {
                    //        //        dbsheet_comboBox.Items.Add(tableName);
                    //        //    }

                    //        // 默认选择第一个表
                    //        if (dbsheet_comboBox.Items.Count > 0)
                    //        {
                    //            dbsheet_comboBox.SelectedIndex = 0;
                    //            // 更新 DataGridView
                    //            UpdateDataGridView(dbsheet_comboBox.SelectedItem.ToString());
                    //        }
                    //        else
                    //        {
                    //            MessageBox.Show("数据库中没有表");
                    //            dbsheet_dataGridView.DataSource = null;
                    //        }
                    //    }
                    //    break;

                    default:
                        break;
                }
            }
        }

        // 更新 DataGridView
        private void UpdateDataGridView(string tableName)
        {
            var dataTable = GetDataTable(tableName);
            dbsheet_dataGridView.DataSource = dataTable;
        }

        // 获取数据库指定数据表
        private DataTable GetDataTable(string tableName)
        {
            string connectionString = null;
            switch (dbtype_comboBox.SelectedIndex)
            {
                //MySQL
                case 0:
                    connectionString = $"server={dbaddress_textBox.Text};user={dbuser_textBox.Text};database={dbname_textBox.Text};port={dbport_textBox.Text};password={dbpwd_textBox.Text}";
                    using (var connection = new MySqlConnection(connectionString))
                    {
                        connection.Open();
                        using (var command = new MySqlCommand($"SELECT * FROM {tableName}", connection))
                        {
                            using (var adapter = new MySqlDataAdapter(command))
                            {
                                var dataTable = new DataTable();
                                adapter.Fill(dataTable);
                                return dataTable;
                            }
                        }
                    }

                //SQL Server
                case 1:
                    connectionString = $"Data Source={dbaddress_textBox.Text};Initial Catalog={dbname_textBox.Text};User ID={dbuser_textBox.Text};Password={dbpwd_textBox.Text}";
                    using (var connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        using (var command = new SqlCommand($"SELECT * FROM {tableName}", connection))
                        {
                            using (var adapter = new SqlDataAdapter(command))
                            {
                                var dataTable = new DataTable();
                                adapter.Fill(dataTable);
                                return dataTable;
                            }
                        }
                    }

                //Access
                case 2:
                    if (string.IsNullOrEmpty(dbpwd_textBox.Text))
                    {
                        connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbaddress_textBox.Text};Persist Security Info=False;";
                    }
                    else
                    {
                        connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbaddress_textBox.Text};Persist Security Info=False;Jet OLEDB:Database Password={dbpwd_textBox.Text};";
                    }

                    using (var connection = new OleDbConnection(connectionString))
                    {
                        connection.Open();
                        using (var command = new OleDbCommand($"SELECT * FROM {tableName}", connection))
                        {
                            using (var adapter = new OleDbDataAdapter(command))
                            {
                                var dataTable = new DataTable();
                                adapter.Fill(dataTable);
                                return dataTable;
                            }
                        }
                    }

                //SQLite
                case 3:
                    if (string.IsNullOrEmpty(dbpwd_textBox.Text))
                    {
                        connectionString = $"Data Source={dbaddress_textBox.Text};Version=3;";
                    }
                    else
                    {
                        connectionString = $"Data Source={dbaddress_textBox.Text};Version=3;Password={dbpwd_textBox.Text};";
                    }

                    using (var connection = new SQLiteConnection(connectionString))
                    {
                        connection.Open();
                        using (var command = new SQLiteCommand($"SELECT * FROM {tableName}", connection))
                        {
                            using (var adapter = new SQLiteDataAdapter(command))
                            {
                                var dataTable = new DataTable();
                                adapter.Fill(dataTable);
                                return dataTable;
                            }
                        }
                    }

                //PostgreSQL
                case 4:
                    connectionString = $"Host={dbaddress_textBox.Text};Port={dbport_textBox.Text};Username={dbuser_textBox.Text};Password={dbpwd_textBox.Text};Database={dbname_textBox.Text}";
                    using (var connection = new NpgsqlConnection(connectionString))
                    {
                        connection.Open();
                        using (var command = new NpgsqlCommand($"SELECT * FROM {tableName}", connection))
                        {
                            using (var adapter = new NpgsqlDataAdapter(command))
                            {
                                var dataTable = new DataTable();
                                adapter.Fill(dataTable);
                                return dataTable;
                            }
                        }
                    }

                //Oracle
                case 5:
                    connectionString = $"Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={dbaddress_textBox.Text})(PORT={dbport_textBox.Text}))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME={dbname_textBox.Text})));User Id={dbuser_textBox.Text};Password={dbpwd_textBox.Text};";
                    using (var connection = new OracleConnection(connectionString))
                    {
                        connection.Open();
                        using (var command = new OracleCommand($"SELECT * FROM {tableName}", connection))
                        {
                            using (var adapter = new OracleDataAdapter(command))
                            {
                                var dataTable = new DataTable();
                                adapter.Fill(dataTable);
                                return dataTable;
                            }
                        }
                    }

                ////DB2
                //case 6:
                //    connectionString = $"Server={dbaddress_textBox.Text}:{dbport_textBox.Text};Database={dbname_textBox.Text};UID={dbuser_textBox.Text};PWD={dbpwd_textBox.Text};";

                //    using (var connection = new DB2Connection(connectionString))
                //    {
                //        connection.Open();
                //        using (var command = new DB2Command($"SELECT * FROM {tableName}", connection))
                //        {
                //            using (var adapter = new DB2DataAdapter(command))
                //            {
                //                var dataTable = new DataTable();
                //                adapter.Fill(dataTable);
                //                return dataTable;
                //            }
                //        }
                //    }

                default:
                    return null;
            }
        }

        //数据库表选择改变事件
        private void dbsheet_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dbsheet_comboBox.SelectedIndex == -1)
            {
                dbsheet_dataGridView.DataSource = null;
            }
            else
            {
                string selectedTableName = dbsheet_comboBox.SelectedItem.ToString();
                UpdateDataGridView(selectedTableName);
            }
        }

        //数据库类型选择改变事件
        private void dbtype_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Tab5Clear();
            switch (dbtype_comboBox.SelectedIndex)
            {
                //MySQL
                case 0:
                    dbport_textBox.Text = "3306";
                    break;

                //SQL Server
                case 1:
                    dbport_textBox.Text = "1433";
                    break;

                //Access
                case 2:
                    dbport_textBox.Text = "";
                    dbaddress_textBox.Text = "请双击本框选择Access数据库文件";
                    break;

                //Sqlite
                case 3:
                    dbport_textBox.Text = "";
                    dbaddress_textBox.Text = "请双击本框选择SQLite数据库文件";
                    break;

                //PostgreSQL
                case 4:
                    dbport_textBox.Text = "5432";
                    break;

                //Oracle
                case 5:
                    dbport_textBox.Text = "1521";
                    break;

                ////DB2
                //case 6:                               
                //    dbport_textBox.Text = "50000";
                //    break;

                default:
                    break;
            }
        }

        private void dbclear_button_Click(object sender, EventArgs e)
        {
            Tab5Clear();
        }

        //功能5页面统清空
        private void Tab5Clear()
        {
            if (dbtype_comboBox.SelectedIndex == 2 || dbtype_comboBox.SelectedIndex == 3)
            {
                dbaddress_textBox.Text = "请双击本框选择Access数据库文件";
                dbname_textBox.ReadOnly = true;
                dbport_textBox.ReadOnly = true;
                dbuser_textBox.ReadOnly = true;
                dbuser_textBox.Text = "";
                dbname_textBox.Text = "";
                dbpwd_textBox.Text = "";
                tableNames.Clear();
                find_keyword_textBox.Text = "";
                dbsheet_comboBox.DataSource = null;
                dbsheet_dataGridView.DataSource = null;
                database_result_label.Text = string.Empty;
                dbexport_result_label.Text = string.Empty;
            }
            else
            {
                dbname_textBox.ReadOnly = false;
                dbport_textBox.ReadOnly = false;
                dbuser_textBox.ReadOnly = false;
                dbaddress_textBox.Text = "";
                dbuser_textBox.Text = "";
                dbname_textBox.Text = "";
                dbpwd_textBox.Text = "";
                tableNames.Clear();
                find_keyword_textBox.Text = "";
                dbsheet_comboBox.DataSource = null;
                dbsheet_dataGridView.DataSource = null;
                database_result_label.Text = string.Empty;
                dbexport_result_label.Text = string.Empty;
            }
        }

        //数据导出按钮响应事件
        private void dbexport_button_Click(object sender, EventArgs e)
        {
            string newsheetname = dbname_textBox.Text + "." + dbsheet_comboBox.Text;
            try
            {
                dbexport_result_label.Text = "正在导出......";
                ExportDataGridViewToExcel(dbsheet_dataGridView, newsheetname);
                dbexport_result_label.Text = "导出完成！";
                MessageBox.Show("导出成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                dbexport_result_label.Text = "导出失败，原因为：" + ex.Message;
            }
        }

        //预览数据导出到Excel
        internal void ExportDataGridViewToExcel(DataGridView dataGridView, string newsheetname)
        {
            ThisAddIn.app.DisplayAlerts = false;
            ThisAddIn.app.ScreenUpdating = false;

            workbook = ThisAddIn.app.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.Worksheets.Add();
            worksheet.Name = newsheetname;
            worksheet.Activate();
            try
            {
                // 写入标题
                for (int i = 0; i < dataGridView.ColumnCount; i++)
                {
                    worksheet.Cells[1, i + 1] = dataGridView.Columns[i].HeaderText;
                }

                // 写入数据
                for (int rowIndex = 0; rowIndex < dataGridView.Rows.Count; rowIndex++)
                {
                    for (int colIndex = 0; colIndex < dataGridView.Columns.Count; colIndex++)
                    {
                        // 获取单元格的值
                        var cellValue = dataGridView.Rows[rowIndex].Cells[colIndex].Value?.ToString() ?? "";

                        // 将值写入 Excel 单元格
                        worksheet.Cells[rowIndex + 2, colIndex + 1] = cellValue; // +2 因为第一行是标题行
                    }
                }

                // 刷新 Excel 表格
                worksheet.Cells.EntireColumn.AutoFit();
                worksheet.Activate();
            }
            catch (Exception ex)
            {
                // 处理异常
                MessageBox.Show($"Error: {ex.Message}");
            }
            finally
            {
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;
            }
        }

        // 双击地址文本框事件
        private void dbaddress_textBox_DoubleClick(object sender, EventArgs e)
        {

            switch (dbtype_comboBox.SelectedIndex)
            {
                case 2:
                    using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
                    {
                        openFileDialog1.Title = "请选择要打开的Access数据库文件";
                        openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                        openFileDialog1.Filter = "Access数据库文件|*.accdb;*.mdb";
                        if (openFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            dbaddress_textBox.Text = openFileDialog1.FileName;
                        }
                    }
                    break;
                case 3:
                    using (OpenFileDialog openFileDialog1 = new OpenFileDialog())
                    {
                        openFileDialog1.Title = "请选择要打开的SQLite数据库文件";
                        openFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                        openFileDialog1.Filter = "SQLite数据库文件|*.db;*.db3;*.sqlite|全部文件(*.*)|*.*";
                        if (openFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            dbaddress_textBox.Text = openFileDialog1.FileName;
                        }
                    }
                    break;
            }

        }

        //地址文本框改变事件
        private void dbaddress_textBox_TextChanged(object sender, EventArgs e)
        {
            if (dbaddress_textBox.Text == "请双击本框选择SQLite数据库文件" || dbaddress_textBox.Text == "请双击本框选择Access数据库文件")
            {
                dbaddress_textBox.ForeColor = Color.DarkGray;
                dbaddress_textBox.Font = new Font(dbaddress_textBox.Font.FontFamily, 8, FontStyle.Italic);

            }
            else
            {
                dbaddress_textBox.ForeColor = SystemColors.WindowText;
                dbaddress_textBox.Font = new Font(dbaddress_textBox.Font.FontFamily, 10.5f, FontStyle.Regular);
            }
        }

        //关键字查找数据库中需导出表名
        private List<string> FindComboBoxItems(string searchText)
        {
            // 获取 ComboBox 的原始数据源
            List<string> items = new List<string>();
            foreach (string item in tableNames)
            {
                if (item.ToLower().Contains(searchText.ToLower()))
                {
                    items.Add(item);
                }
            }
            return items;
        }

        //模糊查找按钮点击事件
        private void find_keywordbutton_pictureBox_Click(object sender, EventArgs e)
        {
            if (find_keyword_textBox.Text != "")
            {
                List<string> resultItems = new List<string>();
                resultItems = FindComboBoxItems(find_keyword_textBox.Text);
                if (resultItems.Count == 0 || resultItems.Count == tableNames.Count)
                {
                    dbsheet_comboBox.DataSource = tableNames;
                    dbexport_result_label.Text = "未找到表";
                }
                else
                {
                    dbsheet_comboBox.DataSource = resultItems;
                    ShowLabel(dbexport_result_label, true, $"找到{resultItems.Count.ToString()}个表");
                    StartTimer();

                }

            }
            else
            {
                dbsheet_comboBox.DataSource = tableNames;
            }
        }

        //关键字文本框改变事件
        private void find_keyword_textBox_TextChanged(object sender, EventArgs e)
        {
            if (find_keyword_textBox.Text == "")
            {
                dbsheet_comboBox.DataSource = tableNames;
                find_keywordclear_pictureBox.Visible = false;
                dbexport_result_label.Text = string.Empty;
            }
            else
            {
                find_keywordclear_pictureBox.Visible = true;
                dbexport_result_label.Text = string.Empty;
            }

        }

        //清空关键字查找框内容事件
        private void find_keywordclear_pictureBox_Click(object sender, EventArgs e)
        {
            if (find_keyword_textBox.Text != "")
            {
                find_keyword_textBox.Text = string.Empty;
                find_keywordclear_pictureBox.Visible = false;
            }
        }

        Color qrForeColor = Color.Black;
        Color qrBackColor = Color.White;
        private void foreColor_select_button_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                foreColor_select_button.BackColor = colorDialog1.Color;
                qrForeColor = colorDialog1.Color;
            }
        }

        private void backColor_select_button_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                backColor_select_button.BackColor = colorDialog1.Color;
                qrBackColor = colorDialog1.Color;
            }
        }

        private static int GetGrayScale(Color color)
        {
            // 计算灰度值
            return (int)(0.299 * color.R + 0.587 * color.G + 0.114 * color.B);
        }

        private void chart_reset_button_Click(object sender, EventArgs e)
        {
            chart_select_comboBox.SelectedIndex = 0;
            chart_range_textBox.Text = "";
            chart_pictureBox.Image = null;
            CleanupTempFiles();
        }

        [STAThread]
        private void chart_create_button_Click(object sender, EventArgs e)
        {
            switch (chart_select_comboBox.SelectedIndex)
            {
                case 0:
                    if (string.IsNullOrEmpty(chart_range_textBox.Text))
                    {
                        MessageBox.Show("请先选择数据范围");
                        return;
                    }
                    Excel.Range range =workbook.ActiveSheet.Range[chart_range_textBox.Text];
                    string activeSheetName = workbook.ActiveSheet.Name;
                    chart_pictureBox.Image = null;

                    Thread staThread = new Thread(() => GenerateWordCloud(activeSheetName, range));
                    staThread.SetApartmentState(ApartmentState.STA); // 设置STA模式
                    staThread.IsBackground = true; // 设为后台线程
                    staThread.Start();
                    break;
                case 1:
                    break;
            }
        }

        // 在窗体类中增加字段
        private string _currentWordCloudTempPath;


        // 将词频字典转换为 WordScore 集合
        private static IEnumerable<WordScore> ConvertToWordScores(Dictionary<string, int> frequencyDict)
        {
            return frequencyDict
                .Select(kv => new WordScore(
                    Score: kv.Value,  
                    Word: kv.Key))
                .OrderByDescending(ws => ws.Score); // 按数值排序
        }

        private void GenerateWordCloud(string activeSheetName, Excel.Range range)
        {
            try
            {
                ThisAddIn.app.DisplayAlerts = false;
                ThisAddIn.app.ScreenUpdating = false;
                this.Invoke(new Action(() => 
                {
                    //ThisAddIn.app.ScreenUpdating = false;
                    //ThisAddIn.app.DisplayAlerts = false;
                    chart_create_button.Enabled = false;
                    chart_reset_button.Enabled = false;
                    chart_select_comboBox.Enabled = false;
                    chart_range_textBox.ReadOnly = true;
                }));

                // 统计字符串频率
                Dictionary<string, int> frequencyDict = new Dictionary<string, int>();
                foreach (Excel.Range cell in range.Cells)
                {
                    string cellText = cell.Text.ToString().Trim();
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        if (frequencyDict.ContainsKey(cellText))
                            frequencyDict[cellText]++;
                        else
                            frequencyDict[cellText] = 1;
                    }
                }

                if (frequencyDict.Count == 0)
                {
                    MessageBox.Show("选中的范围没有文本内容");
                    return;
                }

                string wordCloudSheetName = "_wordcloud";
                int i = 1;
                while (SheetExist(wordCloudSheetName)) 
                {
                    wordCloudSheetName = "_wordcloud"+i.ToString();
                    i++;
                }
                Excel.Worksheet wordCloudSheet = workbook.Worksheets.Add(Before: workbook.ActiveSheet);
                wordCloudSheet.Name = wordCloudSheetName;
                wordCloudSheet.Range["A1"].Value = "key_word";
                wordCloudSheet.Range["B1"].Value = "count";
                int rowIndex = 2;
                var frequencyDictSorted = frequencyDict.OrderByDescending(x => x.Value);
                foreach (var pair in frequencyDictSorted)
                {
                    wordCloudSheet.Cells[rowIndex, 1] = pair.Key;
                    wordCloudSheet.Cells[rowIndex, 2] = pair.Value;
                    rowIndex++;
                }


                // 强制刷新工作表计算
                wordCloudSheet.Application.Calculate();

                 //生成词云图

                // 词云图方向
                TextOrientations[] orientations =
                [
                    TextOrientations.PreferHorizontal,
                    TextOrientations.PreferVertical,
                    TextOrientations.HorizontalOnly,
                    TextOrientations.VerticalOnly,
                    TextOrientations.Random
                ];

                string fontName="Microsoft YaHei";
                int textDirection = 0;
                

                this.Invoke(new Action(() =>
                {
                    if (comboBoxFonts.SelectedItem is FontInfo selectedFont)
                    {
                        fontName = selectedFont.EnglishName;
                    }
                    textDirection = comboBoxTextDirection.SelectedIndex;
                }));

                WordCloud wc = WordCloud.Create(new WordCloudOptions(600, 400, ConvertToWordScores(frequencyDict))
                {
                    TextOrientation = orientations[textDirection],
                    FontManager = new FontManager([SKTypeface.FromFamilyName($"fontName")]),    
                });
                byte[] pngBytes = wc.ToSKBitmap().Encode(SKEncodedImageFormat.Png, 100).AsSpan().ToArray();

                // 获取 Windows 临时文件夹路径
                _currentWordCloudTempPath = Path.Combine(Path.GetTempPath(), "wordCloud.png");
                File.WriteAllBytes(_currentWordCloudTempPath, pngBytes);

                // 保存SVG
                if (checkBoxSVG.Checked)
                {
                    this.Invoke(new Action(() => // 使用Invoke切换到UI线程
                    {
                        using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                        {
                            saveFileDialog.Title = "请选择保存SVG图片路径";
                            saveFileDialog.Filter = "SVG图片(*.svg)|*.svg";
                            saveFileDialog.DefaultExt = ".svg";
                            saveFileDialog.AddExtension = true;
                            saveFileDialog.FileName = "worldCloud.svg";
                            saveFileDialog.OverwritePrompt = true;
                            if (saveFileDialog.ShowDialog() == DialogResult.OK)
                            {
                                string svg = saveFileDialog.FileName;
                                File.WriteAllText(svg, wc.ToSvg());
                            }
                        }
                    }));
                }

                this.Invoke(new Action(() =>
                {
                    chart_pictureBox.Image = Image.FromFile(_currentWordCloudTempPath);

                }));

                // 插入图片并获取图片对象
                Excel.Shape wordCloudShape = wordCloudSheet.Shapes.AddPicture
                (
                    Filename: _currentWordCloudTempPath,
                    LinkToFile: MsoTriState.msoFalse,  // 不链接到文件（嵌入到文档）
                    SaveWithDocument: MsoTriState.msoTrue,
                    Left: 500,
                    Top: 100,
                    Width: -1,  // -1 表示保持原始宽度
                    Height: -1   // -1 表示保持原始高度
                );
             }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误: " + ex.Message);
            }
            finally
            {
                this.Invoke(new Action(() =>
                {
                    chart_create_button.Enabled = true;
                    chart_reset_button.Enabled = true;
                    chart_select_comboBox.Enabled = true;
                    chart_range_textBox.ReadOnly = false;
                    ThisAddIn.app.ScreenUpdating = true;
                    ThisAddIn.app.DisplayAlerts = true;
                }));
                workbook.Worksheets[activeSheetName].Select();
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;
            }
        }

        // 临时文件清理方法
        private void CleanupTempFiles()
        {
            try
            {
                if (!string.IsNullOrEmpty(_currentWordCloudTempPath) && File.Exists(_currentWordCloudTempPath))
                {
                    File.Delete(_currentWordCloudTempPath);
                    _currentWordCloudTempPath = null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"清理临时文件失败: {ex.Message}");
            }

            // 清理图片资源
            if (chart_pictureBox.Image != null)
            {
                chart_pictureBox.Image.Dispose();
                chart_pictureBox.Image = null;
            }
        }

        private void chart_select_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(chart_select_comboBox.SelectedIndex)
            {
                case 0:
                    chart_range_textBox.Text = "";
                    groupBoxWordCloud.Visible = true;
                    break;
                case 1:
                    chart_range_textBox.Text = "";
                    groupBoxWordCloud.Visible = false;
                    break;
            }
        }
    }

    internal class MysqlDB
    {
        internal static List<string> GetTableNames(string connString)
        {
            using (var connection = new MySqlConnection(connString))
            {
                try
                {
                    connection.Open();
                    using (var command = new MySqlCommand("SHOW TABLES", connection))             //"SHOW TABLES"
                    {

                        using (var reader = command.ExecuteReader())
                        {
                            List<string> tableNames = new List<string>();
                            while (reader.Read())
                            {
                                tableNames.Add(reader.GetString(0));
                            }
                            return tableNames;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("数据库连接失败：" + ex.Message);
                    return new List<string>() { ex.Message + ":" };
                }
            }
        }
    }

    internal class SQLServerDB
    {
        internal static List<string> GetTableNames(string connString)
        {
            using (SqlConnection connection = new SqlConnection(connString))
            {
                try
                {
                    connection.Open();

                    // SQL 查询语句，用于获取所有表名
                    string query = @"
                    SELECT 
                        t.NAME AS TableName
                    FROM 
                        sys.tables t
                    WHERE 
                        t.is_ms_shipped = 0;"; // 排除系统表

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            List<string> tableNames = new List<string>();
                            while (reader.Read())
                            {
                                tableNames.Add(reader.GetString(0));
                            }
                            return tableNames;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("数据库连接失败：" + ex.Message);
                    return new List<string>() { ex.Message + ":" };
                }
            }
        }
    }

    internal class AccessDB
    {
        internal static List<string> GetTableNames(string connString)
        {
            using (var connection = new OleDbConnection(connString))
            {
                try
                {
                    connection.Open();

                    // 使用 OpenSchema 方法获取所有表的信息
                    DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    if (schemaTable != null && schemaTable.Rows.Count > 0)
                    {
                        List<string> tableNames = new List<string>();
                        foreach (DataRow row in schemaTable.Rows)
                        {
                            string tableName = row["TABLE_NAME"].ToString();
                            tableNames.Add(tableName);
                        }
                        return tableNames;
                    }
                    else
                    {
                        return new List<string>() { "没有找到任何表" };
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("数据库连接失败：" + ex.Message);
                    return new List<string>() { ex.Message + ":" };
                }
            }
        }
    }

    internal class SqliteDB
    {
        internal static List<string> GetTableNames(string connString)
        {
            using (var connection = new SQLiteConnection(connString))
            {
                try
                {
                    connection.Open();

                    // SQL 查询语句，用于获取所有表名
                    string query = @"
                    SELECT 
                        name
                    FROM 
                        sqlite_master
                    WHERE 
                        type='table'
                        AND name NOT LIKE 'sqlite_%'";

                    using (var command = new SQLiteCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            List<string> tableNames = new List<string>();
                            while (reader.Read())
                            {
                                tableNames.Add(reader.GetString(0));
                            }
                            return tableNames;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("数据库连接失败：" + ex.Message);
                    return new List<string>() { ex.Message };
                }
            }
        }
    }

    internal class PostgreSqlDB
    {
        internal static List<string> GetTableNames(string connString)
        {
            using (var connection = new NpgsqlConnection(connString))
            {
                try
                {
                    connection.Open();
                    using (var command = new NpgsqlCommand("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public'", connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            List<string> tableNames = new List<string>();
                            while (reader.Read())
                            {
                                tableNames.Add(reader.GetString(0));
                            }
                            return tableNames;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("数据库连接失败：" + ex.Message);
                    return new List<string>() { ex.Message };
                }
            }
        }
    }

    internal class OracleDB
    {
        internal static List<string> GetTableNames(string connString)
        {
            using (OracleConnection connection = new OracleConnection(connString))
            {
                try
                {
                    connection.Open();
                    // 获取当前用户下的所有表名
                    using (OracleCommand command = new OracleCommand("SELECT table_name FROM user_tables", connection))
                    {
                        using (OracleDataReader reader = command.ExecuteReader())
                        {
                            List<string> tableNames = new List<string>();
                            while (reader.Read())
                            {
                                tableNames.Add(reader.GetString(0));
                            }
                            return tableNames;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("数据库连接失败：" + ex.Message);
                    return new List<string>() { ex.Message + ":" };
                }
            }
        }
    }

    //public class Db2DB
    //{
    //    internal static List<string> GetTableNames(string connString)
    //    {
    //        using (DB2Connection connection = new DB2Connection(connString))
    //        {
    //            try
    //            {
    //                connection.Open();

    //                // 获取当前用户下的所有表名
    //                string query = @"SELECT TABNAME FROM SYSCAT.TABLES WHERE TABSCHEMA NOT IN ('SYSIBM', 'SYSSPATIAL', 'SYSSTAT', 'SYSCAT', 'SYSSQL', 'SYSBAR', 'SYSLIB', 'SYSPUBLIC','IBMCONSOLE','SYSTOOLS') AND TYPE = 'T' AND OWNERTYPE='U'";
    //                using (DB2Command command = new(query, connection))
    //                {
    //                    using (DB2DataReader reader = command.ExecuteReader())
    //                    {
    //                        List<string> tableNames = new List<string>();
    //                        while (reader.Read())
    //                        {
    //                            tableNames.Add(reader.GetString(0));
    //                        }
    //                        return tableNames;
    //                    }
    //                }
    //            }
    //            catch (Exception ex)
    //            {
    //                MessageBox.Show("数据库连接失败：" + ex.Message);
    //                return new List<string>() { ex.Message + ":" };
    //            }
    //        }
    //    }
    //}
}
