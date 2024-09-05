using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualBasic.FileIO;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;


namespace ExcelAddIn
{
    public partial class Form4 : Form
    {
        Excel.Worksheet worksheet = ThisAddIn.app.ActiveWorkbook.Worksheets["_rename"];
        internal  int regulation_number = 1;
        public static int runButtonClicked = 0;
        public static int  resetButtonClicked= 0;
        private string command=Ribbon1.runcommand;
        private bool isCheckedAll = false;

        public Form4()
        {
            InitializeComponent();
        }

        private async void Form4_Load(object sender, EventArgs e)
        {
            if (command == "file")
            {
                BeginInvoke(new MethodInvoker(() => ListBoxItemsLoad()));
            }
            else
            {
                Invoke(new MethodInvoker(() => 
                {
                    file_type_checkedListBox.Enabled = false;
                    select_all_checkBox.Visible = false;
                }));
                
            }
            
            file_type_checkedListBox.CheckOnClick = true;
            delete_select_radioButton.Select();

            runButtonClicked = 0;
            resetButtonClicked = 0;
            filename_regular_label2.Visible = false;
            filename_regular_ComboBox2.Visible = false;
            filename_regular_textBox2.Visible = false;
            regulation_add_pictureBox2.Visible = false;
            regulation_reduce_pictureBox2.Visible = false;

            filename_regular_label3.Visible = false;
            filename_regular_ComboBox3.Visible = false;
            filename_regular_textBox3.Visible = false;
            regulation_reduce_pictureBox3.Visible = false;
            result_dm_label.Text = "";
        }

        //初始化CheckListBox的Items
        private void ListBoxItemsLoad()
        {
            file_type_checkedListBox.Items.Clear();
            List<string> items = new List<string>();
            int usedrowsCount = worksheet.UsedRange.Rows.Count;
            
            for (int i = 2; i <= usedrowsCount; i++)
            {
                string fileName = Path.Combine(worksheet.Cells[i, 1].Value, worksheet.Cells[i, 2].Value);
                if (File.Exists(fileName))
                {
                    string extName = Path.GetExtension(fileName);
                    if (!items.Contains(extName))
                    {
                        items.Add(extName);
                    }
                }
            }
            items.Sort();
            if (items.Count > 0)
            {
                foreach (string item in items) 
                {
                    file_type_checkedListBox.Items.Add(item);                
                }
                select_all_checkBox.Visible = true;
            }
            else
            {
                select_all_checkBox.Visible = false;
            }
        }

        //添加第2行规则
        private void regulation_add_pictureBox1_Click(object sender, EventArgs e)
        {
            regulation_add_pictureBox1.Visible = false;
            filename_regular_label2.Visible = true;
            filename_regular_ComboBox2.SelectedIndex = -1;
            filename_regular_ComboBox2.Visible = true;
            filename_regular_textBox2.Text = "";
            filename_regular_textBox2.Visible = true;
            regulation_add_pictureBox2.Visible = true;
            regulation_reduce_pictureBox2.Visible = true;
        } 

        //删除第2行规则
        private void regulation_reduce_pictureBox2_Click(object sender, EventArgs e)
        {
            regulation_add_pictureBox1.Visible = true;
            filename_regular_label2.Visible = false;
            filename_regular_ComboBox2.SelectedIndex = -1;
            filename_regular_ComboBox2.Visible = false;
            filename_regular_textBox2.Text = "";
            filename_regular_textBox2.Visible = false;
            regulation_add_pictureBox2.Visible = false;
            regulation_reduce_pictureBox2.Visible = false;
        }

        //添加第3行规则
        private void regulation_add_pictureBox2_Click(object sender, EventArgs e)
        {
            regulation_add_pictureBox2.Visible = false;
            regulation_reduce_pictureBox2.Visible=false;
            filename_regular_label3.Visible = true;
            filename_regular_ComboBox3.SelectedIndex = -1;
            filename_regular_ComboBox3.Visible = true;
            filename_regular_textBox3.Text = "";
            filename_regular_textBox3.Visible = true;
            regulation_reduce_pictureBox3.Visible = true;
        }

        //删除第3行规则
        private void regulation_reduce_pictureBox3_Click(object sender, EventArgs e)
        {
            regulation_add_pictureBox2.Visible = true;
            regulation_reduce_pictureBox2.Visible = true;
            filename_regular_label3.Visible = false;
            filename_regular_ComboBox3.SelectedIndex = -1;
            filename_regular_ComboBox3.Visible = false;
            filename_regular_textBox3.Text = "";
            filename_regular_textBox3.Visible = false;
            regulation_reduce_pictureBox3.Visible = false;
        }

        private void quit_button_Click(object sender, EventArgs e)
        {
            result_dm_label.Text = "正在关闭窗口......";
            this.Close();
        }

        private void reset_button_Click(object sender, EventArgs e)
        {
            ListBoxItemsLoad();

            filename_regular_ComboBox1.SelectedIndex = -1;
            filename_regular_textBox1.Text = "";
            regulation_add_pictureBox1.Visible = true;

            filename_regular_label2.Visible = false;
            filename_regular_ComboBox2.SelectedIndex = -1;
            filename_regular_ComboBox2.Visible = false;
            filename_regular_textBox2.Text = "";
            filename_regular_textBox2.Visible = false;
            regulation_add_pictureBox2.Visible = false;
            regulation_reduce_pictureBox2.Visible = false;

            filename_regular_label3.Visible = false;
            filename_regular_ComboBox3.SelectedIndex = -1;
            filename_regular_ComboBox3.Visible = false;
            filename_regular_textBox3.Text = "";
            filename_regular_textBox3.Visible = false;
            regulation_reduce_pictureBox3.Visible = false;
            delete_select_radioButton.Select();
            resetButtonClicked += 1;
        }

        private void run_button_Click(object sender, EventArgs e)
        {
            List<string> rules_about_extension = new List<string>();
            List<string> rules_about_startwith = new List<string>();
            List<string> rules_about_endwith = new List<string>();
            List<string> rules_about_contains = new List<string>();
            List<string> rules_about_notcontains = new List<string>();

            for(int i=0;i<file_type_checkedListBox.Items.Count;i++)
            {
                if(file_type_checkedListBox.GetItemChecked(i))
                {
                    rules_about_extension.Add(file_type_checkedListBox.Items[i].ToString());
                }
            }

            ProcessRules(filename_regular_ComboBox1, filename_regular_textBox1, rules_about_startwith, rules_about_endwith, rules_about_contains, rules_about_notcontains);
            ProcessRules(filename_regular_ComboBox2, filename_regular_textBox2, rules_about_startwith, rules_about_endwith, rules_about_contains, rules_about_notcontains);
            ProcessRules(filename_regular_ComboBox3, filename_regular_textBox3, rules_about_startwith, rules_about_endwith, rules_about_contains, rules_about_notcontains);
            
            //检测“包含”和“不包含”是否设置相同字符串
            List<string> duplicates=rules_about_contains.Intersect(rules_about_notcontains).ToList();
            if(duplicates.Count()>0)
            {
                if(MessageBox.Show("规则中“包含”和“不包含”不能设置相同字符串，确认后将删除重复规则后继续，取消将返回重新设定规则", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Error)==DialogResult.OK)
                {
                    rules_about_contains.Remove(duplicates[0]);
                }
            }

            
            if (rules_about_extension.Count == 0 && rules_about_startwith.Count == 0 && rules_about_endwith.Count == 0 && rules_about_contains.Count == 0 && rules_about_notcontains.Count == 0)
            {
                result_dm_label.Text = "请先设置规则！";
                return;
            }
            else
            {
                try
                {
                    switch (command)
                    {
                        case "file":
                            List<string> file_list_all = new List<string>();                     //_rename表中的所有文件名
                            List<string> filter_list_ext = new List<string>();                //文件类型规则过滤出的结果
                            List<string> filter_list_startwith = new List<string>();             //第1次规则过滤出的结果（文件开头匹配））
                            List<string> filter_list_endwith = new List<string>();           //第2次规则过滤出的结果（文件结尾匹配）
                            List<string> filter_list_contains = new List<string>();           //第3次规则过滤出的结果（文件包含匹配）
                            List<string> filter_list_notcontains = new List<string>();          //第4次规则过滤出的结果（文件不包含匹配）
                            List<string> resultList = new List<string>();                        //最终删除表

                            //读取_rename表中所有文件名
                            for (int i = 2; i <= worksheet.UsedRange.Rows.Count; i++)
                            {
                                string file_path = Path.Combine(worksheet.Cells[i, 1].Value, worksheet.Cells[i, 2].Value);
                                file_list_all.Add(file_path);
                            }

                            //筛选文件扩展名规则
                            if (rules_about_extension.Count > 0)
                            {
                                foreach (string searching_file_first in file_list_all)
                                {
                                    if (rules_about_extension.Contains(Path.GetExtension(searching_file_first)))
                                    {
                                        filter_list_ext.Add(searching_file_first);
                                    }
                                }
                            }
                            else filter_list_ext = file_list_all;

                            if (rules_about_startwith.Count == 0 && rules_about_endwith.Count == 0 && rules_about_contains.Count == 0 && rules_about_notcontains.Count == 0)
                            {
                                resultList = filter_list_ext;
                            }
                            else
                            {
                                filter_list_startwith = MatchingList(filter_list_ext, rules_about_startwith, "startwith");
                                filter_list_endwith = MatchingList(filter_list_ext, rules_about_endwith, "endwith");
                                filter_list_contains = MatchingList(filter_list_ext, rules_about_contains, "contains");
                                filter_list_notcontains = MatchingList(filter_list_ext, rules_about_notcontains, "notcontains");

                                //合并多个过滤条件选出文件名并去重
                                resultList = filter_list_startwith.Concat(filter_list_endwith).Concat(filter_list_contains)
                                    .Concat(filter_list_notcontains).Distinct().ToList();
                            }

                            if (resultList.Count != 0)
                            {
                                switch (this.delete_select_radioButton.Checked)
                                {
                                    case true:
                                        if (MessageBox.Show($"删除文件是高风险操作，本次将移除{resultList.Count}个文件至回收站！", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                                        {
                                            foreach (string item in resultList)
                                            {
                                                if (File.Exists(item))
                                                {
                                                    try
                                                    {
                                                        // 将文件移动到回收站
                                                        FileSystem.DeleteFile(item, UIOption.AllDialogs, RecycleOption.SendToRecycleBin);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        result_dm_label.Text = "删除文件时发生错误: " + ex.Message;
                                                        continue;
                                                    }
                                                }else result_dm_label.Text = $"文件{item}不存在,继续删除下一个";                                                
                                            }
                                            result_dm_label.Text = "文件已移动到回收站";
                                        }
                                        else return;
                                        break;
                                    case false:
                                        if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                                        {
                                            if (MessageBox.Show($"本次将移除{resultList.Count}个文件至目标文件夹！如有文件名重复，将重命名移动文件", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                                            {
                                                string destinationFolder = folderBrowserDialog1.SelectedPath;
                                                MoveFiles(resultList, destinationFolder, "file");
                                                result_dm_label.Text = $"文件已移动到目标文件夹：{destinationFolder}";
                                            }
                                            else return;
                                        }
                                        break;
                                }
                            }
                            break;
                        case "folder":
                            List<string> folder_list_all = new List<string>();                     //_rename表中的所有文件夹
                            List<string> filter_folderlist_startwith = new List<string>();             //第1次规则过滤出的结果（文件夹开头匹配））
                            List<string> filter_folderlist_endwith = new List<string>();           //第2次规则过滤出的结果（文件夹结尾匹配）
                            List<string> filter_folderlist_contains = new List<string>();           //第3次规则过滤出的结果（文件夹包含匹配）
                            List<string> filter_folderlist_notcontains = new List<string>();          //第4次规则过滤出的结果（文件夹不包含匹配）
                            List<string> resultFolderList = new List<string>();                        //最终结果表表

                            //读取_rename表中所有文件夹名
                            for (int i = 2; i <= worksheet.UsedRange.Rows.Count; i++)
                            {
                                string file_path = worksheet.Cells[i, 1].Value+"\\"+worksheet.Cells[i, 2].Value;
                                folder_list_all.Add(file_path);
                            }

                            if (rules_about_startwith.Count == 0 && rules_about_endwith.Count == 0 && rules_about_contains.Count == 0 && rules_about_notcontains.Count == 0)
                            {
                                result_dm_label.Text = "未检测到任何规则，请检查规则设置";
                                return;
                            }
                            else
                            {
                                filter_list_startwith = MatchingList(folder_list_all, rules_about_startwith, "startwith");
                                filter_list_endwith = MatchingList(folder_list_all, rules_about_endwith, "endwith");
                                filter_list_contains = MatchingList(folder_list_all, rules_about_contains, "contains");
                                filter_list_notcontains = MatchingList(folder_list_all, rules_about_notcontains, "notcontains");

                                //合并多个过滤条件选出文件名并去重
                                resultList = filter_list_startwith.Concat(filter_list_endwith).Concat(filter_list_contains)
                                    .Concat(filter_list_notcontains).Distinct().ToList();
                            }

                            if (resultList.Count != 0)
                            {
                                switch (this.delete_select_radioButton.Checked)
                                {
                                    case true:
                                        if (MessageBox.Show($"删除文件夹是高风险操作，文件夹内的文件将一并删除！本次将移除{resultList.Count}个文件夹至回收站！", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                                        {
                                            foreach (string item in resultList)
                                            {
                                                if (Directory.Exists(item))
                                                {
                                                    try
                                                    {
                                                        // 将文件移动到回收站
                                                        FileSystem.DeleteDirectory(item, UIOption.AllDialogs, RecycleOption.SendToRecycleBin);
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        result_dm_label.Text = "删除文件夹时发生错误: " + ex.Message;
                                                        continue;
                                                    }
                                                }
                                                else result_dm_label.Text = $"要删除的文件夹{item}不存在，继续删除下一个";                                                
                                            }
                                            result_dm_label.Text = "文件夹已移动到回收站";
                                        }
                                        else return;
                                        break;
                                    case false:
                                        if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                                        {
                                            if (MessageBox.Show($"本次将移动{resultList.Count}个文件至目标文件夹！如有文件夹名重复，将重命名移动文件夹", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                                            {
                                                string destinationFolder = folderBrowserDialog1.SelectedPath;
                                                MoveFiles(resultList, destinationFolder, "folder");
                                                result_dm_label.Text = $"文件夹已移动到目标文件夹：{destinationFolder}";
                                            }
                                            else return;
                                        }
                                        break;
                                }
                            }
                            break;
                    }
                    
                }
                catch(Exception ex)
                {
                    result_dm_label.Text=$"{ex.Message}！";
                }
                finally
                {
                    runButtonClicked+=1;
                }
            }
        }

        //读取窗体输入的规则内容
        private void ProcessRules(ComboBox comboBox, System.Windows.Forms.TextBox textBox, List<string> startWith, List<string> endWith, List<string> contains, List<string> notContains)
        {
            if (comboBox.SelectedIndex != -1 && !string.IsNullOrEmpty(textBox.Text))
            {
                switch (comboBox.SelectedIndex)
                {
                    case 0:
                        startWith.Add(textBox.Text);
                        break;
                    case 1:
                        endWith.Add(textBox.Text);
                        break;
                    case 2:
                        contains.Add(textBox.Text);
                        break;
                    case 3:
                        notContains.Add(textBox.Text);
                        break;
                }
            }
        }


        //筛选器，对符合筛选条件的文件名进行筛选
        private List<string> MatchingList(List<string> listFiles,List<string> listRules,string type)
        {
            List<string> matchingList = new List<string>();
            foreach (string listFile in listFiles)
            {
                string matching_filename = Path.GetFileNameWithoutExtension(listFile);                
                switch (type)
                {
                    //判断是否匹配始于规则的文件名
                    case "startwith":
                        if (listRules.Count > 0)
                        {
                            foreach (string listRule in listRules)
                            {
                                if (matching_filename.StartsWith(listRule))
                                {
                                    matchingList.Add(listFile);
                                }
                            }
                        }
                        break;

                    //判断是否匹配止于规则的文件名
                    case "endwith":
                        if (listRules.Count > 0)
                        {
                            foreach (string listRule in listRules)
                            {
                                if (matching_filename.EndsWith(listRule))
                                {
                                    matchingList.Add(listFile);
                                }
                            }
                        }

                        break;

                    //判断是否匹配包含规则的文件名
                    case "contains":
                        if (listRules.Count > 0)
                        {
                            foreach (string listRule in listRules)
                            {
                                if (matching_filename.Contains(listRule))
                                {
                                    matchingList.Add(listFile);
                                }
                            }
                        }

                        break;

                    //判断是否匹配不包含规则的文件名
                    case "notcontains":
                        if (listRules.Count > 0)
                        {
                            foreach (string listRule in listRules)
                            {
                                if (!matching_filename.Contains(listRule))
                                {
                                    matchingList.Add(listFile);
                                }
                            }
                        }
                        break;
                }                
            }
            return matchingList;
        }                                       

        //移动文件/文件夹
        private void MoveFiles(List<string> listFiles, string destinationFolder,string type)
        {           
            switch (type)
            {
                case "file":
                    foreach (string listFile in listFiles)
                    {
                        if (File.Exists(listFile))
                        {
                            try
                            {
                                string sourceFileName = Path.GetFileName(listFile);
                                string destinationPath = Path.Combine(destinationFolder, sourceFileName);

                                // 如果目标路径已经存在同名文件，则重命名目标文件
                                int i = 1;
                                string originalFileName = Path.GetFileNameWithoutExtension(listFile);
                                while (File.Exists(destinationPath))
                                {
                                    string newFileName = $"{originalFileName}({i}){Path.GetExtension(destinationPath)}";
                                    destinationPath = Path.Combine(destinationFolder, newFileName);
                                    i++;
                                }
                                // 执行文件移动操作
                                File.Move(listFile, destinationPath);
                            }
                            catch (Exception ex)
                            {
                                result_dm_label.Text = ex.Message;
                                continue;
                            }
                        }else result_dm_label.Text = "源文件{listFile}不存在，继续移动下一个";
                    }
                    break;
                case "folder":
                    foreach (string listFile in listFiles)
                    {
                        if (Directory.Exists(listFile))
                        {
                            try
                            {
                                string originalDirectoryName =listFile.Split('\\').Last();
                                string destinationPath = destinationFolder+"\\"+originalDirectoryName;

                                // 如果目标路径已经存在同名文件夹，则重命名目标文件夹
                                int i = 1;
                                 
                                while (Directory.Exists(destinationPath))
                                {
                                    destinationPath = $"{originalDirectoryName}({i}){destinationFolder}";
                                    i++;
                                }
                                // 执行文件移动操作
                                Directory.Move(listFile, destinationPath);
                            }
                            catch (Exception ex)
                            {
                                result_dm_label.Text = ex.Message;
                                continue;
                            }
                        }
                    }
                    break;
            }            
        }

        private void file_type_checkedListBox_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (isCheckedAll) return;
            
            if (e.NewValue == CheckState.Unchecked)
            {
                select_all_checkBox.Checked = false;
            }
            else
            {
                // 使用BeginInvoke来延迟执行状态检查和更新
                this.BeginInvoke(new MethodInvoker(() =>
                {
                    bool allChecked = true;
                    for (int i = 0; i < file_type_checkedListBox.Items.Count; i++)
                    {
                        if (file_type_checkedListBox.GetItemCheckState(i) != CheckState.Checked)
                        {
                            allChecked = false;
                            break;
                        }
                    }
                    if (allChecked)
                    {
                        // 全选时联动select_all_checkBox被选中
                        select_all_checkBox.Checked = true;
                    }
                }));
            }
        }

        //全选/取消全选
        private void select_all_checkBox_Click(object sender, EventArgs e)
        {
            isCheckedAll = true;
            for (int i = 0; i < file_type_checkedListBox.Items.Count; i++)
            {
                file_type_checkedListBox.SetItemChecked(i, select_all_checkBox.Checked);

            }
            isCheckedAll = false;
        }
    }



    /// <summary>
    /// 写了一个类，用来存储控件的信息和添加删除方法，本程序中未使用该类
    /// 仅为参考，方便其他功能调用
    /// </summary>
    public class ControlInfo
    {
        public string ControlName{get;set;}

        public enum ControlType
        {
            TextBox,
            ComboBox,
            CheckBox,
            RadioButton,
            Label,
            PictureBox,
            ListBox,
            Button
        }

        public void SetControlName(string name)
        {
            ControlName = name;
        }


        public string GetControlName()
        {
            return ControlName;
        }

         public bool ControlExists(string controlName,Form form)
        {
            foreach (Control control in form.Controls)
            {
                if (control.Name == controlName)
                {
                    return true;
                }
            }
            return false;
        }

        public static Control CreateControl(string controlName, ControlType controlType, System.Drawing.Point location)
        {
            Control newControl = null;

            // 根据控件类型创建控件
            switch (controlType)
            {
                case ControlType.Button:
                    newControl = new System.Windows.Forms.Button();
                    ((System.Windows.Forms.Button)newControl).Text = "按钮"; // 设置按钮文本
                    break;
                case ControlType.TextBox:
                    newControl = new System.Windows.Forms.TextBox();
                    break;
                case ControlType.Label:
                    newControl = new System.Windows.Forms.Label();
                    ((System.Windows.Forms.Label)newControl).Text = "标签"; // 设置标签文本
                    break;
                case ControlType.PictureBox:
                    newControl = new PictureBox();
                    break;
                case ControlType.ComboBox:
                    newControl = new ComboBox();
                    break;
                case ControlType.ListBox:
                    newControl = new System.Windows.Forms.ListBox();
                    break;
                default:
                    MessageBox.Show("不支持的控件类型: " + controlType);
                    return null;
            }
            // 设置控件的名称和位置
            newControl.Name = controlName;
            newControl.Location = location;

            return newControl;
        }
    }
}
