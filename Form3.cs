using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Forms;


namespace ExcelAddIn
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            result_label.Visible = false;
            format_radioButton1.Select();
            fold_path_textBox.Text = "双击选择文件夹";
            fold_path_textBox.ForeColor = Color.LightGray;
            fold_path_textBox.Font = new Font(fold_path_textBox.Font, System.Drawing.FontStyle.Italic);

        }


        static string get_dir_path;
        static string[] files;

        //文本框获得焦点
        private void fold_path_textBox_GotFocus(object sender, EventArgs e)
        {
            fold_path_textBox.Text = "";
        }

        //双击打开文件夹选择框
        private void fold_path_textBox_DoubleClick(object sender, EventArgs e)
        {
            fold_path_textBox.Text = "";
            folderBrowserDialog1.Description = "请选择文件所在文件夹";
            ;
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string select_path = folderBrowserDialog1.SelectedPath;
                fold_path_textBox.Text = select_path;
                fold_path_textBox.ForeColor = Color.Black;
                fold_path_textBox.Font = new Font(fold_path_textBox.Font, System.Drawing.FontStyle.Regular);
                get_dir_path = select_path;
                ReadFileNames(get_dir_path); 
                result_label.Text = "文件夹中现有格式参考：" + Path.GetFileName(files[0]);
                result_label.Visible = true;
            }
            else
            {
                System.Windows.MessageBox.Show("未选择文件夹");
                if (string.IsNullOrEmpty(fold_path_textBox.Text))
                {
                    fold_path_textBox.Text = "双击选择文件夹";
                    fold_path_textBox.ForeColor = Color.LightGray;
                    fold_path_textBox.Font = new Font(fold_path_textBox.Font, System.Drawing.FontStyle.Italic);
                }                
            }
        }

        //单击文本框时
        private void fold_path_textBox_Click(object sender, EventArgs e)
        {
            if (fold_path_textBox.Text == "双击选择文件夹")
            {
                fold_path_textBox.Text="" ;
                result_label.Text = "";
                result_label.Visible = false;
            }
        }




        private void run_button_Click(object sender, EventArgs e)
        {
            fold_path_textBox.Enabled=false;
            run_button.Enabled = false;
            quit_button.Enabled = false;
            format_radioButton1.Enabled = false;
            format_radioButton2.Enabled = false;

            if(files.Length > 0)
            {
                string bat_name = get_dir_path + "\\run.bat";
                FileInfo run_file = new FileInfo(bat_name);
                if (run_file.Exists)
                {
                    run_file.Delete();
                }
                if (format_radioButton1.Checked)
                {
                    string[] files_name = new string[files.Length];
                    string[] files_path = new string[files.Length];


                    for (int i = 0; i < files.Length; i++)
                    {
                        //从读取文件全名中分解文件名
                        files_name[i] = Path.GetFileName(files[i]);

                        //从读取文件全名中分解路径，并判断每段文件夹名中是否包含空格
                        files_path[i] = Path.GetDirectoryName(files[i]);
                        string file_path = files_path[i];
                        if(file_path.Contains(" "))
                        {
                            string[] split_file_path = file_path.Split('\\');
                            for (int a = 0; a < split_file_path.Length; a++)
                            {
                                if (split_file_path[a].Contains(" "))
                                {
                                    split_file_path[a] = "\"" + split_file_path[a] + "\"";
                                }
                            }
                            file_path = string.Join("\\", split_file_path);
                            files_path[i] = file_path;
                        }
                    }

                    //写入新文件名，并判断文件名中是否有空格
                    string[] new_files_name = new string[files.Length];
                    for (int n = 0; n < files_name.Length; n++)
                    {
                        string new_file_name = files_name[n].Split('-')[1];
                        if (new_file_name.StartsWith(" ")|| new_file_name.EndsWith(" "))
                        {
                            new_file_name = new_file_name.Trim();
                        }
                        if (new_file_name.Contains(" "))
                        {
                            new_file_name = "\"" + new_file_name + "\"";
                        }
                        new_files_name[n] = new_file_name;
                        if (files_name[n].Contains(" "))
                        {
                            files_name[n] = "\"" + files_name[n] + "\"";
                        }
                    }
                    using (StreamWriter bat_file = new StreamWriter(bat_name, false, Encoding.GetEncoding("gb2312")))
                    {
                        for (int t = 0; t < files_name.Length; t++)
                        {
                            string code = "ren " + files_path[t] + "\\" + files_name[t] + " " + new_files_name[t];
                            bat_file.WriteLine(code, Encoding.GetEncoding("gb2312"));
                        }
                        bat_file.Close();
                    }                    
                }
                else
                {
                    string[] files_name = new string[files.Length];
                    string[] files_path = new string[files.Length];

                    for (int i = 0; i < files.Length; i++)
                    {
                        //从读取文件全名中分解文件名
                        files_name[i] = Path.GetFileName(files[i]);

                        //从读取文件全名中分解路径，并判断每段文件夹名中是否包含空格
                        files_path[i] = Path.GetDirectoryName(files[i]);
                        string file_path = files_path[i];
                        if (file_path.Contains(" "))
                        {
                            string[] split_file_path = file_path.Split('\\');
                            for (int a = 0; a < split_file_path.Length; a++)
                            {
                                if (split_file_path[a].Contains(" "))
                                {
                                    split_file_path[a] = "\"" + split_file_path[a] + "\"";
                                }
                            }
                            file_path = string.Join("\\", split_file_path);
                            files_path[i] = file_path;
                        }
                    }

                    //写入新文件名，并判断文件名中是否有空格
                    string[] new_files_name = new string[files.Length];
                    for (int n = 0; n < files_name.Length; n++)
                    {
                        string new_file_ext = Path.GetExtension(files_name[n]);
                        string new_file_name = files_name[n].Split('-')[0];
                        if(new_file_name.EndsWith(" ")|| new_file_name.StartsWith(" "))
                        {
                            new_file_name = new_file_name.Trim();
                        }
                        if (new_file_name.Contains(" "))
                        {
                            new_file_name = "\"" + new_file_name + "\"";
                        }
                        new_files_name[n] = new_file_name+new_file_ext;
                        if (files_name[n].Contains(" "))
                        {
                            files_name[n] = "\"" + files_name[n] + "\"";
                        }
                    }

                    using (StreamWriter bat_file = new StreamWriter(bat_name, false, Encoding.GetEncoding("gb2312")))
                    {
                        for (int t = 0; t < files_name.Length; t++)
                        {
                            string code = "ren " + files_path[t] + "\\" + files_name[t] + " " + new_files_name[t];
                            bat_file.WriteLine(code, Encoding.GetEncoding("gb2312"));
                        }
                        bat_file.Close();
                    }
                }

                //调用批处理文件改文件名，并且不显示cmd窗口
                ProcessStartInfo startInfo = new ProcessStartInfo()
                {
                    FileName = bat_name,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                Process proc = Process.Start(startInfo);
                proc.WaitForExit();
                proc.Close();
                System.IO.File.Delete(bat_name);
                System.Windows.MessageBox.Show("文件名修改完毕");
                Process.Start(get_dir_path);

                result_label.Visible = false;
                fold_path_textBox.Text = "双击选择文件夹";
                fold_path_textBox.ForeColor = Color.LightGray;
                fold_path_textBox.Font = new Font(fold_path_textBox.Font, System.Drawing.FontStyle.Italic);
            }
            else
            {
                result_label.Text = "未正确选择文件夹";
                result_label.Visible = true;
            }
            fold_path_textBox.Enabled = true;
            run_button.Enabled = true;
            quit_button.Enabled = true;
            format_radioButton1.Enabled = true;
            format_radioButton2.Enabled = true;
        }

        //读取文件
        static void ReadFileNames(string folderPath)
        {
            string[] fileNames = Directory.GetFiles(folderPath,"*.mp3",SearchOption.AllDirectories);
            files = new string[fileNames.Length];

            for (int i = 0; i < fileNames.Length; i++)
            {
                files[i] = fileNames[i];
            }
        }



        //退出按钮
        private void quit_button_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
