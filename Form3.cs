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
                string[] files_name = new string[files.Length];
                string[] files_path = new string[files.Length];
                switch (format_radioButton1.Checked)
                {
                    case true:
                        for (int i = 0; i < files.Length; i++)
                        {
                            //从读取文件全名中分解路径
                            files_path[i] = Path.GetDirectoryName(files[i]);

                            //从读取文件全名中分解文件名
                            files_name[i] = Path.GetFileName(files[i]);

                            //改新文件名
                            string new_file_name = files_name[i].Split('-')[1];
                            if(new_file_name.EndsWith(" ") || new_file_name.StartsWith(" "))
                            {
                                new_file_name=new_file_name.Trim();
                            }
                            int exist_file = 0;
                            while (File.Exists(Path.Combine(files_path[i], new_file_name)))
                            {
                                exist_file++;
                                new_file_name= Path.GetFileNameWithoutExtension(new_file_name)+ "("+exist_file.ToString()+")"+ Path.GetExtension(new_file_name);
                            }
                            Directory.Move(Path.Combine(files_path[i],files_name[i]),Path.Combine(files_path[i],new_file_name));
                        }
                        break;
                    case false:
                        for (int i = 0; i < files.Length; i++)
                        {
                            //从读取文件全名中分解路径
                            files_path[i] = Path.GetDirectoryName(files[i]);

                            //从读取文件全名中分解文件名
                            files_name[i] = Path.GetFileName(files[i]);

                            //改新文件名
                            string new_file_name = files_name[i].Split('-')[0];
                            string new_file_ext = Path.GetExtension(files_name[i]);                            
                            if (new_file_name.EndsWith(" ") || new_file_name.StartsWith(" "))
                            {
                                new_file_name=new_file_name.Trim();
                            }
                            int exist_file = 0;

                            while (File.Exists(Path.Combine(files_path[i], new_file_name+new_file_ext)))
                            {
                                exist_file++;
                                new_file_name = new_file_name + "(" + exist_file.ToString() + ")";
                            }
                            Directory.Move(Path.Combine(files_path[i],files_name[i]),Path.Combine(files_path[i],new_file_name+new_file_ext));
                        }
                        break;
                }

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
            List<string> filelist = new List<string>( Directory.GetFiles(folderPath,"*.mp3",SearchOption.AllDirectories));
            filelist.RemoveAll(file => (System.IO.File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden);
            string[] fileNames= filelist.ToArray();
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
