namespace ExcelAddIn
{
    partial class Form6
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form6));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.clear_pictureBox = new System.Windows.Forms.PictureBox();
            this.subfolder_checkBox = new System.Windows.Forms.CheckBox();
            this.batch_result_label = new System.Windows.Forms.Label();
            this.batch_run_button = new System.Windows.Forms.Button();
            this.folder_path_textBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.xml_treeView = new System.Windows.Forms.TreeView();
            this.last_pictureBox = new System.Windows.Forms.PictureBox();
            this.next_pictureBox = new System.Windows.Forms.PictureBox();
            this.preview_pictureBox = new System.Windows.Forms.PictureBox();
            this.begin_pictureBox = new System.Windows.Forms.PictureBox();
            this.sequence_label = new System.Windows.Forms.Label();
            this.single_result_label = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.xml_path_textBox = new System.Windows.Forms.TextBox();
            this.run_button = new System.Windows.Forms.Button();
            this.reset_button = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.version_label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.clear_pictureBox)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.last_pictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.next_pictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.preview_pictureBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.begin_pictureBox)).BeginInit();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControl1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tabControl1.ItemSize = new System.Drawing.Size(50, 120);
            this.tabControl1.Location = new System.Drawing.Point(1, 3);
            this.tabControl1.Multiline = true;
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(799, 392);
            this.tabControl1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabControl1.TabIndex = 0;
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Font = new System.Drawing.Font("微软雅黑", 10.8F);
            this.tabPage1.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.tabPage1.Location = new System.Drawing.Point(124, 4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(671, 384);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "一、电子发票读取";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.AutoSize = true;
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(671, 388);
            this.panel1.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.clear_pictureBox);
            this.groupBox2.Controls.Add(this.subfolder_checkBox);
            this.groupBox2.Controls.Add(this.batch_result_label);
            this.groupBox2.Controls.Add(this.batch_run_button);
            this.groupBox2.Controls.Add(this.folder_path_textBox);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Location = new System.Drawing.Point(3, 232);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(665, 149);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "批量写入Excel";
            // 
            // clear_pictureBox
            // 
            this.clear_pictureBox.Image = global::ExcelAddIn.Properties.Resources.clear;
            this.clear_pictureBox.Location = new System.Drawing.Point(434, 61);
            this.clear_pictureBox.Name = "clear_pictureBox";
            this.clear_pictureBox.Size = new System.Drawing.Size(15, 15);
            this.clear_pictureBox.TabIndex = 12;
            this.clear_pictureBox.TabStop = false;
            this.clear_pictureBox.Click += new System.EventHandler(this.clear_pictureBox_Click);
            // 
            // subfolder_checkBox
            // 
            this.subfolder_checkBox.AutoSize = true;
            this.subfolder_checkBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.subfolder_checkBox.Location = new System.Drawing.Point(40, 88);
            this.subfolder_checkBox.Name = "subfolder_checkBox";
            this.subfolder_checkBox.Size = new System.Drawing.Size(99, 21);
            this.subfolder_checkBox.TabIndex = 11;
            this.subfolder_checkBox.Text = "包含子文件夹";
            this.subfolder_checkBox.UseVisualStyleBackColor = true;
            // 
            // batch_result_label
            // 
            this.batch_result_label.AutoSize = true;
            this.batch_result_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.batch_result_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.batch_result_label.Location = new System.Drawing.Point(46, 120);
            this.batch_result_label.Name = "batch_result_label";
            this.batch_result_label.Size = new System.Drawing.Size(0, 17);
            this.batch_result_label.TabIndex = 9;
            // 
            // batch_run_button
            // 
            this.batch_run_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.batch_run_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.batch_run_button.Location = new System.Drawing.Point(467, 51);
            this.batch_run_button.Name = "batch_run_button";
            this.batch_run_button.Size = new System.Drawing.Size(74, 34);
            this.batch_run_button.TabIndex = 7;
            this.batch_run_button.Text = "批量导入";
            this.batch_run_button.UseVisualStyleBackColor = true;
            this.batch_run_button.Click += new System.EventHandler(this.batch_run_button_Click);
            // 
            // folder_path_textBox
            // 
            this.folder_path_textBox.Location = new System.Drawing.Point(40, 55);
            this.folder_path_textBox.Name = "folder_path_textBox";
            this.folder_path_textBox.Size = new System.Drawing.Size(415, 27);
            this.folder_path_textBox.TabIndex = 6;
            this.folder_path_textBox.TextChanged += new System.EventHandler(this.folder_path_textBox_TextChanged);
            this.folder_path_textBox.DoubleClick += new System.EventHandler(this.folder_path_textBox_DoubleClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(44, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(144, 20);
            this.label2.TabIndex = 5;
            this.label2.Text = "双击下框选择文件夹";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.xml_treeView);
            this.groupBox1.Controls.Add(this.last_pictureBox);
            this.groupBox1.Controls.Add(this.next_pictureBox);
            this.groupBox1.Controls.Add(this.preview_pictureBox);
            this.groupBox1.Controls.Add(this.begin_pictureBox);
            this.groupBox1.Controls.Add(this.sequence_label);
            this.groupBox1.Controls.Add(this.single_result_label);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.xml_path_textBox);
            this.groupBox1.Controls.Add(this.run_button);
            this.groupBox1.Controls.Add(this.reset_button);
            this.groupBox1.Location = new System.Drawing.Point(3, 0);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(665, 226);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "单文件写入Excel";
            // 
            // xml_treeView
            // 
            this.xml_treeView.Location = new System.Drawing.Point(347, 17);
            this.xml_treeView.Name = "xml_treeView";
            this.xml_treeView.Size = new System.Drawing.Size(312, 193);
            this.xml_treeView.TabIndex = 11;
            // 
            // last_pictureBox
            // 
            this.last_pictureBox.Image = global::ExcelAddIn.Properties.Resources.last;
            this.last_pictureBox.Location = new System.Drawing.Point(200, 97);
            this.last_pictureBox.Name = "last_pictureBox";
            this.last_pictureBox.Size = new System.Drawing.Size(20, 20);
            this.last_pictureBox.TabIndex = 10;
            this.last_pictureBox.TabStop = false;
            this.toolTip1.SetToolTip(this.last_pictureBox, "最后一个");
            this.last_pictureBox.Click += new System.EventHandler(this.last_pictureBox_Click);
            // 
            // next_pictureBox
            // 
            this.next_pictureBox.Image = global::ExcelAddIn.Properties.Resources.next1;
            this.next_pictureBox.Location = new System.Drawing.Point(173, 97);
            this.next_pictureBox.Name = "next_pictureBox";
            this.next_pictureBox.Size = new System.Drawing.Size(20, 20);
            this.next_pictureBox.TabIndex = 9;
            this.next_pictureBox.TabStop = false;
            this.toolTip1.SetToolTip(this.next_pictureBox, "下一个");
            this.next_pictureBox.Click += new System.EventHandler(this.next_pictureBox_Click);
            // 
            // preview_pictureBox
            // 
            this.preview_pictureBox.Image = global::ExcelAddIn.Properties.Resources.preview;
            this.preview_pictureBox.Location = new System.Drawing.Point(122, 97);
            this.preview_pictureBox.Name = "preview_pictureBox";
            this.preview_pictureBox.Size = new System.Drawing.Size(20, 20);
            this.preview_pictureBox.TabIndex = 8;
            this.preview_pictureBox.TabStop = false;
            this.toolTip1.SetToolTip(this.preview_pictureBox, "上一个");
            this.preview_pictureBox.Click += new System.EventHandler(this.preview_pictureBox_Click);
            // 
            // begin_pictureBox
            // 
            this.begin_pictureBox.Image = global::ExcelAddIn.Properties.Resources.begin;
            this.begin_pictureBox.Location = new System.Drawing.Point(95, 97);
            this.begin_pictureBox.Name = "begin_pictureBox";
            this.begin_pictureBox.Size = new System.Drawing.Size(20, 20);
            this.begin_pictureBox.TabIndex = 7;
            this.begin_pictureBox.TabStop = false;
            this.toolTip1.SetToolTip(this.begin_pictureBox, "第一个");
            this.begin_pictureBox.Click += new System.EventHandler(this.begin_pictureBox_Click);
            // 
            // sequence_label
            // 
            this.sequence_label.AutoSize = true;
            this.sequence_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.sequence_label.Location = new System.Drawing.Point(149, 97);
            this.sequence_label.Name = "sequence_label";
            this.sequence_label.Size = new System.Drawing.Size(17, 20);
            this.sequence_label.TabIndex = 6;
            this.sequence_label.Text = "0";
            this.sequence_label.TextChanged += new System.EventHandler(this.sequence_label_TextChanged);
            // 
            // single_result_label
            // 
            this.single_result_label.AutoSize = true;
            this.single_result_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.single_result_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.single_result_label.Location = new System.Drawing.Point(34, 133);
            this.single_result_label.Name = "single_result_label";
            this.single_result_label.Size = new System.Drawing.Size(0, 17);
            this.single_result_label.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(234, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "双击下框选择文件或输入文件路径";
            // 
            // xml_path_textBox
            // 
            this.xml_path_textBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.xml_path_textBox.Location = new System.Drawing.Point(27, 68);
            this.xml_path_textBox.Name = "xml_path_textBox";
            this.xml_path_textBox.Size = new System.Drawing.Size(297, 23);
            this.xml_path_textBox.TabIndex = 1;
            this.xml_path_textBox.TextChanged += new System.EventHandler(this.xml_path_textBox_TextChanged);
            this.xml_path_textBox.DoubleClick += new System.EventHandler(this.xml_path_textBox_DoubleClick);
            // 
            // run_button
            // 
            this.run_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.run_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.run_button.Location = new System.Drawing.Point(49, 163);
            this.run_button.Name = "run_button";
            this.run_button.Size = new System.Drawing.Size(67, 32);
            this.run_button.TabIndex = 3;
            this.run_button.Text = "运行";
            this.run_button.UseVisualStyleBackColor = true;
            this.run_button.Click += new System.EventHandler(this.run_button_Click);
            // 
            // reset_button
            // 
            this.reset_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.reset_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.reset_button.Location = new System.Drawing.Point(158, 163);
            this.reset_button.Name = "reset_button";
            this.reset_button.Size = new System.Drawing.Size(68, 32);
            this.reset_button.TabIndex = 4;
            this.reset_button.Text = "重置";
            this.reset_button.UseVisualStyleBackColor = true;
            this.reset_button.Click += new System.EventHandler(this.reset_button_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.richTextBox1);
            this.tabPage2.Font = new System.Drawing.Font("微软雅黑", 10.8F);
            this.tabPage2.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.tabPage2.Location = new System.Drawing.Point(124, 4);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(671, 384);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "二、使用帮助";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // version_label1
            // 
            this.version_label1.AutoSize = true;
            this.version_label1.BackColor = System.Drawing.Color.DarkGreen;
            this.version_label1.Font = new System.Drawing.Font("Franklin Gothic Medium", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.version_label1.ForeColor = System.Drawing.Color.Transparent;
            this.version_label1.Location = new System.Drawing.Point(25, 415);
            this.version_label1.Name = "version_label1";
            this.version_label1.Size = new System.Drawing.Size(49, 16);
            this.version_label1.TabIndex = 4;
            this.version_label1.Text = "version";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.DarkGreen;
            this.label4.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.label4.Font = new System.Drawing.Font("Franklin Gothic Medium", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.SystemColors.Control;
            this.label4.Location = new System.Drawing.Point(298, 413);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(285, 18);
            this.label4.TabIndex = 5;
            this.label4.Text = "Copyright © He Kun.  All rights reserved";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.DarkGreen;
            this.pictureBox1.Location = new System.Drawing.Point(1, 397);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(799, 50);
            this.pictureBox1.TabIndex = 2;
            this.pictureBox1.TabStop = false;
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.richTextBox1.Location = new System.Drawing.Point(54, 27);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(566, 320);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = resources.GetString("richTextBox1.Text");
            // 
            // tabPage3
            // 
            this.tabPage3.Font = new System.Drawing.Font("微软雅黑", 10.8F);
            this.tabPage3.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.tabPage3.Location = new System.Drawing.Point(124, 4);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(671, 384);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "三、退出工具";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // Form6
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.version_label1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form6";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form6";
            this.Load += new System.EventHandler(this.Form6_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.clear_pictureBox)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.last_pictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.next_pictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.preview_pictureBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.begin_pictureBox)).EndInit();
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label version_label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox xml_path_textBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button run_button;
        private System.Windows.Forms.Button reset_button;
        private System.Windows.Forms.TextBox folder_path_textBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button batch_run_button;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Label batch_result_label;
        private System.Windows.Forms.Label single_result_label;
        private System.Windows.Forms.Label sequence_label;
        private System.Windows.Forms.PictureBox last_pictureBox;
        private System.Windows.Forms.PictureBox next_pictureBox;
        private System.Windows.Forms.PictureBox preview_pictureBox;
        private System.Windows.Forms.PictureBox begin_pictureBox;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.CheckBox subfolder_checkBox;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.PictureBox clear_pictureBox;
        private System.Windows.Forms.TreeView xml_treeView;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.TabPage tabPage3;
    }
}