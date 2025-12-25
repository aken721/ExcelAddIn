namespace ExcelAddIn
{
    partial class Form10
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form10));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.sheets_name_comboBox = new System.Windows.Forms.ComboBox();
            this.model_select_button = new System.Windows.Forms.Button();
            this.doc_folder_button = new System.Windows.Forms.Button();
            this.docModel_label = new System.Windows.Forms.Label();
            this.docGenerated_label = new System.Windows.Forms.Label();
            this.result_doc_label = new System.Windows.Forms.Label();
            this.docRun_button = new System.Windows.Forms.Button();
            this.docReset_button = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.placeholder_comboBox = new System.Windows.Forms.ComboBox();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.label5 = new System.Windows.Forms.Label();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.pictureOriginal_comboBox = new System.Windows.Forms.ComboBox();
            this.flowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.label6 = new System.Windows.Forms.Label();
            this.height_textBox = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.flowLayoutPanel3 = new System.Windows.Forms.FlowLayoutPanel();
            this.label8 = new System.Windows.Forms.Label();
            this.width_textBox = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.pictureSize_checkBox = new System.Windows.Forms.CheckBox();
            this.docQuit_button = new System.Windows.Forms.Button();
            this.flowLayoutPanel4 = new System.Windows.Forms.FlowLayoutPanel();
            this.radioButtonAll = new System.Windows.Forms.RadioButton();
            this.radioButtonSelected = new System.Windows.Forms.RadioButton();
            this.flowLayoutPanel1.SuspendLayout();
            this.flowLayoutPanel2.SuspendLayout();
            this.flowLayoutPanel3.SuspendLayout();
            this.flowLayoutPanel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label1.Location = new System.Drawing.Point(281, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(138, 28);
            this.label1.TabIndex = 0;
            this.label1.Text = "文档批量生成";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label2.Location = new System.Drawing.Point(56, 53);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(107, 20);
            this.label2.TabIndex = 1;
            this.label2.Text = "选择文本所在表";
            // 
            // sheets_name_comboBox
            // 
            this.sheets_name_comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.sheets_name_comboBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.sheets_name_comboBox.FormattingEnabled = true;
            this.sheets_name_comboBox.Location = new System.Drawing.Point(169, 53);
            this.sheets_name_comboBox.Name = "sheets_name_comboBox";
            this.sheets_name_comboBox.Size = new System.Drawing.Size(302, 28);
            this.sheets_name_comboBox.TabIndex = 2;
            // 
            // model_select_button
            // 
            this.model_select_button.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.model_select_button.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.model_select_button.Location = new System.Drawing.Point(57, 131);
            this.model_select_button.Name = "model_select_button";
            this.model_select_button.Size = new System.Drawing.Size(103, 35);
            this.model_select_button.TabIndex = 4;
            this.model_select_button.Text = "word模板文件";
            this.model_select_button.UseVisualStyleBackColor = true;
            this.model_select_button.Click += new System.EventHandler(this.model_select_button_Click);
            // 
            // doc_folder_button
            // 
            this.doc_folder_button.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.doc_folder_button.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.doc_folder_button.Location = new System.Drawing.Point(57, 180);
            this.doc_folder_button.Name = "doc_folder_button";
            this.doc_folder_button.Size = new System.Drawing.Size(103, 35);
            this.doc_folder_button.TabIndex = 5;
            this.doc_folder_button.Text = "批量生成目录";
            this.doc_folder_button.UseVisualStyleBackColor = true;
            this.doc_folder_button.Click += new System.EventHandler(this.doc_folder_button_Click);
            // 
            // docModel_label
            // 
            this.docModel_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.docModel_label.Location = new System.Drawing.Point(166, 131);
            this.docModel_label.Name = "docModel_label";
            this.docModel_label.Size = new System.Drawing.Size(305, 35);
            this.docModel_label.TabIndex = 6;
            // 
            // docGenerated_label
            // 
            this.docGenerated_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.docGenerated_label.Location = new System.Drawing.Point(164, 180);
            this.docGenerated_label.Name = "docGenerated_label";
            this.docGenerated_label.Size = new System.Drawing.Size(307, 35);
            this.docGenerated_label.TabIndex = 7;
            // 
            // result_doc_label
            // 
            this.result_doc_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.result_doc_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.result_doc_label.Location = new System.Drawing.Point(65, 220);
            this.result_doc_label.Name = "result_doc_label";
            this.result_doc_label.Size = new System.Drawing.Size(481, 54);
            this.result_doc_label.TabIndex = 8;
            // 
            // docRun_button
            // 
            this.docRun_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.docRun_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.docRun_button.Location = new System.Drawing.Point(137, 385);
            this.docRun_button.Name = "docRun_button";
            this.docRun_button.Size = new System.Drawing.Size(75, 29);
            this.docRun_button.TabIndex = 9;
            this.docRun_button.Text = "生成";
            this.docRun_button.UseVisualStyleBackColor = true;
            this.docRun_button.Click += new System.EventHandler(this.docRun_button_Click);
            // 
            // docReset_button
            // 
            this.docReset_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.docReset_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.docReset_button.Location = new System.Drawing.Point(323, 385);
            this.docReset_button.Name = "docReset_button";
            this.docReset_button.Size = new System.Drawing.Size(75, 29);
            this.docReset_button.TabIndex = 10;
            this.docReset_button.Text = "重置";
            this.docReset_button.UseVisualStyleBackColor = true;
            this.docReset_button.Click += new System.EventHandler(this.docReset_button_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.ForeColor = System.Drawing.Color.Teal;
            this.label3.Location = new System.Drawing.Point(65, 276);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(562, 104);
            this.label3.TabIndex = 11;
            this.label3.Text = "使用说明：\r\n1. 在excel表格中准备数据（第一行为列标题，第一列为生成的文件名，第一列不参与占位符替换）；\r\n2. 利用word编辑模板文件，需要批量修改部" +
    "分用占位符+列名标记。示例：尊敬的【客户姓名】，您的合同金额为【合同金额】元；\r\n3. 占位符可根据习惯选定使用，并在本窗口内选择对应的占位符标识；\r\n4. 选" +
    "择文档批量生成的输出路径。";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label4.Location = new System.Drawing.Point(56, 94);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(107, 20);
            this.label4.TabIndex = 12;
            this.label4.Text = "指定占位符标志";
            // 
            // placeholder_comboBox
            // 
            this.placeholder_comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.placeholder_comboBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.placeholder_comboBox.FormattingEnabled = true;
            this.placeholder_comboBox.Items.AddRange(new object[] {
            "{列名}",
            "[列名]",
            "(列名)",
            "【列名】",
            "（列名）",
            "**列名**",
            "//列名//",
            "##列名##"});
            this.placeholder_comboBox.Location = new System.Drawing.Point(169, 93);
            this.placeholder_comboBox.Name = "placeholder_comboBox";
            this.placeholder_comboBox.Size = new System.Drawing.Size(302, 28);
            this.placeholder_comboBox.TabIndex = 13;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label5.Location = new System.Drawing.Point(3, 3);
            this.label5.Margin = new System.Windows.Forms.Padding(3, 3, 3, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(121, 20);
            this.label5.TabIndex = 14;
            this.label5.Text = "替换图片尺寸设置";
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.label5);
            this.flowLayoutPanel1.Controls.Add(this.pictureOriginal_comboBox);
            this.flowLayoutPanel1.Controls.Add(this.flowLayoutPanel2);
            this.flowLayoutPanel1.Controls.Add(this.flowLayoutPanel3);
            this.flowLayoutPanel1.Controls.Add(this.pictureSize_checkBox);
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(552, 55);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(134, 173);
            this.flowLayoutPanel1.TabIndex = 15;
            // 
            // pictureOriginal_comboBox
            // 
            this.pictureOriginal_comboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.pictureOriginal_comboBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.pictureOriginal_comboBox.FormattingEnabled = true;
            this.pictureOriginal_comboBox.Items.AddRange(new object[] {
            "原尺寸插入",
            "缩放插入"});
            this.pictureOriginal_comboBox.Location = new System.Drawing.Point(3, 26);
            this.pictureOriginal_comboBox.Name = "pictureOriginal_comboBox";
            this.pictureOriginal_comboBox.Size = new System.Drawing.Size(121, 28);
            this.pictureOriginal_comboBox.TabIndex = 18;
            this.pictureOriginal_comboBox.SelectedIndexChanged += new System.EventHandler(this.pictureOriginal_comboBox_SelectedIndexChanged);
            // 
            // flowLayoutPanel2
            // 
            this.flowLayoutPanel2.Controls.Add(this.label6);
            this.flowLayoutPanel2.Controls.Add(this.height_textBox);
            this.flowLayoutPanel2.Controls.Add(this.label7);
            this.flowLayoutPanel2.Location = new System.Drawing.Point(3, 60);
            this.flowLayoutPanel2.Name = "flowLayoutPanel2";
            this.flowLayoutPanel2.Size = new System.Drawing.Size(127, 33);
            this.flowLayoutPanel2.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label6.Location = new System.Drawing.Point(3, 3);
            this.label6.Margin = new System.Windows.Forms.Padding(3, 3, 3, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(23, 20);
            this.label6.TabIndex = 0;
            this.label6.Text = "高";
            // 
            // height_textBox
            // 
            this.height_textBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.height_textBox.Location = new System.Drawing.Point(32, 3);
            this.height_textBox.Name = "height_textBox";
            this.height_textBox.Size = new System.Drawing.Size(40, 26);
            this.height_textBox.TabIndex = 1;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(78, 3);
            this.label7.Margin = new System.Windows.Forms.Padding(3, 3, 3, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(37, 20);
            this.label7.TabIndex = 2;
            this.label7.Text = "厘米";
            // 
            // flowLayoutPanel3
            // 
            this.flowLayoutPanel3.Controls.Add(this.label8);
            this.flowLayoutPanel3.Controls.Add(this.width_textBox);
            this.flowLayoutPanel3.Controls.Add(this.label9);
            this.flowLayoutPanel3.Location = new System.Drawing.Point(3, 99);
            this.flowLayoutPanel3.Name = "flowLayoutPanel3";
            this.flowLayoutPanel3.Size = new System.Drawing.Size(127, 29);
            this.flowLayoutPanel3.TabIndex = 16;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label8.Location = new System.Drawing.Point(3, 3);
            this.label8.Margin = new System.Windows.Forms.Padding(3, 3, 3, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(23, 20);
            this.label8.TabIndex = 0;
            this.label8.Text = "宽";
            // 
            // width_textBox
            // 
            this.width_textBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.width_textBox.Location = new System.Drawing.Point(32, 3);
            this.width_textBox.Name = "width_textBox";
            this.width_textBox.ReadOnly = true;
            this.width_textBox.Size = new System.Drawing.Size(40, 26);
            this.width_textBox.TabIndex = 1;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(78, 3);
            this.label9.Margin = new System.Windows.Forms.Padding(3, 3, 3, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(37, 20);
            this.label9.TabIndex = 2;
            this.label9.Text = "厘米";
            // 
            // pictureSize_checkBox
            // 
            this.pictureSize_checkBox.AutoSize = true;
            this.pictureSize_checkBox.Checked = true;
            this.pictureSize_checkBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.pictureSize_checkBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.pictureSize_checkBox.Location = new System.Drawing.Point(3, 134);
            this.pictureSize_checkBox.Name = "pictureSize_checkBox";
            this.pictureSize_checkBox.Size = new System.Drawing.Size(99, 21);
            this.pictureSize_checkBox.TabIndex = 17;
            this.pictureSize_checkBox.Text = "保持原始比例";
            this.pictureSize_checkBox.UseVisualStyleBackColor = true;
            this.pictureSize_checkBox.CheckedChanged += new System.EventHandler(this.pictureSize_checkBox_CheckedChanged);
            // 
            // docQuit_button
            // 
            this.docQuit_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.docQuit_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.docQuit_button.Location = new System.Drawing.Point(489, 383);
            this.docQuit_button.Name = "docQuit_button";
            this.docQuit_button.Size = new System.Drawing.Size(75, 29);
            this.docQuit_button.TabIndex = 16;
            this.docQuit_button.Text = "退出";
            this.docQuit_button.UseVisualStyleBackColor = true;
            this.docQuit_button.Click += new System.EventHandler(this.docQuit_button_Click);
            // 
            // flowLayoutPanel4
            // 
            this.flowLayoutPanel4.Controls.Add(this.radioButtonAll);
            this.flowLayoutPanel4.Controls.Add(this.radioButtonSelected);
            this.flowLayoutPanel4.Location = new System.Drawing.Point(477, 55);
            this.flowLayoutPanel4.Name = "flowLayoutPanel4";
            this.flowLayoutPanel4.Size = new System.Drawing.Size(69, 160);
            this.flowLayoutPanel4.TabIndex = 17;
            // 
            // radioButtonAll
            // 
            this.radioButtonAll.AutoSize = true;
            this.radioButtonAll.Location = new System.Drawing.Point(3, 6);
            this.radioButtonAll.Margin = new System.Windows.Forms.Padding(3, 6, 3, 6);
            this.radioButtonAll.Name = "radioButtonAll";
            this.radioButtonAll.Size = new System.Drawing.Size(59, 16);
            this.radioButtonAll.TabIndex = 0;
            this.radioButtonAll.TabStop = true;
            this.radioButtonAll.Text = "所有行";
            this.radioButtonAll.UseVisualStyleBackColor = true;
            // 
            // radioButtonSelected
            // 
            this.radioButtonSelected.AutoSize = true;
            this.radioButtonSelected.Location = new System.Drawing.Point(3, 34);
            this.radioButtonSelected.Margin = new System.Windows.Forms.Padding(3, 6, 3, 6);
            this.radioButtonSelected.Name = "radioButtonSelected";
            this.radioButtonSelected.Size = new System.Drawing.Size(59, 16);
            this.radioButtonSelected.TabIndex = 1;
            this.radioButtonSelected.TabStop = true;
            this.radioButtonSelected.Text = "选中行";
            this.radioButtonSelected.UseVisualStyleBackColor = true;
            // 
            // Form10
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(698, 426);
            this.Controls.Add(this.flowLayoutPanel4);
            this.Controls.Add(this.docQuit_button);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.placeholder_comboBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.docReset_button);
            this.Controls.Add(this.docRun_button);
            this.Controls.Add(this.result_doc_label);
            this.Controls.Add(this.docGenerated_label);
            this.Controls.Add(this.docModel_label);
            this.Controls.Add(this.doc_folder_button);
            this.Controls.Add(this.model_select_button);
            this.Controls.Add(this.sheets_name_comboBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form10";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "文档批量生成";
            this.Load += new System.EventHandler(this.Form10_Load);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.flowLayoutPanel2.ResumeLayout(false);
            this.flowLayoutPanel2.PerformLayout();
            this.flowLayoutPanel3.ResumeLayout(false);
            this.flowLayoutPanel3.PerformLayout();
            this.flowLayoutPanel4.ResumeLayout(false);
            this.flowLayoutPanel4.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox sheets_name_comboBox;
        private System.Windows.Forms.Button model_select_button;
        private System.Windows.Forms.Button doc_folder_button;
        private System.Windows.Forms.Label docModel_label;
        private System.Windows.Forms.Label docGenerated_label;
        private System.Windows.Forms.Label result_doc_label;
        private System.Windows.Forms.Button docRun_button;
        private System.Windows.Forms.Button docReset_button;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox placeholder_comboBox;
        private System.Windows.Forms.Timer timer1;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox height_textBox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel3;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox width_textBox;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.CheckBox pictureSize_checkBox;
        private System.Windows.Forms.ComboBox pictureOriginal_comboBox;
        private System.Windows.Forms.Button docQuit_button;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel4;
        private System.Windows.Forms.RadioButton radioButtonAll;
        private System.Windows.Forms.RadioButton radioButtonSelected;
    }
}