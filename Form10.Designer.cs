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
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label1.Location = new System.Drawing.Point(223, 14);
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
            this.sheets_name_comboBox.Size = new System.Drawing.Size(352, 28);
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
            this.docModel_label.Size = new System.Drawing.Size(352, 35);
            this.docModel_label.TabIndex = 6;
            // 
            // docGenerated_label
            // 
            this.docGenerated_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.docGenerated_label.Location = new System.Drawing.Point(164, 180);
            this.docGenerated_label.Name = "docGenerated_label";
            this.docGenerated_label.Size = new System.Drawing.Size(354, 35);
            this.docGenerated_label.TabIndex = 7;
            // 
            // result_doc_label
            // 
            this.result_doc_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.result_doc_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.result_doc_label.Location = new System.Drawing.Point(65, 218);
            this.result_doc_label.Name = "result_doc_label";
            this.result_doc_label.Size = new System.Drawing.Size(456, 54);
            this.result_doc_label.TabIndex = 8;
            // 
            // docRun_button
            // 
            this.docRun_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.docRun_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.docRun_button.Location = new System.Drawing.Point(137, 381);
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
            this.docReset_button.Location = new System.Drawing.Point(323, 381);
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
            this.label3.Location = new System.Drawing.Point(65, 272);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(470, 104);
            this.label3.TabIndex = 11;
            this.label3.Text = "使用说明：\r\n1. 在excel表格中准备数据（第一行为列标题，第一列为生成的文件名）；\r\n2. 利用word编辑模板文件，需要批量修改部分用占位符+列名标记。示" +
    "例：尊敬的【客户姓名】，您的合同金额为【合同金额】元；\r\n3. 占位符可根据习惯选定使用，并在本窗口内选择对应的占位符标识；\r\n4. 选择文档批量生成的输出路径" +
    "。";
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
            this.placeholder_comboBox.Size = new System.Drawing.Size(349, 28);
            this.placeholder_comboBox.TabIndex = 13;
            // 
            // Form10
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(597, 426);
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
    }
}