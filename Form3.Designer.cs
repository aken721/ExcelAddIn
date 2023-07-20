namespace ExcelAddIn
{
    partial class Form3
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form3));
            this.label1 = new System.Windows.Forms.Label();
            this.format_radioButton1 = new System.Windows.Forms.RadioButton();
            this.format_radioButton2 = new System.Windows.Forms.RadioButton();
            this.run_button = new System.Windows.Forms.Button();
            this.quit_button = new System.Windows.Forms.Button();
            this.result_label = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.fold_path_textBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.DarkRed;
            this.label1.Location = new System.Drawing.Point(41, 73);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(208, 24);
            this.label1.TabIndex = 0;
            this.label1.Text = "请双击输入框选择文件夹";
            // 
            // format_radioButton1
            // 
            this.format_radioButton1.AutoSize = true;
            this.format_radioButton1.ForeColor = System.Drawing.Color.DarkRed;
            this.format_radioButton1.Location = new System.Drawing.Point(144, 153);
            this.format_radioButton1.Margin = new System.Windows.Forms.Padding(4);
            this.format_radioButton1.Name = "format_radioButton1";
            this.format_radioButton1.Size = new System.Drawing.Size(104, 24);
            this.format_radioButton1.TabIndex = 2;
            this.format_radioButton1.TabStop = true;
            this.format_radioButton1.Text = "歌手 - 歌名";
            this.format_radioButton1.UseVisualStyleBackColor = true;
            // 
            // format_radioButton2
            // 
            this.format_radioButton2.AutoSize = true;
            this.format_radioButton2.ForeColor = System.Drawing.Color.DarkRed;
            this.format_radioButton2.Location = new System.Drawing.Point(340, 153);
            this.format_radioButton2.Margin = new System.Windows.Forms.Padding(4);
            this.format_radioButton2.Name = "format_radioButton2";
            this.format_radioButton2.Size = new System.Drawing.Size(104, 24);
            this.format_radioButton2.TabIndex = 3;
            this.format_radioButton2.TabStop = true;
            this.format_radioButton2.Text = "歌名 - 歌手";
            this.format_radioButton2.UseVisualStyleBackColor = true;
            // 
            // run_button
            // 
            this.run_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.run_button.ForeColor = System.Drawing.Color.DarkGreen;
            this.run_button.Location = new System.Drawing.Point(116, 265);
            this.run_button.Margin = new System.Windows.Forms.Padding(4);
            this.run_button.Name = "run_button";
            this.run_button.Size = new System.Drawing.Size(88, 33);
            this.run_button.TabIndex = 4;
            this.run_button.Text = "运 行";
            this.run_button.UseVisualStyleBackColor = true;
            this.run_button.Click += new System.EventHandler(this.run_button_Click);
            // 
            // quit_button
            // 
            this.quit_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.quit_button.ForeColor = System.Drawing.Color.DarkGreen;
            this.quit_button.Location = new System.Drawing.Point(362, 265);
            this.quit_button.Margin = new System.Windows.Forms.Padding(4);
            this.quit_button.Name = "quit_button";
            this.quit_button.Size = new System.Drawing.Size(88, 33);
            this.quit_button.TabIndex = 5;
            this.quit_button.Text = "退 出";
            this.quit_button.UseVisualStyleBackColor = true;
            this.quit_button.Click += new System.EventHandler(this.quit_button_Click);
            // 
            // result_label
            // 
            this.result_label.AutoSize = true;
            this.result_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.result_label.ForeColor = System.Drawing.Color.DarkBlue;
            this.result_label.Location = new System.Drawing.Point(140, 208);
            this.result_label.Name = "result_label";
            this.result_label.Size = new System.Drawing.Size(46, 24);
            this.result_label.TabIndex = 6;
            this.result_label.Text = "结果";
            // 
            // fold_path_textBox
            // 
            this.fold_path_textBox.Location = new System.Drawing.Point(256, 73);
            this.fold_path_textBox.Name = "fold_path_textBox";
            this.fold_path_textBox.Size = new System.Drawing.Size(265, 27);
            this.fold_path_textBox.TabIndex = 7;
            this.fold_path_textBox.Click += new System.EventHandler(this.fold_path_textBox_Click);
            this.fold_path_textBox.DoubleClick += new System.EventHandler(this.fold_path_textBox_DoubleClick);
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(587, 382);
            this.Controls.Add(this.fold_path_textBox);
            this.Controls.Add(this.result_label);
            this.Controls.Add(this.quit_button);
            this.Controls.Add(this.run_button);
            this.Controls.Add(this.format_radioButton2);
            this.Controls.Add(this.format_radioButton1);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Form3";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MP3文件名批量修改";
            this.Load += new System.EventHandler(this.Form3_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton format_radioButton1;
        private System.Windows.Forms.RadioButton format_radioButton2;
        private System.Windows.Forms.Button run_button;
        private System.Windows.Forms.Button quit_button;
        private System.Windows.Forms.Label result_label;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TextBox fold_path_textBox;
    }
}