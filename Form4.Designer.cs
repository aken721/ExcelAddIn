namespace ExcelAddIn
{
    partial class Form4
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form4));
            this.file_type_checkedListBox = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.filename_regular_ComboBox1 = new System.Windows.Forms.ComboBox();
            this.filename_regular_textBox1 = new System.Windows.Forms.TextBox();
            this.regulation_add_pictureBox1 = new System.Windows.Forms.PictureBox();
            this.filename_regular_label1 = new System.Windows.Forms.Label();
            this.filename_regular_label2 = new System.Windows.Forms.Label();
            this.filename_regular_label3 = new System.Windows.Forms.Label();
            this.filename_regular_ComboBox2 = new System.Windows.Forms.ComboBox();
            this.filename_regular_ComboBox3 = new System.Windows.Forms.ComboBox();
            this.filename_regular_textBox2 = new System.Windows.Forms.TextBox();
            this.filename_regular_textBox3 = new System.Windows.Forms.TextBox();
            this.regulation_reduce_pictureBox2 = new System.Windows.Forms.PictureBox();
            this.regulation_add_pictureBox2 = new System.Windows.Forms.PictureBox();
            this.regulation_reduce_pictureBox3 = new System.Windows.Forms.PictureBox();
            this.run_button = new System.Windows.Forms.Button();
            this.reset_button = new System.Windows.Forms.Button();
            this.quit_button = new System.Windows.Forms.Button();
            this.move_select_radioButton = new System.Windows.Forms.RadioButton();
            this.delete_select_radioButton = new System.Windows.Forms.RadioButton();
            this.del_mov_title_label = new System.Windows.Forms.Label();
            this.tips_label = new System.Windows.Forms.Label();
            this.result_dm_label = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            ((System.ComponentModel.ISupportInitialize)(this.regulation_add_pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.regulation_reduce_pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.regulation_add_pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.regulation_reduce_pictureBox3)).BeginInit();
            this.SuspendLayout();
            // 
            // file_type_checkedListBox
            // 
            this.file_type_checkedListBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.file_type_checkedListBox.FormattingEnabled = true;
            this.file_type_checkedListBox.Location = new System.Drawing.Point(168, 131);
            this.file_type_checkedListBox.Name = "file_type_checkedListBox";
            this.file_type_checkedListBox.Size = new System.Drawing.Size(168, 130);
            this.file_type_checkedListBox.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label1.Location = new System.Drawing.Point(55, 133);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(107, 20);
            this.label1.TabIndex = 1;
            this.label1.Text = "请选择文件类型";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label2.Location = new System.Drawing.Point(400, 133);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(219, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "请定义文件名规则（不含扩展名）";
            // 
            // filename_regular_ComboBox1
            // 
            this.filename_regular_ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.filename_regular_ComboBox1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.filename_regular_ComboBox1.FormattingEnabled = true;
            this.filename_regular_ComboBox1.Items.AddRange(new object[] {
            "始于",
            "止于",
            "包含",
            "不包含"});
            this.filename_regular_ComboBox1.Location = new System.Drawing.Point(457, 172);
            this.filename_regular_ComboBox1.Name = "filename_regular_ComboBox1";
            this.filename_regular_ComboBox1.Size = new System.Drawing.Size(70, 28);
            this.filename_regular_ComboBox1.TabIndex = 3;
            // 
            // filename_regular_textBox1
            // 
            this.filename_regular_textBox1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.filename_regular_textBox1.Location = new System.Drawing.Point(533, 172);
            this.filename_regular_textBox1.Name = "filename_regular_textBox1";
            this.filename_regular_textBox1.Size = new System.Drawing.Size(156, 26);
            this.filename_regular_textBox1.TabIndex = 4;
            // 
            // regulation_add_pictureBox1
            // 
            this.regulation_add_pictureBox1.Image = global::ExcelAddIn.Properties.Resources.add;
            this.regulation_add_pictureBox1.Location = new System.Drawing.Point(706, 172);
            this.regulation_add_pictureBox1.Name = "regulation_add_pictureBox1";
            this.regulation_add_pictureBox1.Size = new System.Drawing.Size(20, 20);
            this.regulation_add_pictureBox1.TabIndex = 5;
            this.regulation_add_pictureBox1.TabStop = false;
            this.regulation_add_pictureBox1.Click += new System.EventHandler(this.regulation_add_pictureBox1_Click);
            // 
            // filename_regular_label1
            // 
            this.filename_regular_label1.AutoSize = true;
            this.filename_regular_label1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.filename_regular_label1.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.filename_regular_label1.Location = new System.Drawing.Point(406, 172);
            this.filename_regular_label1.Name = "filename_regular_label1";
            this.filename_regular_label1.Size = new System.Drawing.Size(45, 20);
            this.filename_regular_label1.TabIndex = 7;
            this.filename_regular_label1.Text = "规则1";
            // 
            // filename_regular_label2
            // 
            this.filename_regular_label2.AutoSize = true;
            this.filename_regular_label2.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.filename_regular_label2.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.filename_regular_label2.Location = new System.Drawing.Point(406, 226);
            this.filename_regular_label2.Name = "filename_regular_label2";
            this.filename_regular_label2.Size = new System.Drawing.Size(45, 20);
            this.filename_regular_label2.TabIndex = 8;
            this.filename_regular_label2.Text = "规则2";
            // 
            // filename_regular_label3
            // 
            this.filename_regular_label3.AutoSize = true;
            this.filename_regular_label3.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.filename_regular_label3.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.filename_regular_label3.Location = new System.Drawing.Point(406, 280);
            this.filename_regular_label3.Name = "filename_regular_label3";
            this.filename_regular_label3.Size = new System.Drawing.Size(45, 20);
            this.filename_regular_label3.TabIndex = 9;
            this.filename_regular_label3.Text = "规则3";
            // 
            // filename_regular_ComboBox2
            // 
            this.filename_regular_ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.filename_regular_ComboBox2.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.filename_regular_ComboBox2.FormattingEnabled = true;
            this.filename_regular_ComboBox2.Items.AddRange(new object[] {
            "始于",
            "止于",
            "包含",
            "不包含"});
            this.filename_regular_ComboBox2.Location = new System.Drawing.Point(457, 226);
            this.filename_regular_ComboBox2.Name = "filename_regular_ComboBox2";
            this.filename_regular_ComboBox2.Size = new System.Drawing.Size(70, 28);
            this.filename_regular_ComboBox2.TabIndex = 10;
            // 
            // filename_regular_ComboBox3
            // 
            this.filename_regular_ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.filename_regular_ComboBox3.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.filename_regular_ComboBox3.FormattingEnabled = true;
            this.filename_regular_ComboBox3.Items.AddRange(new object[] {
            "始于",
            "止于",
            "包含",
            "不包含"});
            this.filename_regular_ComboBox3.Location = new System.Drawing.Point(457, 280);
            this.filename_regular_ComboBox3.Name = "filename_regular_ComboBox3";
            this.filename_regular_ComboBox3.Size = new System.Drawing.Size(70, 28);
            this.filename_regular_ComboBox3.TabIndex = 11;
            // 
            // filename_regular_textBox2
            // 
            this.filename_regular_textBox2.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.filename_regular_textBox2.Location = new System.Drawing.Point(533, 226);
            this.filename_regular_textBox2.Name = "filename_regular_textBox2";
            this.filename_regular_textBox2.Size = new System.Drawing.Size(156, 26);
            this.filename_regular_textBox2.TabIndex = 12;
            // 
            // filename_regular_textBox3
            // 
            this.filename_regular_textBox3.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.filename_regular_textBox3.Location = new System.Drawing.Point(533, 280);
            this.filename_regular_textBox3.Name = "filename_regular_textBox3";
            this.filename_regular_textBox3.Size = new System.Drawing.Size(156, 26);
            this.filename_regular_textBox3.TabIndex = 13;
            // 
            // regulation_reduce_pictureBox2
            // 
            this.regulation_reduce_pictureBox2.Image = global::ExcelAddIn.Properties.Resources.reduce;
            this.regulation_reduce_pictureBox2.Location = new System.Drawing.Point(735, 226);
            this.regulation_reduce_pictureBox2.Name = "regulation_reduce_pictureBox2";
            this.regulation_reduce_pictureBox2.Size = new System.Drawing.Size(20, 20);
            this.regulation_reduce_pictureBox2.TabIndex = 15;
            this.regulation_reduce_pictureBox2.TabStop = false;
            this.regulation_reduce_pictureBox2.Click += new System.EventHandler(this.regulation_reduce_pictureBox2_Click);
            // 
            // regulation_add_pictureBox2
            // 
            this.regulation_add_pictureBox2.Image = global::ExcelAddIn.Properties.Resources.add;
            this.regulation_add_pictureBox2.Location = new System.Drawing.Point(706, 226);
            this.regulation_add_pictureBox2.Name = "regulation_add_pictureBox2";
            this.regulation_add_pictureBox2.Size = new System.Drawing.Size(20, 20);
            this.regulation_add_pictureBox2.TabIndex = 14;
            this.regulation_add_pictureBox2.TabStop = false;
            this.regulation_add_pictureBox2.Click += new System.EventHandler(this.regulation_add_pictureBox2_Click);
            // 
            // regulation_reduce_pictureBox3
            // 
            this.regulation_reduce_pictureBox3.Image = global::ExcelAddIn.Properties.Resources.reduce;
            this.regulation_reduce_pictureBox3.Location = new System.Drawing.Point(706, 280);
            this.regulation_reduce_pictureBox3.Name = "regulation_reduce_pictureBox3";
            this.regulation_reduce_pictureBox3.Size = new System.Drawing.Size(20, 20);
            this.regulation_reduce_pictureBox3.TabIndex = 17;
            this.regulation_reduce_pictureBox3.TabStop = false;
            this.regulation_reduce_pictureBox3.Click += new System.EventHandler(this.regulation_reduce_pictureBox3_Click);
            // 
            // run_button
            // 
            this.run_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.run_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.run_button.Location = new System.Drawing.Point(213, 389);
            this.run_button.Name = "run_button";
            this.run_button.Size = new System.Drawing.Size(76, 33);
            this.run_button.TabIndex = 18;
            this.run_button.Text = "运行";
            this.run_button.UseVisualStyleBackColor = true;
            this.run_button.Click += new System.EventHandler(this.run_button_Click);
            // 
            // reset_button
            // 
            this.reset_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.reset_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.reset_button.Location = new System.Drawing.Point(373, 389);
            this.reset_button.Name = "reset_button";
            this.reset_button.Size = new System.Drawing.Size(76, 33);
            this.reset_button.TabIndex = 19;
            this.reset_button.Text = "重置";
            this.reset_button.UseVisualStyleBackColor = true;
            this.reset_button.Click += new System.EventHandler(this.reset_button_Click);
            // 
            // quit_button
            // 
            this.quit_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.quit_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.quit_button.Location = new System.Drawing.Point(533, 389);
            this.quit_button.Name = "quit_button";
            this.quit_button.Size = new System.Drawing.Size(76, 33);
            this.quit_button.TabIndex = 20;
            this.quit_button.Text = "退出";
            this.quit_button.UseVisualStyleBackColor = true;
            this.quit_button.Click += new System.EventHandler(this.quit_button_Click);
            // 
            // move_select_radioButton
            // 
            this.move_select_radioButton.AutoSize = true;
            this.move_select_radioButton.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.move_select_radioButton.ForeColor = System.Drawing.Color.DarkMagenta;
            this.move_select_radioButton.Location = new System.Drawing.Point(414, 84);
            this.move_select_radioButton.Name = "move_select_radioButton";
            this.move_select_radioButton.Size = new System.Drawing.Size(55, 23);
            this.move_select_radioButton.TabIndex = 21;
            this.move_select_radioButton.TabStop = true;
            this.move_select_radioButton.Text = "移动";
            this.move_select_radioButton.UseVisualStyleBackColor = true;
            // 
            // delete_select_radioButton
            // 
            this.delete_select_radioButton.AutoSize = true;
            this.delete_select_radioButton.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.delete_select_radioButton.ForeColor = System.Drawing.Color.DarkMagenta;
            this.delete_select_radioButton.Location = new System.Drawing.Point(328, 84);
            this.delete_select_radioButton.Name = "delete_select_radioButton";
            this.delete_select_radioButton.Size = new System.Drawing.Size(55, 23);
            this.delete_select_radioButton.TabIndex = 22;
            this.delete_select_radioButton.TabStop = true;
            this.delete_select_radioButton.Text = "删除";
            this.delete_select_radioButton.UseVisualStyleBackColor = true;
            // 
            // del_mov_title_label
            // 
            this.del_mov_title_label.AutoSize = true;
            this.del_mov_title_label.Font = new System.Drawing.Font("微软雅黑", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.del_mov_title_label.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.del_mov_title_label.Location = new System.Drawing.Point(323, 25);
            this.del_mov_title_label.Name = "del_mov_title_label";
            this.del_mov_title_label.Size = new System.Drawing.Size(159, 28);
            this.del_mov_title_label.TabIndex = 23;
            this.del_mov_title_label.Text = "批量删除或移动";
            // 
            // tips_label
            // 
            this.tips_label.AutoSize = true;
            this.tips_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tips_label.ForeColor = System.Drawing.Color.Teal;
            this.tips_label.Location = new System.Drawing.Point(46, 283);
            this.tips_label.Name = "tips_label";
            this.tips_label.Size = new System.Drawing.Size(290, 68);
            this.tips_label.TabIndex = 24;
            this.tips_label.Text = "注意：1. 文件名和文件类型为逻辑“与”规则，即同\r\n             时满足。\r\n         2. 文件名规则内部为逻辑“或”设计， 满足单\r\n " +
    "            条即执行。";
            // 
            // result_dm_label
            // 
            this.result_dm_label.AutoSize = true;
            this.result_dm_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.result_dm_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.result_dm_label.Location = new System.Drawing.Point(415, 331);
            this.result_dm_label.Name = "result_dm_label";
            this.result_dm_label.Size = new System.Drawing.Size(0, 20);
            this.result_dm_label.TabIndex = 25;
            // 
            // Form4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.result_dm_label);
            this.Controls.Add(this.tips_label);
            this.Controls.Add(this.del_mov_title_label);
            this.Controls.Add(this.delete_select_radioButton);
            this.Controls.Add(this.move_select_radioButton);
            this.Controls.Add(this.quit_button);
            this.Controls.Add(this.reset_button);
            this.Controls.Add(this.run_button);
            this.Controls.Add(this.regulation_reduce_pictureBox3);
            this.Controls.Add(this.regulation_reduce_pictureBox2);
            this.Controls.Add(this.regulation_add_pictureBox2);
            this.Controls.Add(this.filename_regular_textBox3);
            this.Controls.Add(this.filename_regular_textBox2);
            this.Controls.Add(this.filename_regular_ComboBox3);
            this.Controls.Add(this.filename_regular_ComboBox2);
            this.Controls.Add(this.filename_regular_label3);
            this.Controls.Add(this.filename_regular_label2);
            this.Controls.Add(this.filename_regular_label1);
            this.Controls.Add(this.regulation_add_pictureBox1);
            this.Controls.Add(this.filename_regular_textBox1);
            this.Controls.Add(this.filename_regular_ComboBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.file_type_checkedListBox);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form4";
            this.Text = "删除或移动文件";
            this.Load += new System.EventHandler(this.Form4_Load);
            ((System.ComponentModel.ISupportInitialize)(this.regulation_add_pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.regulation_reduce_pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.regulation_add_pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.regulation_reduce_pictureBox3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox file_type_checkedListBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox filename_regular_ComboBox1;
        private System.Windows.Forms.TextBox filename_regular_textBox1;
        private System.Windows.Forms.PictureBox regulation_add_pictureBox1;
        private System.Windows.Forms.Label filename_regular_label1;
        private System.Windows.Forms.Label filename_regular_label2;
        private System.Windows.Forms.Label filename_regular_label3;
        private System.Windows.Forms.ComboBox filename_regular_ComboBox2;
        private System.Windows.Forms.ComboBox filename_regular_ComboBox3;
        private System.Windows.Forms.TextBox filename_regular_textBox2;
        private System.Windows.Forms.TextBox filename_regular_textBox3;
        private System.Windows.Forms.PictureBox regulation_reduce_pictureBox2;
        private System.Windows.Forms.PictureBox regulation_add_pictureBox2;
        private System.Windows.Forms.PictureBox regulation_reduce_pictureBox3;
        private System.Windows.Forms.Button run_button;
        private System.Windows.Forms.Button reset_button;
        private System.Windows.Forms.Button quit_button;
        private System.Windows.Forms.RadioButton move_select_radioButton;
        private System.Windows.Forms.RadioButton delete_select_radioButton;
        private System.Windows.Forms.Label del_mov_title_label;
        private System.Windows.Forms.Label tips_label;
        private System.Windows.Forms.Label result_dm_label;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}