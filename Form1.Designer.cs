namespace ExcelAddIn
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.splitProgressBar_label = new System.Windows.Forms.Label();
            this.clear_button = new System.Windows.Forms.Button();
            this.splitsheet_delete_button = new System.Windows.Forms.Button();
            this.splitsheet_export_button = new System.Windows.Forms.Button();
            this.split_sheet_progressBar = new System.Windows.Forms.ProgressBar();
            this.split_button = new System.Windows.Forms.Button();
            this.split_sheet_result_label = new System.Windows.Forms.Label();
            this.field_name_combobox = new System.Windows.Forms.ComboBox();
            this.sheet_name_combobox = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.single_merge_button = new System.Windows.Forms.Button();
            this.mergeProgressBar_label = new System.Windows.Forms.Label();
            this.merge_sheet_progressBar = new System.Windows.Forms.ProgressBar();
            this.merge_sheet_result_label = new System.Windows.Forms.Label();
            this.multi_merge_sheet_checkBox = new System.Windows.Forms.CheckBox();
            this.dir_select_button = new System.Windows.Forms.Button();
            this.multi_merge_button = new System.Windows.Forms.Button();
            this.dir_select_textbox = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.batch_delete_button = new System.Windows.Forms.Button();
            this.batch_export_button = new System.Windows.Forms.Button();
            this.all_select_checkbox = new System.Windows.Forms.CheckBox();
            this.sheet_listbox = new System.Windows.Forms.ListBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.contents_button = new System.Windows.Forms.Button();
            this.payslip_button = new System.Windows.Forms.Button();
            this.regex_button = new System.Windows.Forms.Button();
            this.transposition_button = new System.Windows.Forms.Button();
            this.add_sheet_button = new System.Windows.Forms.Button();
            this.move_sheet_button = new System.Windows.Forms.Button();
            this.run_result_label = new System.Windows.Forms.Label();
            this.regex_clear_button = new System.Windows.Forms.Button();
            this.regex_run_button = new System.Windows.Forms.Button();
            this.regex_rule_textbox = new System.Windows.Forms.TextBox();
            this.regex_rule_label = new System.Windows.Forms.Label();
            this.what_type_combobox = new System.Windows.Forms.ComboBox();
            this.what_type_label = new System.Windows.Forms.Label();
            this.which_field_combobox = new System.Windows.Forms.ComboBox();
            this.which_field_label = new System.Windows.Forms.Label();
            this.function_title_label = new System.Windows.Forms.Label();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.tabPage6 = new System.Windows.Forms.TabPage();
            this.split_sheet_timer = new System.Windows.Forms.Timer(this.components);
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.merge_sheet_timer = new System.Windows.Forms.Timer(this.components);
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.tabPage5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabControl1
            // 
            this.tabControl1.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Controls.Add(this.tabPage5);
            this.tabControl1.Controls.Add(this.tabPage6);
            this.tabControl1.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed;
            this.tabControl1.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tabControl1.HotTrack = true;
            this.tabControl1.ItemSize = new System.Drawing.Size(50, 120);
            this.tabControl1.Location = new System.Drawing.Point(0, 1);
            this.tabControl1.Multiline = true;
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(803, 400);
            this.tabControl1.SizeMode = System.Windows.Forms.TabSizeMode.Fixed;
            this.tabControl1.TabIndex = 0;
            this.tabControl1.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.tabControl1_DrawItem);
            this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.splitProgressBar_label);
            this.tabPage1.Controls.Add(this.clear_button);
            this.tabPage1.Controls.Add(this.splitsheet_delete_button);
            this.tabPage1.Controls.Add(this.splitsheet_export_button);
            this.tabPage1.Controls.Add(this.split_sheet_progressBar);
            this.tabPage1.Controls.Add(this.split_button);
            this.tabPage1.Controls.Add(this.split_sheet_result_label);
            this.tabPage1.Controls.Add(this.field_name_combobox);
            this.tabPage1.Controls.Add(this.sheet_name_combobox);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tabPage1.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.tabPage1.Location = new System.Drawing.Point(124, 4);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(675, 392);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "一、分表功能";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // splitProgressBar_label
            // 
            this.splitProgressBar_label.AutoSize = true;
            this.splitProgressBar_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.splitProgressBar_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.splitProgressBar_label.Location = new System.Drawing.Point(88, 179);
            this.splitProgressBar_label.Name = "splitProgressBar_label";
            this.splitProgressBar_label.Size = new System.Drawing.Size(65, 20);
            this.splitProgressBar_label.TabIndex = 11;
            this.splitProgressBar_label.Text = "分表进度";
            // 
            // clear_button
            // 
            this.clear_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.clear_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.clear_button.Location = new System.Drawing.Point(547, 126);
            this.clear_button.Name = "clear_button";
            this.clear_button.Size = new System.Drawing.Size(45, 27);
            this.clear_button.TabIndex = 10;
            this.clear_button.Text = "清空";
            this.clear_button.UseVisualStyleBackColor = true;
            this.clear_button.Click += new System.EventHandler(this.clear_button_Click);
            // 
            // splitsheet_delete_button
            // 
            this.splitsheet_delete_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.splitsheet_delete_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.splitsheet_delete_button.Location = new System.Drawing.Point(464, 348);
            this.splitsheet_delete_button.Name = "splitsheet_delete_button";
            this.splitsheet_delete_button.Size = new System.Drawing.Size(76, 32);
            this.splitsheet_delete_button.TabIndex = 9;
            this.splitsheet_delete_button.Text = "删除分表";
            this.splitsheet_delete_button.UseVisualStyleBackColor = true;
            this.splitsheet_delete_button.Click += new System.EventHandler(this.splitsheet_delete_button_Click);
            // 
            // splitsheet_export_button
            // 
            this.splitsheet_export_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.splitsheet_export_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.splitsheet_export_button.Location = new System.Drawing.Point(296, 348);
            this.splitsheet_export_button.Name = "splitsheet_export_button";
            this.splitsheet_export_button.Size = new System.Drawing.Size(76, 32);
            this.splitsheet_export_button.TabIndex = 8;
            this.splitsheet_export_button.Text = "分表导出";
            this.splitsheet_export_button.UseVisualStyleBackColor = true;
            this.splitsheet_export_button.Click += new System.EventHandler(this.splitsheet_export_button_Click);
            // 
            // split_sheet_progressBar
            // 
            this.split_sheet_progressBar.Location = new System.Drawing.Point(200, 184);
            this.split_sheet_progressBar.Name = "split_sheet_progressBar";
            this.split_sheet_progressBar.Size = new System.Drawing.Size(340, 15);
            this.split_sheet_progressBar.TabIndex = 7;
            // 
            // split_button
            // 
            this.split_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.split_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.split_button.Location = new System.Drawing.Point(122, 348);
            this.split_button.Name = "split_button";
            this.split_button.Size = new System.Drawing.Size(76, 32);
            this.split_button.TabIndex = 6;
            this.split_button.Text = "分表";
            this.split_button.UseVisualStyleBackColor = true;
            this.split_button.Click += new System.EventHandler(this.split_button_Click);
            // 
            // split_sheet_result_label
            // 
            this.split_sheet_result_label.AutoSize = true;
            this.split_sheet_result_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.split_sheet_result_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.split_sheet_result_label.Location = new System.Drawing.Point(88, 214);
            this.split_sheet_result_label.Name = "split_sheet_result_label";
            this.split_sheet_result_label.Size = new System.Drawing.Size(65, 20);
            this.split_sheet_result_label.TabIndex = 5;
            this.split_sheet_result_label.Text = "分表结果";
            // 
            // field_name_combobox
            // 
            this.field_name_combobox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.field_name_combobox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.field_name_combobox.FormattingEnabled = true;
            this.field_name_combobox.Location = new System.Drawing.Point(263, 126);
            this.field_name_combobox.Name = "field_name_combobox";
            this.field_name_combobox.Size = new System.Drawing.Size(277, 28);
            this.field_name_combobox.TabIndex = 4;
            // 
            // sheet_name_combobox
            // 
            this.sheet_name_combobox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.sheet_name_combobox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.sheet_name_combobox.FormattingEnabled = true;
            this.sheet_name_combobox.Location = new System.Drawing.Point(263, 68);
            this.sheet_name_combobox.Name = "sheet_name_combobox";
            this.sheet_name_combobox.Size = new System.Drawing.Size(277, 28);
            this.sheet_name_combobox.TabIndex = 3;
            this.sheet_name_combobox.TextChanged += new System.EventHandler(this.sheet_name_combobox_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.ForeColor = System.Drawing.Color.Teal;
            this.label3.Location = new System.Drawing.Point(88, 250);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(470, 76);
            this.label3.TabIndex = 2;
            this.label3.Text = "注意：\r\n1.分表功能是基于每次所分的表进行操作，所以导出和删除均只能在运行分\r\n表后才有效；\r\n2.如需在不运行分表钱导出或删除，请用导出删除页相关功能。";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(88, 126);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(163, 20);
            this.label2.TabIndex = 1;
            this.label2.Text = "请选择分表所依据的字段";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(88, 68);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "请选择要分的原始表";
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.splitContainer1);
            this.tabPage2.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tabPage2.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.tabPage2.Location = new System.Drawing.Point(124, 4);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(675, 392);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "二、并表功能";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Location = new System.Drawing.Point(3, 3);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.single_merge_button);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.mergeProgressBar_label);
            this.splitContainer1.Panel2.Controls.Add(this.merge_sheet_progressBar);
            this.splitContainer1.Panel2.Controls.Add(this.merge_sheet_result_label);
            this.splitContainer1.Panel2.Controls.Add(this.multi_merge_sheet_checkBox);
            this.splitContainer1.Panel2.Controls.Add(this.dir_select_button);
            this.splitContainer1.Panel2.Controls.Add(this.multi_merge_button);
            this.splitContainer1.Panel2.Controls.Add(this.dir_select_textbox);
            this.splitContainer1.Panel2.Controls.Add(this.label5);
            this.splitContainer1.Size = new System.Drawing.Size(669, 386);
            this.splitContainer1.SplitterDistance = 165;
            this.splitContainer1.TabIndex = 0;
            // 
            // single_merge_button
            // 
            this.single_merge_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.single_merge_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.single_merge_button.Location = new System.Drawing.Point(256, 37);
            this.single_merge_button.Name = "single_merge_button";
            this.single_merge_button.Size = new System.Drawing.Size(120, 61);
            this.single_merge_button.TabIndex = 0;
            this.single_merge_button.Text = "并表\r\n（同一工作簿）";
            this.single_merge_button.UseVisualStyleBackColor = true;
            this.single_merge_button.Click += new System.EventHandler(this.single_merge_button_Click);
            // 
            // mergeProgressBar_label
            // 
            this.mergeProgressBar_label.AutoSize = true;
            this.mergeProgressBar_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mergeProgressBar_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.mergeProgressBar_label.Location = new System.Drawing.Point(50, 87);
            this.mergeProgressBar_label.Name = "mergeProgressBar_label";
            this.mergeProgressBar_label.Size = new System.Drawing.Size(65, 20);
            this.mergeProgressBar_label.TabIndex = 8;
            this.mergeProgressBar_label.Text = "完成进度";
            // 
            // merge_sheet_progressBar
            // 
            this.merge_sheet_progressBar.ForeColor = System.Drawing.Color.Purple;
            this.merge_sheet_progressBar.Location = new System.Drawing.Point(139, 94);
            this.merge_sheet_progressBar.Name = "merge_sheet_progressBar";
            this.merge_sheet_progressBar.Size = new System.Drawing.Size(385, 13);
            this.merge_sheet_progressBar.TabIndex = 7;
            // 
            // merge_sheet_result_label
            // 
            this.merge_sheet_result_label.AutoSize = true;
            this.merge_sheet_result_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.merge_sheet_result_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.merge_sheet_result_label.Location = new System.Drawing.Point(50, 124);
            this.merge_sheet_result_label.Name = "merge_sheet_result_label";
            this.merge_sheet_result_label.Size = new System.Drawing.Size(65, 20);
            this.merge_sheet_result_label.TabIndex = 6;
            this.merge_sheet_result_label.Text = "并表结果";
            // 
            // multi_merge_sheet_checkBox
            // 
            this.multi_merge_sheet_checkBox.AutoSize = true;
            this.multi_merge_sheet_checkBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.multi_merge_sheet_checkBox.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.multi_merge_sheet_checkBox.Location = new System.Drawing.Point(90, 204);
            this.multi_merge_sheet_checkBox.Name = "multi_merge_sheet_checkBox";
            this.multi_merge_sheet_checkBox.Size = new System.Drawing.Size(434, 24);
            this.multi_merge_sheet_checkBox.TabIndex = 5;
            this.multi_merge_sheet_checkBox.Text = "当前已打开工作簿的已有表列入合并范围（仅适用多工作簿合并）";
            this.multi_merge_sheet_checkBox.UseVisualStyleBackColor = true;
            // 
            // dir_select_button
            // 
            this.dir_select_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dir_select_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.dir_select_button.Location = new System.Drawing.Point(465, 47);
            this.dir_select_button.Name = "dir_select_button";
            this.dir_select_button.Size = new System.Drawing.Size(59, 26);
            this.dir_select_button.TabIndex = 4;
            this.dir_select_button.Text = "选择";
            this.dir_select_button.UseVisualStyleBackColor = true;
            this.dir_select_button.Click += new System.EventHandler(this.dir_select_button_Click);
            // 
            // multi_merge_button
            // 
            this.multi_merge_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.multi_merge_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.multi_merge_button.Location = new System.Drawing.Point(256, 152);
            this.multi_merge_button.Name = "multi_merge_button";
            this.multi_merge_button.Size = new System.Drawing.Size(111, 34);
            this.multi_merge_button.TabIndex = 2;
            this.multi_merge_button.Text = "工作簿汇总";
            this.multi_merge_button.UseVisualStyleBackColor = true;
            this.multi_merge_button.Click += new System.EventHandler(this.multi_merge_button_Click);
            // 
            // dir_select_textbox
            // 
            this.dir_select_textbox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dir_select_textbox.Location = new System.Drawing.Point(217, 47);
            this.dir_select_textbox.Name = "dir_select_textbox";
            this.dir_select_textbox.ReadOnly = true;
            this.dir_select_textbox.Size = new System.Drawing.Size(241, 26);
            this.dir_select_textbox.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(50, 51);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(163, 20);
            this.label5.TabIndex = 0;
            this.label5.Text = "请选择工作簿所在文件夹";
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.batch_delete_button);
            this.tabPage3.Controls.Add(this.batch_export_button);
            this.tabPage3.Controls.Add(this.all_select_checkbox);
            this.tabPage3.Controls.Add(this.sheet_listbox);
            this.tabPage3.Controls.Add(this.label6);
            this.tabPage3.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tabPage3.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.tabPage3.Location = new System.Drawing.Point(124, 4);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(675, 392);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "三、批量导删";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // batch_delete_button
            // 
            this.batch_delete_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.batch_delete_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.batch_delete_button.Location = new System.Drawing.Point(362, 282);
            this.batch_delete_button.Name = "batch_delete_button";
            this.batch_delete_button.Size = new System.Drawing.Size(89, 32);
            this.batch_delete_button.TabIndex = 4;
            this.batch_delete_button.Text = "删除所选表";
            this.batch_delete_button.UseVisualStyleBackColor = true;
            this.batch_delete_button.Click += new System.EventHandler(this.batch_delete_button_Click);
            // 
            // batch_export_button
            // 
            this.batch_export_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.batch_export_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.batch_export_button.Location = new System.Drawing.Point(163, 282);
            this.batch_export_button.Name = "batch_export_button";
            this.batch_export_button.Size = new System.Drawing.Size(89, 32);
            this.batch_export_button.TabIndex = 3;
            this.batch_export_button.Text = "导出所选表";
            this.batch_export_button.UseVisualStyleBackColor = true;
            this.batch_export_button.Click += new System.EventHandler(this.batch_export_button_Click);
            // 
            // all_select_checkbox
            // 
            this.all_select_checkbox.AutoSize = true;
            this.all_select_checkbox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.all_select_checkbox.Location = new System.Drawing.Point(235, 224);
            this.all_select_checkbox.Name = "all_select_checkbox";
            this.all_select_checkbox.Size = new System.Drawing.Size(84, 24);
            this.all_select_checkbox.TabIndex = 2;
            this.all_select_checkbox.Text = "全部选中";
            this.all_select_checkbox.UseVisualStyleBackColor = true;
            this.all_select_checkbox.Click += new System.EventHandler(this.all_select_checkbox_Click);
            // 
            // sheet_listbox
            // 
            this.sheet_listbox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.sheet_listbox.FormattingEnabled = true;
            this.sheet_listbox.ItemHeight = 20;
            this.sheet_listbox.Location = new System.Drawing.Point(235, 98);
            this.sheet_listbox.Name = "sheet_listbox";
            this.sheet_listbox.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.sheet_listbox.Size = new System.Drawing.Size(235, 64);
            this.sheet_listbox.TabIndex = 1;
            this.sheet_listbox.SelectedValueChanged += new System.EventHandler(this.sheet_listbox_SelectedValueChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(136, 98);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(93, 40);
            this.label6.TabIndex = 0;
            this.label6.Text = "选择要导出\r\n或删除的表：";
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.splitContainer2);
            this.tabPage4.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tabPage4.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.tabPage4.Location = new System.Drawing.Point(124, 4);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Size = new System.Drawing.Size(675, 392);
            this.tabPage4.TabIndex = 6;
            this.tabPage4.Text = "四、实用功能";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.splitContainer2.IsSplitterFixed = true;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.contents_button);
            this.splitContainer2.Panel1.Controls.Add(this.payslip_button);
            this.splitContainer2.Panel1.Controls.Add(this.regex_button);
            this.splitContainer2.Panel1.Controls.Add(this.transposition_button);
            this.splitContainer2.Panel1.Controls.Add(this.add_sheet_button);
            this.splitContainer2.Panel1.Controls.Add(this.move_sheet_button);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.run_result_label);
            this.splitContainer2.Panel2.Controls.Add(this.regex_clear_button);
            this.splitContainer2.Panel2.Controls.Add(this.regex_run_button);
            this.splitContainer2.Panel2.Controls.Add(this.regex_rule_textbox);
            this.splitContainer2.Panel2.Controls.Add(this.regex_rule_label);
            this.splitContainer2.Panel2.Controls.Add(this.what_type_combobox);
            this.splitContainer2.Panel2.Controls.Add(this.what_type_label);
            this.splitContainer2.Panel2.Controls.Add(this.which_field_combobox);
            this.splitContainer2.Panel2.Controls.Add(this.which_field_label);
            this.splitContainer2.Panel2.Controls.Add(this.function_title_label);
            this.splitContainer2.Size = new System.Drawing.Size(675, 392);
            this.splitContainer2.SplitterDistance = 318;
            this.splitContainer2.TabIndex = 0;
            // 
            // contents_button
            // 
            this.contents_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.contents_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.contents_button.Location = new System.Drawing.Point(185, 288);
            this.contents_button.Name = "contents_button";
            this.contents_button.Size = new System.Drawing.Size(90, 50);
            this.contents_button.TabIndex = 5;
            this.contents_button.Text = "建立目录页新表";
            this.contents_button.UseVisualStyleBackColor = true;
            this.contents_button.Click += new System.EventHandler(this.contents_button_Click);
            // 
            // payslip_button
            // 
            this.payslip_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.payslip_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.payslip_button.Location = new System.Drawing.Point(45, 289);
            this.payslip_button.Name = "payslip_button";
            this.payslip_button.Size = new System.Drawing.Size(90, 50);
            this.payslip_button.TabIndex = 4;
            this.payslip_button.Text = "一键生成工资条";
            this.payslip_button.UseVisualStyleBackColor = true;
            this.payslip_button.Click += new System.EventHandler(this.payslip_button_Click);
            // 
            // regex_button
            // 
            this.regex_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.regex_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.regex_button.Location = new System.Drawing.Point(185, 182);
            this.regex_button.Name = "regex_button";
            this.regex_button.Size = new System.Drawing.Size(90, 50);
            this.regex_button.TabIndex = 3;
            this.regex_button.Text = "正则表达式";
            this.regex_button.UseVisualStyleBackColor = true;
            this.regex_button.Click += new System.EventHandler(this.regex_button_Click);
            // 
            // transposition_button
            // 
            this.transposition_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.transposition_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.transposition_button.Location = new System.Drawing.Point(45, 183);
            this.transposition_button.Name = "transposition_button";
            this.transposition_button.Size = new System.Drawing.Size(90, 50);
            this.transposition_button.TabIndex = 2;
            this.transposition_button.Text = "转置工作表";
            this.transposition_button.UseVisualStyleBackColor = true;
            this.transposition_button.Click += new System.EventHandler(this.transposition_button_Click);
            // 
            // add_sheet_button
            // 
            this.add_sheet_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.add_sheet_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.add_sheet_button.Location = new System.Drawing.Point(185, 82);
            this.add_sheet_button.Name = "add_sheet_button";
            this.add_sheet_button.Size = new System.Drawing.Size(90, 50);
            this.add_sheet_button.TabIndex = 1;
            this.add_sheet_button.Text = "一键建立多个工作表";
            this.add_sheet_button.UseVisualStyleBackColor = true;
            this.add_sheet_button.Click += new System.EventHandler(this.add_sheet_button_Click);
            // 
            // move_sheet_button
            // 
            this.move_sheet_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.move_sheet_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.move_sheet_button.Location = new System.Drawing.Point(45, 82);
            this.move_sheet_button.Name = "move_sheet_button";
            this.move_sheet_button.Size = new System.Drawing.Size(90, 50);
            this.move_sheet_button.TabIndex = 0;
            this.move_sheet_button.Text = "多工作簿表转同工作簿";
            this.move_sheet_button.UseVisualStyleBackColor = true;
            this.move_sheet_button.Click += new System.EventHandler(this.move_sheet_button_Click);
            // 
            // run_result_label
            // 
            this.run_result_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.run_result_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.run_result_label.Location = new System.Drawing.Point(41, 253);
            this.run_result_label.Name = "run_result_label";
            this.run_result_label.Size = new System.Drawing.Size(246, 51);
            this.run_result_label.TabIndex = 9;
            this.run_result_label.Text = "运行结果";
            // 
            // regex_clear_button
            // 
            this.regex_clear_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.regex_clear_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.regex_clear_button.Location = new System.Drawing.Point(227, 318);
            this.regex_clear_button.Name = "regex_clear_button";
            this.regex_clear_button.Size = new System.Drawing.Size(60, 30);
            this.regex_clear_button.TabIndex = 8;
            this.regex_clear_button.Text = "清空";
            this.regex_clear_button.UseVisualStyleBackColor = true;
            this.regex_clear_button.Click += new System.EventHandler(this.regex_clear_button_Click);
            // 
            // regex_run_button
            // 
            this.regex_run_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.regex_run_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.regex_run_button.Location = new System.Drawing.Point(60, 317);
            this.regex_run_button.Name = "regex_run_button";
            this.regex_run_button.Size = new System.Drawing.Size(60, 30);
            this.regex_run_button.TabIndex = 7;
            this.regex_run_button.Text = "运行";
            this.regex_run_button.UseVisualStyleBackColor = true;
            this.regex_run_button.Click += new System.EventHandler(this.regex_run_button_Click);
            // 
            // regex_rule_textbox
            // 
            this.regex_rule_textbox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.regex_rule_textbox.Location = new System.Drawing.Point(112, 200);
            this.regex_rule_textbox.Name = "regex_rule_textbox";
            this.regex_rule_textbox.Size = new System.Drawing.Size(175, 26);
            this.regex_rule_textbox.TabIndex = 6;
            // 
            // regex_rule_label
            // 
            this.regex_rule_label.AutoSize = true;
            this.regex_rule_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.regex_rule_label.Location = new System.Drawing.Point(41, 203);
            this.regex_rule_label.Name = "regex_rule_label";
            this.regex_rule_label.Size = new System.Drawing.Size(65, 20);
            this.regex_rule_label.TabIndex = 5;
            this.regex_rule_label.Text = "过滤规则";
            // 
            // what_type_combobox
            // 
            this.what_type_combobox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.what_type_combobox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.what_type_combobox.FormattingEnabled = true;
            this.what_type_combobox.Items.AddRange(new object[] {
            "数字",
            "英文",
            "中文",
            "网址",
            "身份证号",
            "电子邮箱",
            "电话号码",
            "IP地址",
            "自定义"});
            this.what_type_combobox.Location = new System.Drawing.Point(112, 147);
            this.what_type_combobox.Name = "what_type_combobox";
            this.what_type_combobox.Size = new System.Drawing.Size(175, 28);
            this.what_type_combobox.TabIndex = 4;
            this.what_type_combobox.SelectedIndexChanged += new System.EventHandler(this.what_type_combobox_SelectedIndexChanged);
            // 
            // what_type_label
            // 
            this.what_type_label.AutoSize = true;
            this.what_type_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.what_type_label.Location = new System.Drawing.Point(41, 150);
            this.what_type_label.Name = "what_type_label";
            this.what_type_label.Size = new System.Drawing.Size(65, 20);
            this.what_type_label.TabIndex = 3;
            this.what_type_label.Text = "提取内容";
            // 
            // which_field_combobox
            // 
            this.which_field_combobox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.which_field_combobox.FormattingEnabled = true;
            this.which_field_combobox.Location = new System.Drawing.Point(112, 82);
            this.which_field_combobox.Name = "which_field_combobox";
            this.which_field_combobox.Size = new System.Drawing.Size(175, 28);
            this.which_field_combobox.TabIndex = 2;
            this.which_field_combobox.VisibleChanged += new System.EventHandler(this.which_field_combobox_VisibleChanged);
            // 
            // which_field_label
            // 
            this.which_field_label.AutoSize = true;
            this.which_field_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.which_field_label.Location = new System.Drawing.Point(41, 85);
            this.which_field_label.Name = "which_field_label";
            this.which_field_label.Size = new System.Drawing.Size(65, 20);
            this.which_field_label.TabIndex = 1;
            this.which_field_label.Text = "提取哪列";
            // 
            // function_title_label
            // 
            this.function_title_label.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.function_title_label.AutoSize = true;
            this.function_title_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.function_title_label.Location = new System.Drawing.Point(41, 30);
            this.function_title_label.Name = "function_title_label";
            this.function_title_label.Size = new System.Drawing.Size(65, 19);
            this.function_title_label.TabIndex = 0;
            this.function_title_label.Text = "功能标题";
            this.function_title_label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // tabPage5
            // 
            this.tabPage5.Controls.Add(this.richTextBox1);
            this.tabPage5.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tabPage5.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.tabPage5.Location = new System.Drawing.Point(124, 4);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Size = new System.Drawing.Size(675, 392);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "五、使用帮助";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // richTextBox1
            // 
            this.richTextBox1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.richTextBox1.Location = new System.Drawing.Point(77, 48);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(519, 320);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = resources.GetString("richTextBox1.Text");
            // 
            // tabPage6
            // 
            this.tabPage6.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.tabPage6.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.tabPage6.Location = new System.Drawing.Point(124, 4);
            this.tabPage6.Name = "tabPage6";
            this.tabPage6.Size = new System.Drawing.Size(675, 392);
            this.tabPage6.TabIndex = 5;
            this.tabPage6.Text = "六、退出工具";
            this.tabPage6.UseVisualStyleBackColor = true;
            // 
            // split_sheet_timer
            // 
            this.split_sheet_timer.Tick += new System.EventHandler(this.split_sheet_timer_Tick);
            // 
            // merge_sheet_timer
            // 
            this.merge_sheet_timer.Tick += new System.EventHandler(this.merge_sheet_timer_Tick);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.Color.DarkGreen;
            this.pictureBox1.Location = new System.Drawing.Point(0, 400);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(799, 50);
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.DarkGreen;
            this.label4.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.label4.Font = new System.Drawing.Font("Franklin Gothic Medium", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.SystemColors.Control;
            this.label4.Location = new System.Drawing.Point(257, 417);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(285, 18);
            this.label4.TabIndex = 2;
            this.label4.Text = "Copyright © He Kun.  All rights reserved";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.BackColor = System.Drawing.Color.DarkGreen;
            this.label7.Font = new System.Drawing.Font("Franklin Gothic Medium", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.Transparent;
            this.label7.Location = new System.Drawing.Point(12, 419);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(90, 16);
            this.label7.TabIndex = 3;
            this.label7.Text = "Version  2.2.1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.tabControl1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "表操作工具箱";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer2)).EndInit();
            this.splitContainer2.ResumeLayout(false);
            this.tabPage5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage5;
        private System.Windows.Forms.TabPage tabPage6;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button split_button;
        private System.Windows.Forms.Label split_sheet_result_label;
        private System.Windows.Forms.ComboBox field_name_combobox;
        private System.Windows.Forms.ComboBox sheet_name_combobox;
        private System.Windows.Forms.Button splitsheet_delete_button;
        private System.Windows.Forms.Button splitsheet_export_button;
        private System.Windows.Forms.ProgressBar split_sheet_progressBar;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Button single_merge_button;
        private System.Windows.Forms.CheckBox multi_merge_sheet_checkBox;
        private System.Windows.Forms.Button dir_select_button;
        private System.Windows.Forms.Button multi_merge_button;
        private System.Windows.Forms.TextBox dir_select_textbox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button batch_delete_button;
        private System.Windows.Forms.Button batch_export_button;
        private System.Windows.Forms.CheckBox all_select_checkbox;
        private System.Windows.Forms.ListBox sheet_listbox;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button clear_button;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.Button contents_button;
        private System.Windows.Forms.Button payslip_button;
        private System.Windows.Forms.Button regex_button;
        private System.Windows.Forms.Button transposition_button;
        private System.Windows.Forms.Button add_sheet_button;
        private System.Windows.Forms.Button move_sheet_button;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Label function_title_label;
        private System.Windows.Forms.ComboBox what_type_combobox;
        private System.Windows.Forms.Label what_type_label;
        private System.Windows.Forms.ComboBox which_field_combobox;
        private System.Windows.Forms.Label which_field_label;
        private System.Windows.Forms.Button regex_clear_button;
        private System.Windows.Forms.Button regex_run_button;
        private System.Windows.Forms.TextBox regex_rule_textbox;
        private System.Windows.Forms.Label regex_rule_label;
        private System.Windows.Forms.Label run_result_label;
        private System.Windows.Forms.Timer split_sheet_timer;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ProgressBar merge_sheet_progressBar;
        private System.Windows.Forms.Label merge_sheet_result_label;
        private System.Windows.Forms.Timer merge_sheet_timer;
        private System.Windows.Forms.Label splitProgressBar_label;
        private System.Windows.Forms.Label mergeProgressBar_label;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label7;
    }
}