namespace ExcelAddIn
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.subject_label = new System.Windows.Forms.Label();
            this.body_label = new System.Windows.Forms.Label();
            this.attachment_label = new System.Windows.Forms.Label();
            this.mailto_label = new System.Windows.Forms.Label();
            this.mailfrom_label = new System.Windows.Forms.Label();
            this.smtp_label = new System.Windows.Forms.Label();
            this.subject_textBox = new System.Windows.Forms.TextBox();
            this.body_richTextBox = new System.Windows.Forms.RichTextBox();
            this.attachment_no_radioButton = new System.Windows.Forms.RadioButton();
            this.attachment_yes_radioButton = new System.Windows.Forms.RadioButton();
            this.attachment_checkBox = new System.Windows.Forms.CheckBox();
            this.mailto_textBox = new System.Windows.Forms.TextBox();
            this.mailto_comboBox = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.mailfrom_textBox = new System.Windows.Forms.TextBox();
            this.mailfrom_comboBox = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.mailpassword_label = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.mailpassword_textBox = new System.Windows.Forms.TextBox();
            this.smtp_textBox = new System.Windows.Forms.TextBox();
            this.port_label = new System.Windows.Forms.Label();
            this.port_textBox = new System.Windows.Forms.TextBox();
            this.send_button = new System.Windows.Forms.Button();
            this.clear_button = new System.Windows.Forms.Button();
            this.quit_button = new System.Windows.Forms.Button();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.mailpassword_pictureBox = new System.Windows.Forms.PictureBox();
            this.send_progressBar = new System.Windows.Forms.ProgressBar();
            this.send_progress_label = new System.Windows.Forms.Label();
            this.ssl_checkBox = new System.Windows.Forms.CheckBox();
            this.attachment_openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.attachment_folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.attachment_textBox = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.mailpassword_pictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // subject_label
            // 
            this.subject_label.AutoSize = true;
            this.subject_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.subject_label.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.subject_label.Location = new System.Drawing.Point(136, 41);
            this.subject_label.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.subject_label.Name = "subject_label";
            this.subject_label.Size = new System.Drawing.Size(82, 24);
            this.subject_label.TabIndex = 17;
            this.subject_label.Text = "邮件主题";
            // 
            // body_label
            // 
            this.body_label.AutoSize = true;
            this.body_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.body_label.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.body_label.Location = new System.Drawing.Point(136, 89);
            this.body_label.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.body_label.Name = "body_label";
            this.body_label.Size = new System.Drawing.Size(82, 24);
            this.body_label.TabIndex = 18;
            this.body_label.Text = "邮件内容";
            // 
            // attachment_label
            // 
            this.attachment_label.AutoSize = true;
            this.attachment_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.attachment_label.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.attachment_label.Location = new System.Drawing.Point(136, 256);
            this.attachment_label.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.attachment_label.Name = "attachment_label";
            this.attachment_label.Size = new System.Drawing.Size(87, 24);
            this.attachment_label.TabIndex = 19;
            this.attachment_label.Text = "附件匹配 ";
            // 
            // mailto_label
            // 
            this.mailto_label.AutoSize = true;
            this.mailto_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mailto_label.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.mailto_label.Location = new System.Drawing.Point(136, 310);
            this.mailto_label.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.mailto_label.Name = "mailto_label";
            this.mailto_label.Size = new System.Drawing.Size(82, 24);
            this.mailto_label.TabIndex = 21;
            this.mailto_label.Text = "接收邮箱";
            // 
            // mailfrom_label
            // 
            this.mailfrom_label.AutoSize = true;
            this.mailfrom_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mailfrom_label.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.mailfrom_label.Location = new System.Drawing.Point(136, 361);
            this.mailfrom_label.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.mailfrom_label.Name = "mailfrom_label";
            this.mailfrom_label.Size = new System.Drawing.Size(82, 24);
            this.mailfrom_label.TabIndex = 24;
            this.mailfrom_label.Text = "发送邮箱";
            // 
            // smtp_label
            // 
            this.smtp_label.AutoSize = true;
            this.smtp_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.smtp_label.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.smtp_label.Location = new System.Drawing.Point(136, 414);
            this.smtp_label.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.smtp_label.Name = "smtp_label";
            this.smtp_label.Size = new System.Drawing.Size(59, 24);
            this.smtp_label.TabIndex = 29;
            this.smtp_label.Text = "SMTP";
            this.smtp_label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // subject_textBox
            // 
            this.subject_textBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.subject_textBox.Location = new System.Drawing.Point(232, 38);
            this.subject_textBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.subject_textBox.Name = "subject_textBox";
            this.subject_textBox.Size = new System.Drawing.Size(613, 31);
            this.subject_textBox.TabIndex = 0;
            // 
            // body_richTextBox
            // 
            this.body_richTextBox.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.body_richTextBox.Location = new System.Drawing.Point(231, 92);
            this.body_richTextBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.body_richTextBox.Name = "body_richTextBox";
            this.body_richTextBox.Size = new System.Drawing.Size(615, 145);
            this.body_richTextBox.TabIndex = 1;
            this.body_richTextBox.Text = "";
            // 
            // attachment_no_radioButton
            // 
            this.attachment_no_radioButton.AutoSize = true;
            this.attachment_no_radioButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.attachment_no_radioButton.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.attachment_no_radioButton.Location = new System.Drawing.Point(232, 260);
            this.attachment_no_radioButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.attachment_no_radioButton.Name = "attachment_no_radioButton";
            this.attachment_no_radioButton.Size = new System.Drawing.Size(105, 24);
            this.attachment_no_radioButton.TabIndex = 2;
            this.attachment_no_radioButton.TabStop = true;
            this.attachment_no_radioButton.Text = "不发送附件";
            this.attachment_no_radioButton.UseVisualStyleBackColor = true;
            // 
            // attachment_yes_radioButton
            // 
            this.attachment_yes_radioButton.AutoSize = true;
            this.attachment_yes_radioButton.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.attachment_yes_radioButton.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.attachment_yes_radioButton.Location = new System.Drawing.Point(371, 260);
            this.attachment_yes_radioButton.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.attachment_yes_radioButton.Name = "attachment_yes_radioButton";
            this.attachment_yes_radioButton.Size = new System.Drawing.Size(90, 24);
            this.attachment_yes_radioButton.TabIndex = 3;
            this.attachment_yes_radioButton.TabStop = true;
            this.attachment_yes_radioButton.Text = "发送附件";
            this.attachment_yes_radioButton.UseVisualStyleBackColor = true;
            this.attachment_yes_radioButton.CheckedChanged += new System.EventHandler(this.attachment_yes_radioButton_CheckedChanged);
            // 
            // attachment_checkBox
            // 
            this.attachment_checkBox.AutoSize = true;
            this.attachment_checkBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.attachment_checkBox.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.attachment_checkBox.Location = new System.Drawing.Point(744, 259);
            this.attachment_checkBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.attachment_checkBox.Name = "attachment_checkBox";
            this.attachment_checkBox.Size = new System.Drawing.Size(121, 24);
            this.attachment_checkBox.TabIndex = 5;
            this.attachment_checkBox.Text = "发送不同附件";
            this.attachment_checkBox.UseVisualStyleBackColor = true;
            this.attachment_checkBox.CheckedChanged += new System.EventHandler(this.attachment_checkBox_CheckedChanged);
            this.attachment_checkBox.CheckStateChanged += new System.EventHandler(this.attachment_checkBox_CheckStateChanged);
            // 
            // mailto_textBox
            // 
            this.mailto_textBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mailto_textBox.Location = new System.Drawing.Point(231, 308);
            this.mailto_textBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.mailto_textBox.Name = "mailto_textBox";
            this.mailto_textBox.Size = new System.Drawing.Size(335, 27);
            this.mailto_textBox.TabIndex = 6;
            // 
            // mailto_comboBox
            // 
            this.mailto_comboBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mailto_comboBox.FormattingEnabled = true;
            this.mailto_comboBox.Location = new System.Drawing.Point(613, 308);
            this.mailto_comboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.mailto_comboBox.Name = "mailto_comboBox";
            this.mailto_comboBox.Size = new System.Drawing.Size(223, 28);
            this.mailto_comboBox.TabIndex = 7;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label7.Location = new System.Drawing.Point(575, 308);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(28, 24);
            this.label7.TabIndex = 22;
            this.label7.Text = "或";
            // 
            // mailfrom_textBox
            // 
            this.mailfrom_textBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mailfrom_textBox.Location = new System.Drawing.Point(231, 359);
            this.mailfrom_textBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.mailfrom_textBox.Name = "mailfrom_textBox";
            this.mailfrom_textBox.Size = new System.Drawing.Size(135, 27);
            this.mailfrom_textBox.TabIndex = 8;
            // 
            // mailfrom_comboBox
            // 
            this.mailfrom_comboBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mailfrom_comboBox.FormattingEnabled = true;
            this.mailfrom_comboBox.Items.AddRange(new object[] {
            "163.com",
            "qq.com",
            "sina.com",
            "126.com",
            "yeah.net",
            "sohu.com",
            "139.com",
            "189.cn",
            "wo.cn",
            "gmail.com",
            "outlook.com",
            "hotmail.com",
            "aliyun.com"});
            this.mailfrom_comboBox.Location = new System.Drawing.Point(389, 358);
            this.mailfrom_comboBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.mailfrom_comboBox.Name = "mailfrom_comboBox";
            this.mailfrom_comboBox.Size = new System.Drawing.Size(131, 28);
            this.mailfrom_comboBox.TabIndex = 9;
            this.mailfrom_comboBox.TextChanged += new System.EventHandler(this.mailfrom_comboBox_TextChanged);
            this.mailfrom_comboBox.GotFocus += new System.EventHandler(this.mailfrom_comboBox_GotFocus);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label8.Location = new System.Drawing.Point(360, 361);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(29, 24);
            this.label8.TabIndex = 25;
            this.label8.Text = "@";
            // 
            // mailpassword_label
            // 
            this.mailpassword_label.AutoSize = true;
            this.mailpassword_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mailpassword_label.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.mailpassword_label.Location = new System.Drawing.Point(545, 361);
            this.mailpassword_label.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.mailpassword_label.Name = "mailpassword_label";
            this.mailpassword_label.Size = new System.Drawing.Size(100, 24);
            this.mailpassword_label.TabIndex = 27;
            this.mailpassword_label.Text = "发件箱密码";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.ForeColor = System.Drawing.Color.Red;
            this.label10.Location = new System.Drawing.Point(123, 309);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(18, 24);
            this.label10.TabIndex = 20;
            this.label10.Text = "*";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label11.ForeColor = System.Drawing.Color.Red;
            this.label11.Location = new System.Drawing.Point(123, 356);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(18, 24);
            this.label11.TabIndex = 23;
            this.label11.Text = "*";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label12.ForeColor = System.Drawing.Color.Red;
            this.label12.Location = new System.Drawing.Point(529, 359);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(18, 24);
            this.label12.TabIndex = 26;
            this.label12.Text = "*";
            // 
            // mailpassword_textBox
            // 
            this.mailpassword_textBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.mailpassword_textBox.Location = new System.Drawing.Point(660, 358);
            this.mailpassword_textBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.mailpassword_textBox.Name = "mailpassword_textBox";
            this.mailpassword_textBox.Size = new System.Drawing.Size(139, 27);
            this.mailpassword_textBox.TabIndex = 10;
            this.mailpassword_textBox.UseSystemPasswordChar = true;
            // 
            // smtp_textBox
            // 
            this.smtp_textBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.smtp_textBox.Location = new System.Drawing.Point(231, 411);
            this.smtp_textBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.smtp_textBox.Name = "smtp_textBox";
            this.smtp_textBox.Size = new System.Drawing.Size(193, 27);
            this.smtp_textBox.TabIndex = 11;
            // 
            // port_label
            // 
            this.port_label.AutoSize = true;
            this.port_label.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.port_label.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.port_label.Location = new System.Drawing.Point(457, 412);
            this.port_label.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.port_label.Name = "port_label";
            this.port_label.Size = new System.Drawing.Size(58, 24);
            this.port_label.TabIndex = 31;
            this.port_label.Text = "PORT";
            this.port_label.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // port_textBox
            // 
            this.port_textBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.port_textBox.Location = new System.Drawing.Point(527, 411);
            this.port_textBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.port_textBox.Name = "port_textBox";
            this.port_textBox.Size = new System.Drawing.Size(65, 27);
            this.port_textBox.TabIndex = 12;
            // 
            // send_button
            // 
            this.send_button.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.send_button.ForeColor = System.Drawing.Color.DarkGreen;
            this.send_button.Location = new System.Drawing.Point(141, 502);
            this.send_button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.send_button.Name = "send_button";
            this.send_button.Size = new System.Drawing.Size(100, 39);
            this.send_button.TabIndex = 14;
            this.send_button.Text = "发送";
            this.send_button.UseVisualStyleBackColor = true;
            this.send_button.Click += new System.EventHandler(this.send_button_Click);
            // 
            // clear_button
            // 
            this.clear_button.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.clear_button.ForeColor = System.Drawing.Color.DarkGreen;
            this.clear_button.Location = new System.Drawing.Point(467, 502);
            this.clear_button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.clear_button.Name = "clear_button";
            this.clear_button.Size = new System.Drawing.Size(100, 39);
            this.clear_button.TabIndex = 15;
            this.clear_button.Text = "清空";
            this.clear_button.UseVisualStyleBackColor = true;
            this.clear_button.Click += new System.EventHandler(this.clear_button_Click);
            // 
            // quit_button
            // 
            this.quit_button.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.quit_button.ForeColor = System.Drawing.Color.DarkGreen;
            this.quit_button.Location = new System.Drawing.Point(733, 502);
            this.quit_button.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.quit_button.Name = "quit_button";
            this.quit_button.Size = new System.Drawing.Size(100, 39);
            this.quit_button.TabIndex = 16;
            this.quit_button.Text = "退出";
            this.quit_button.UseVisualStyleBackColor = true;
            this.quit_button.Click += new System.EventHandler(this.quit_button_Click);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label14.ForeColor = System.Drawing.Color.Red;
            this.label14.Location = new System.Drawing.Point(123, 411);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(18, 24);
            this.label14.TabIndex = 28;
            this.label14.Text = "*";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label15.ForeColor = System.Drawing.Color.Red;
            this.label15.Location = new System.Drawing.Point(440, 411);
            this.label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(18, 24);
            this.label15.TabIndex = 30;
            this.label15.Text = "*";
            // 
            // mailpassword_pictureBox
            // 
            this.mailpassword_pictureBox.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.mailpassword_pictureBox.Image = global::ExcelAddIn.Properties.Resources.eye_hide;
            this.mailpassword_pictureBox.Location = new System.Drawing.Point(803, 356);
            this.mailpassword_pictureBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.mailpassword_pictureBox.Name = "mailpassword_pictureBox";
            this.mailpassword_pictureBox.Size = new System.Drawing.Size(27, 27);
            this.mailpassword_pictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.mailpassword_pictureBox.TabIndex = 23;
            this.mailpassword_pictureBox.TabStop = false;
            this.mailpassword_pictureBox.Click += new System.EventHandler(this.mailpassword_pictureBox_Click);
            // 
            // send_progressBar
            // 
            this.send_progressBar.Location = new System.Drawing.Point(420, 464);
            this.send_progressBar.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.send_progressBar.Name = "send_progressBar";
            this.send_progressBar.Size = new System.Drawing.Size(429, 21);
            this.send_progressBar.TabIndex = 33;
            // 
            // send_progress_label
            // 
            this.send_progress_label.AutoSize = true;
            this.send_progress_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.send_progress_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.send_progress_label.Location = new System.Drawing.Point(124, 464);
            this.send_progress_label.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.send_progress_label.Name = "send_progress_label";
            this.send_progress_label.Size = new System.Drawing.Size(259, 20);
            this.send_progress_label.TabIndex = 32;
            this.send_progress_label.Text = "已发送？封，共？封，发送进度100%";
            // 
            // ssl_checkBox
            // 
            this.ssl_checkBox.AutoSize = true;
            this.ssl_checkBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ssl_checkBox.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.ssl_checkBox.Location = new System.Drawing.Point(607, 416);
            this.ssl_checkBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.ssl_checkBox.Name = "ssl_checkBox";
            this.ssl_checkBox.Size = new System.Drawing.Size(117, 24);
            this.ssl_checkBox.TabIndex = 13;
            this.ssl_checkBox.Text = "SSL加密发送";
            this.ssl_checkBox.UseVisualStyleBackColor = true;
            this.ssl_checkBox.CheckedChanged += new System.EventHandler(this.ssl_checkBox_CheckedChanged);
            // 
            // attachment_openFileDialog
            // 
            this.attachment_openFileDialog.FileName = "openFileDialog1";
            // 
            // attachment_textBox
            // 
            this.attachment_textBox.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.attachment_textBox.ForeColor = System.Drawing.Color.LightGray;
            this.attachment_textBox.Location = new System.Drawing.Point(467, 256);
            this.attachment_textBox.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.attachment_textBox.Name = "attachment_textBox";
            this.attachment_textBox.Size = new System.Drawing.Size(268, 27);
            this.attachment_textBox.TabIndex = 4;
            this.attachment_textBox.Text = "请手工输入文件的完整路径，或双击选择文件";
            this.attachment_textBox.Click += new System.EventHandler(this.attachment_textBox_Click);
            this.attachment_textBox.DoubleClick += new System.EventHandler(this.attachment_textBox_DoubleClick);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(965, 571);
            this.Controls.Add(this.attachment_textBox);
            this.Controls.Add(this.ssl_checkBox);
            this.Controls.Add(this.send_progress_label);
            this.Controls.Add(this.send_progressBar);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.quit_button);
            this.Controls.Add(this.clear_button);
            this.Controls.Add(this.send_button);
            this.Controls.Add(this.port_textBox);
            this.Controls.Add(this.port_label);
            this.Controls.Add(this.smtp_textBox);
            this.Controls.Add(this.mailpassword_pictureBox);
            this.Controls.Add(this.mailpassword_textBox);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.mailpassword_label);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.mailfrom_comboBox);
            this.Controls.Add(this.mailfrom_textBox);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.mailto_comboBox);
            this.Controls.Add(this.mailto_textBox);
            this.Controls.Add(this.attachment_checkBox);
            this.Controls.Add(this.attachment_yes_radioButton);
            this.Controls.Add(this.attachment_no_radioButton);
            this.Controls.Add(this.body_richTextBox);
            this.Controls.Add(this.subject_textBox);
            this.Controls.Add(this.smtp_label);
            this.Controls.Add(this.mailfrom_label);
            this.Controls.Add(this.mailto_label);
            this.Controls.Add(this.attachment_label);
            this.Controls.Add(this.body_label);
            this.Controls.Add(this.subject_label);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Form2";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "邮件群发";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.VisibleChanged += new System.EventHandler(this.Form2_VisibleChanged);
            ((System.ComponentModel.ISupportInitialize)(this.mailpassword_pictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label subject_label;
        private System.Windows.Forms.Label body_label;
        private System.Windows.Forms.Label attachment_label;
        private System.Windows.Forms.Label mailto_label;
        private System.Windows.Forms.Label mailfrom_label;
        private System.Windows.Forms.Label smtp_label;
        private System.Windows.Forms.TextBox subject_textBox;
        private System.Windows.Forms.RichTextBox body_richTextBox;
        private System.Windows.Forms.RadioButton attachment_no_radioButton;
        private System.Windows.Forms.RadioButton attachment_yes_radioButton;
        private System.Windows.Forms.CheckBox attachment_checkBox;
        private System.Windows.Forms.TextBox mailto_textBox;
        private System.Windows.Forms.ComboBox mailto_comboBox;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox mailfrom_textBox;
        private System.Windows.Forms.ComboBox mailfrom_comboBox;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label mailpassword_label;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox mailpassword_textBox;
        private System.Windows.Forms.PictureBox mailpassword_pictureBox;
        private System.Windows.Forms.TextBox smtp_textBox;
        private System.Windows.Forms.Label port_label;
        private System.Windows.Forms.TextBox port_textBox;
        private System.Windows.Forms.Button send_button;
        private System.Windows.Forms.Button clear_button;
        private System.Windows.Forms.Button quit_button;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.ProgressBar send_progressBar;
        private System.Windows.Forms.Label send_progress_label;
        private System.Windows.Forms.CheckBox ssl_checkBox;
        private System.Windows.Forms.OpenFileDialog attachment_openFileDialog;
        private System.Windows.Forms.FolderBrowserDialog attachment_folderBrowserDialog;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.TextBox attachment_textBox;
    }
}