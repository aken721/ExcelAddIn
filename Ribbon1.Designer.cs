namespace ExcelAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.excel_extend = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.send_mail = this.Factory.CreateRibbonButton();
            this.send_message = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.files_read = this.Factory.CreateRibbonButton();
            this.file_rename = this.Factory.CreateRibbonButton();
            this.select_f_or_d = this.Factory.CreateRibbonToggleButton();
            this.switch_FD_label = this.Factory.CreateRibbonLabel();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.rename_mp3 = this.Factory.CreateRibbonButton();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.excel_extend);
            this.group1.Label = "表工具";
            this.group1.Name = "group1";
            // 
            // excel_extend
            // 
            this.excel_extend.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.excel_extend.Image = global::ExcelAddIn.Properties.Resources.excel;
            this.excel_extend.Label = "Excel表操作";
            this.excel_extend.Name = "excel_extend";
            this.excel_extend.ShowImage = true;
            this.excel_extend.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.excel_extend_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.send_mail);
            this.group2.Items.Add(this.send_message);
            this.group2.Label = "群发工具";
            this.group2.Name = "group2";
            // 
            // send_mail
            // 
            this.send_mail.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.send_mail.Image = global::ExcelAddIn.Properties.Resources.email;
            this.send_mail.Label = "Email群发";
            this.send_mail.Name = "send_mail";
            this.send_mail.ShowImage = true;
            this.send_mail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.send_mail_Click);
            // 
            // send_message
            // 
            this.send_message.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.send_message.Enabled = false;
            this.send_message.Image = global::ExcelAddIn.Properties.Resources.message;
            this.send_message.Label = "微信群发";
            this.send_message.Name = "send_message";
            this.send_message.ShowImage = true;
            this.send_message.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.send_message_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.files_read);
            this.group3.Items.Add(this.file_rename);
            this.group3.Items.Add(this.select_f_or_d);
            this.group3.Items.Add(this.switch_FD_label);
            this.group3.Label = "文件/文件夹工具";
            this.group3.Name = "group3";
            // 
            // files_read
            // 
            this.files_read.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.files_read.Image = global::ExcelAddIn.Properties.Resources.read;
            this.files_read.Label = "批读文件名";
            this.files_read.Name = "files_read";
            this.files_read.ShowImage = true;
            this.files_read.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.files_read_Click);
            // 
            // file_rename
            // 
            this.file_rename.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.file_rename.Image = global::ExcelAddIn.Properties.Resources.write;
            this.file_rename.Label = "批量重命名";
            this.file_rename.Name = "file_rename";
            this.file_rename.ShowImage = true;
            this.file_rename.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.file_rename_Click);
            // 
            // select_f_or_d
            // 
            this.select_f_or_d.Image = global::ExcelAddIn.Properties.Resources.Radio_Button_off;
            this.select_f_or_d.Label = "文件名";
            this.select_f_or_d.Name = "select_f_or_d";
            this.select_f_or_d.ShowImage = true;
            this.select_f_or_d.ShowLabel = false;
            this.select_f_or_d.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.select_f_or_d_Click);
            // 
            // switch_FD_label
            // 
            this.switch_FD_label.Label = "文件名";
            this.switch_FD_label.Name = "switch_FD_label";
            // 
            // group4
            // 
            this.group4.Items.Add(this.rename_mp3);
            this.group4.Label = "音乐工具";
            this.group4.Name = "group4";
            // 
            // rename_mp3
            // 
            this.rename_mp3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.rename_mp3.Image = global::ExcelAddIn.Properties.Resources.MP3;
            this.rename_mp3.Label = "MP3批量改名";
            this.rename_mp3.Name = "rename_mp3";
            this.rename_mp3.ShowImage = true;
            this.rename_mp3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rename_mp3_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton excel_extend;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton send_mail;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton send_message;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton files_read;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton file_rename;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton select_f_or_d;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rename_mp3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel switch_FD_label;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
