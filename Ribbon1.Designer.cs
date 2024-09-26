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
            this.components = new System.ComponentModel.Container();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl6 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl7 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl8 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl9 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl10 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl11 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl12 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl13 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl14 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.separator4 = this.Factory.CreateRibbonSeparator();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.separator3 = this.Factory.CreateRibbonSeparator();
            this.switch_FD_label = this.Factory.CreateRibbonLabel();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.box2 = this.Factory.CreateRibbonBox();
            this.sheet_export_comboBox = this.Factory.CreateRibbonComboBox();
            this.export_type_comboBox = this.Factory.CreateRibbonComboBox();
            this.box3 = this.Factory.CreateRibbonBox();
            this.page_orientation_comboBox = this.Factory.CreateRibbonComboBox();
            this.paper_size_comboBox = this.Factory.CreateRibbonComboBox();
            this.page_zoom_comboBox = this.Factory.CreateRibbonComboBox();
            this.group6 = this.Factory.CreateRibbonGroup();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.box1 = this.Factory.CreateRibbonBox();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.buttonGroup2 = this.Factory.CreateRibbonButtonGroup();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.Excel_extend = this.Factory.CreateRibbonButton();
            this.confirm_spotlight = this.Factory.CreateRibbonToggleButton();
            this.Send_mail = this.Factory.CreateRibbonButton();
            this.Files_read = this.Factory.CreateRibbonButton();
            this.File_rename = this.Factory.CreateRibbonButton();
            this.delandmove_button = this.Factory.CreateRibbonButton();
            this.Select_f_or_d = this.Factory.CreateRibbonToggleButton();
            this.to_pdf_button = this.Factory.CreateRibbonButton();
            this.scan_button = this.Factory.CreateRibbonButton();
            this.Rename_mp3 = this.Factory.CreateRibbonButton();
            this.Select_mp3_button = this.Factory.CreateRibbonButton();
            this.Mode_button = this.Factory.CreateRibbonButton();
            this.Play_button = this.Factory.CreateRibbonButton();
            this.Stop_button = this.Factory.CreateRibbonButton();
            this.Next_button = this.Factory.CreateRibbonButton();
            this.Previous_button = this.Factory.CreateRibbonButton();
            this.separator5 = this.Factory.CreateRibbonSeparator();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.box2.SuspendLayout();
            this.box3.SuspendLayout();
            this.group6.SuspendLayout();
            this.box1.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.buttonGroup2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group6);
            this.tab1.Label = "工具箱";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.Excel_extend);
            this.group1.Items.Add(this.separator4);
            this.group1.Items.Add(this.confirm_spotlight);
            this.group1.Items.Add(this.separator5);
            this.group1.Items.Add(this.scan_button);
            this.group1.Label = "表工具";
            this.group1.Name = "group1";
            // 
            // separator4
            // 
            this.separator4.Name = "separator4";
            // 
            // group2
            // 
            this.group2.Items.Add(this.Send_mail);
            this.group2.Label = "群发工具";
            this.group2.Name = "group2";
            // 
            // group3
            // 
            this.group3.Items.Add(this.Files_read);
            this.group3.Items.Add(this.separator1);
            this.group3.Items.Add(this.File_rename);
            this.group3.Items.Add(this.separator3);
            this.group3.Items.Add(this.delandmove_button);
            this.group3.Items.Add(this.Select_f_or_d);
            this.group3.Items.Add(this.switch_FD_label);
            this.group3.Label = "文件/文件夹工具";
            this.group3.Name = "group3";
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // separator3
            // 
            this.separator3.Name = "separator3";
            // 
            // switch_FD_label
            // 
            this.switch_FD_label.Label = "文件名";
            this.switch_FD_label.Name = "switch_FD_label";
            // 
            // group4
            // 
            this.group4.Items.Add(this.to_pdf_button);
            this.group4.Items.Add(this.box2);
            this.group4.Items.Add(this.box3);
            this.group4.Label = "PDF工具";
            this.group4.Name = "group4";
            // 
            // box2
            // 
            this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box2.Items.Add(this.sheet_export_comboBox);
            this.box2.Items.Add(this.export_type_comboBox);
            this.box2.Name = "box2";
            // 
            // sheet_export_comboBox
            // 
            ribbonDropDownItemImpl1.Label = "当前表";
            ribbonDropDownItemImpl2.Label = "全部表";
            this.sheet_export_comboBox.Items.Add(ribbonDropDownItemImpl1);
            this.sheet_export_comboBox.Items.Add(ribbonDropDownItemImpl2);
            this.sheet_export_comboBox.Label = "导出表集";
            this.sheet_export_comboBox.Name = "sheet_export_comboBox";
            this.sheet_export_comboBox.Text = null;
            this.sheet_export_comboBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.sheet_export_comboBox_TextChanged);
            // 
            // export_type_comboBox
            // 
            ribbonDropDownItemImpl3.Label = "多表单文件";
            ribbonDropDownItemImpl4.Label = "多表多文件";
            this.export_type_comboBox.Items.Add(ribbonDropDownItemImpl3);
            this.export_type_comboBox.Items.Add(ribbonDropDownItemImpl4);
            this.export_type_comboBox.Label = "导出方式";
            this.export_type_comboBox.Name = "export_type_comboBox";
            this.export_type_comboBox.Text = null;
            this.export_type_comboBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.export_type_comboBox_TextChanged);
            // 
            // box3
            // 
            this.box3.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box3.Items.Add(this.page_orientation_comboBox);
            this.box3.Items.Add(this.paper_size_comboBox);
            this.box3.Items.Add(this.page_zoom_comboBox);
            this.box3.Name = "box3";
            // 
            // page_orientation_comboBox
            // 
            ribbonDropDownItemImpl5.Label = "纵向";
            ribbonDropDownItemImpl6.Label = "横向";
            this.page_orientation_comboBox.Items.Add(ribbonDropDownItemImpl5);
            this.page_orientation_comboBox.Items.Add(ribbonDropDownItemImpl6);
            this.page_orientation_comboBox.Label = "页面方向";
            this.page_orientation_comboBox.Name = "page_orientation_comboBox";
            this.page_orientation_comboBox.Text = null;
            this.page_orientation_comboBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.page_orientation_comboBox_TextChanged);
            // 
            // paper_size_comboBox
            // 
            ribbonDropDownItemImpl7.Label = "A4";
            ribbonDropDownItemImpl8.Label = "A3";
            ribbonDropDownItemImpl9.Label = "A5";
            ribbonDropDownItemImpl10.Label = "B5";
            this.paper_size_comboBox.Items.Add(ribbonDropDownItemImpl7);
            this.paper_size_comboBox.Items.Add(ribbonDropDownItemImpl8);
            this.paper_size_comboBox.Items.Add(ribbonDropDownItemImpl9);
            this.paper_size_comboBox.Items.Add(ribbonDropDownItemImpl10);
            this.paper_size_comboBox.Label = "页面大小";
            this.paper_size_comboBox.Name = "paper_size_comboBox";
            this.paper_size_comboBox.Text = null;
            this.paper_size_comboBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.paper_size_comboBox_TextChanged);
            // 
            // page_zoom_comboBox
            // 
            ribbonDropDownItemImpl11.Label = "无缩放";
            ribbonDropDownItemImpl12.Label = "表自适应";
            ribbonDropDownItemImpl13.Label = "行自适应";
            ribbonDropDownItemImpl14.Label = "列自适应";
            this.page_zoom_comboBox.Items.Add(ribbonDropDownItemImpl11);
            this.page_zoom_comboBox.Items.Add(ribbonDropDownItemImpl12);
            this.page_zoom_comboBox.Items.Add(ribbonDropDownItemImpl13);
            this.page_zoom_comboBox.Items.Add(ribbonDropDownItemImpl14);
            this.page_zoom_comboBox.Label = "页面缩放";
            this.page_zoom_comboBox.Name = "page_zoom_comboBox";
            this.page_zoom_comboBox.Text = null;
            this.page_zoom_comboBox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.page_zoom_comboBox_TextChanged);
            // 
            // group6
            // 
            this.group6.Items.Add(this.Rename_mp3);
            this.group6.Items.Add(this.separator2);
            this.group6.Items.Add(this.box1);
            this.group6.Label = "音乐工具";
            this.group6.Name = "group6";
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.buttonGroup1);
            this.box1.Items.Add(this.buttonGroup2);
            this.box1.Name = "box1";
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.Select_mp3_button);
            this.buttonGroup1.Items.Add(this.Mode_button);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // buttonGroup2
            // 
            this.buttonGroup2.Items.Add(this.Play_button);
            this.buttonGroup2.Items.Add(this.Stop_button);
            this.buttonGroup2.Items.Add(this.Next_button);
            this.buttonGroup2.Items.Add(this.Previous_button);
            this.buttonGroup2.Name = "buttonGroup2";
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // Excel_extend
            // 
            this.Excel_extend.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Excel_extend.Image = global::ExcelAddIn.Properties.Resources.excel;
            this.Excel_extend.Label = "Excel表操作";
            this.Excel_extend.Name = "Excel_extend";
            this.Excel_extend.ShowImage = true;
            this.Excel_extend.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Excel_extend_Click);
            // 
            // confirm_spotlight
            // 
            this.confirm_spotlight.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.confirm_spotlight.Image = global::ExcelAddIn.Properties.Resources.spotlight_close;
            this.confirm_spotlight.Label = "打开聚光灯";
            this.confirm_spotlight.Name = "confirm_spotlight";
            this.confirm_spotlight.ShowImage = true;
            this.confirm_spotlight.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.confirm_spotlight_Click);
            // 
            // Send_mail
            // 
            this.Send_mail.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Send_mail.Image = global::ExcelAddIn.Properties.Resources.email;
            this.Send_mail.Label = "Email群发";
            this.Send_mail.Name = "Send_mail";
            this.Send_mail.ShowImage = true;
            this.Send_mail.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Send_mail_Click);
            // 
            // Files_read
            // 
            this.Files_read.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Files_read.Image = global::ExcelAddIn.Properties.Resources.read;
            this.Files_read.Label = "批读文件名";
            this.Files_read.Name = "Files_read";
            this.Files_read.ShowImage = true;
            this.Files_read.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Files_read_Click);
            // 
            // File_rename
            // 
            this.File_rename.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.File_rename.Image = global::ExcelAddIn.Properties.Resources.write;
            this.File_rename.Label = "批量重命名";
            this.File_rename.Name = "File_rename";
            this.File_rename.ShowImage = true;
            this.File_rename.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.File_rename_Click);
            // 
            // delandmove_button
            // 
            this.delandmove_button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.delandmove_button.Image = global::ExcelAddIn.Properties.Resources.移动文件;
            this.delandmove_button.Label = "批量删或移";
            this.delandmove_button.Name = "delandmove_button";
            this.delandmove_button.ShowImage = true;
            this.delandmove_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.delandmove_button_Click);
            // 
            // Select_f_or_d
            // 
            this.Select_f_or_d.Image = global::ExcelAddIn.Properties.Resources.Radio_Button_off;
            this.Select_f_or_d.Label = "文件名";
            this.Select_f_or_d.Name = "Select_f_or_d";
            this.Select_f_or_d.ShowImage = true;
            this.Select_f_or_d.ShowLabel = false;
            this.Select_f_or_d.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Select_f_or_d_Click);
            // 
            // to_pdf_button
            // 
            this.to_pdf_button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.to_pdf_button.Image = global::ExcelAddIn.Properties.Resources.pdf;
            this.to_pdf_button.Label = "转为PDF";
            this.to_pdf_button.Name = "to_pdf_button";
            this.to_pdf_button.ShowImage = true;
            this.to_pdf_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.to_pdf_button_Click);
            // 
            // scan_button
            // 
            this.scan_button.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.scan_button.Image = global::ExcelAddIn.Properties.Resources.QR;
            this.scan_button.Label = "识别二维码";
            this.scan_button.Name = "scan_button";
            this.scan_button.ShowImage = true;
            this.scan_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.scan_button_Click);
            // 
            // Rename_mp3
            // 
            this.Rename_mp3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.Rename_mp3.Image = global::ExcelAddIn.Properties.Resources.MP3;
            this.Rename_mp3.Label = "MP3批量改名";
            this.Rename_mp3.Name = "Rename_mp3";
            this.Rename_mp3.ShowImage = true;
            this.Rename_mp3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Rename_mp3_Click);
            // 
            // Select_mp3_button
            // 
            this.Select_mp3_button.Image = global::ExcelAddIn.Properties.Resources.no_open_fold;
            this.Select_mp3_button.Label = "音乐目录";
            this.Select_mp3_button.Name = "Select_mp3_button";
            this.Select_mp3_button.ScreenTip = "选择音乐所在文件夹";
            this.Select_mp3_button.ShowImage = true;
            this.Select_mp3_button.ShowLabel = false;
            this.Select_mp3_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Select_mp3_button_Click);
            // 
            // Mode_button
            // 
            this.Mode_button.Image = global::ExcelAddIn.Properties.Resources.order_play;
            this.Mode_button.Label = "顺序播放";
            this.Mode_button.Name = "Mode_button";
            this.Mode_button.ScreenTip = "顺序播放";
            this.Mode_button.ShowImage = true;
            this.Mode_button.ShowLabel = false;
            this.Mode_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Mode_button_Click);
            // 
            // Play_button
            // 
            this.Play_button.Image = global::ExcelAddIn.Properties.Resources.play;
            this.Play_button.Label = "播放";
            this.Play_button.Name = "Play_button";
            this.Play_button.ScreenTip = "播放";
            this.Play_button.ShowImage = true;
            this.Play_button.ShowLabel = false;
            this.Play_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Play_button_Click);
            // 
            // Stop_button
            // 
            this.Stop_button.Image = global::ExcelAddIn.Properties.Resources.stop;
            this.Stop_button.Label = "停止";
            this.Stop_button.Name = "Stop_button";
            this.Stop_button.ScreenTip = "停止";
            this.Stop_button.ShowImage = true;
            this.Stop_button.ShowLabel = false;
            this.Stop_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Stop_button_Click);
            // 
            // Next_button
            // 
            this.Next_button.Image = global::ExcelAddIn.Properties.Resources.next;
            this.Next_button.Label = "下一首";
            this.Next_button.Name = "Next_button";
            this.Next_button.ScreenTip = "下一曲";
            this.Next_button.ShowImage = true;
            this.Next_button.ShowLabel = false;
            this.Next_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Next_button_Click);
            // 
            // Previous_button
            // 
            this.Previous_button.Image = global::ExcelAddIn.Properties.Resources.previous;
            this.Previous_button.Label = "上一首";
            this.Previous_button.Name = "Previous_button";
            this.Previous_button.ScreenTip = "上一曲";
            this.Previous_button.ShowImage = true;
            this.Previous_button.ShowLabel = false;
            this.Previous_button.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Previous_button_Click);
            // 
            // separator5
            // 
            this.separator5.Name = "separator5";
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
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();
            this.box3.ResumeLayout(false);
            this.box3.PerformLayout();
            this.group6.ResumeLayout(false);
            this.group6.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.buttonGroup2.ResumeLayout(false);
            this.buttonGroup2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Excel_extend;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Send_mail;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Files_read;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton File_rename;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton Select_f_or_d;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Rename_mp3;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel switch_FD_label;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Play_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Select_mp3_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Mode_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Stop_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Next_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Previous_button;
        private Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        private System.Windows.Forms.Timer timer1;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton confirm_spotlight;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton delandmove_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton to_pdf_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox page_orientation_comboBox;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox paper_size_comboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox page_zoom_comboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox sheet_export_comboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box3;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox export_type_comboBox;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton scan_button;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator5;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
