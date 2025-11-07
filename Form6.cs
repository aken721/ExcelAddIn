#nullable enable

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Docnet.Core;
using Docnet.Core.Converters;
using Docnet.Core.Models;
using UglyToad.PdfPig;
using PaddleOCRInvoice;
using PaddleOCRInvoice.Models;
using Excel = Microsoft.Office.Interop.Excel;




namespace ExcelAddIn
{
    public partial class Form6 : Form
    {
        public Form6()
        {
            InitializeComponent();
        }

        private void Form6_Load(object sender, EventArgs e)
        {
            Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            string version = assembly.GetName().Version.ToString();
            version_label1.Text = "Version: " + version;

            //初始化tabcontrol控件
            tabControl1.SelectTab(0);
            sequence_label.Text = string.Empty;
            single_FileFullNames.Clear();
            folder_FileFullNames.Clear();
            clear_pictureBox.Visible = false;
            timer1.Interval = 3000;
            tabControl1.DrawItem += new DrawItemEventHandler(tabControl1_DrawItem);
            
            // 初始化PaddleOCR发票识别服务
            InitializeInvoiceService();
        }
        
        // 初始化发票识别服务
        private void InitializeInvoiceService()
        {
            try
            {
                string basePath = AppDomain.CurrentDomain.BaseDirectory;
                string modelsPath = Path.Combine(basePath, "PaddleOCRInvoice", "Models");
                
                System.Diagnostics.Debug.WriteLine($"【调试】初始化路径: {modelsPath}");
                System.Diagnostics.Debug.WriteLine($"【调试】Models路径是否存在: {Directory.Exists(modelsPath)}");
                
                // 初始化发票识别服务（传递Models路径）
                _invoiceService = new InvoiceRecognitionService(modelsPath);
                
                System.Diagnostics.Debug.WriteLine("【调试】PaddleOCR发票识别服务初始化成功");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"【调试】初始化异常: {ex.ToString()}");
                MessageBox.Show($"发票识别服务初始化失败：\n\n错误信息: {ex.Message}\n\n内部异常: {ex.InnerException?.Message}\n\n堆栈跟踪:\n{ex.StackTrace}\n\nPDF/图片发票识别功能将不可用。", 
                    "初始化失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _invoiceService = null;
            }
        }



        //重绘选项页布局
        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            //调整选项卡文字方向
            SolidBrush _Brush = new SolidBrush(System.Drawing.Color.Black);//单色画刷
            RectangleF _TabTextArea = (RectangleF)tabControl1.GetTabRect(e.Index);//绘制区域
            StringFormat _sf = new StringFormat();//封装文本布局格式信息
            _sf.LineAlignment = StringAlignment.Center;
            _sf.Alignment = StringAlignment.Center;
            // 使用正确的方式获取TabPage的Text属性
            e.Graphics.DrawString(tabControl1.Controls[e.Index].Text, SystemInformation.MenuFont, _Brush, _TabTextArea, _sf);
        }

        //初始化选项卡选项页
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    sequence_label.Text = string.Empty;
                    xml_path_textBox.Text = string.Empty;
                    single_result_label.Text = "";
                    single_FileFullNames.Clear();
                    folder_FileFullNames.Clear();
                    clear_pictureBox.Visible = false;
                    break;
                case 1:
                    sequence_label1.Text = string.Empty;
                    pdf_path_textBox.Text = string.Empty;
                    single_result_label1.Text = "";
                    pdfSingle_FileFullNames.Clear();
                    pdfFolder_FileFullNames.Clear();
                    clear_pictureBox1.Visible = false;
                    break;
                case 2:
                    break;
                case 3:
                    this.Dispose();
                    break;
            }
        }


        /// <summary>
        /// 以下为xml数电发票读取功能模块
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>


        //下一个
        private void next_pictureBox_Click(object sender, EventArgs e)
        {
            if (single_FileFullNames.Count == 0) return;
            int currentSequence = int.Parse(sequence_label.Text);
            xml_path_textBox.Text = single_FileFullNames[currentSequence];
            sequence_label.Text = Convert.ToString(currentSequence + 1);
        }

        //上一个
        private void preview_pictureBox_Click(object sender, EventArgs e)
        {
            if (single_FileFullNames.Count == 0) return;
            int currentSequence = int.Parse(sequence_label.Text);
            xml_path_textBox.Text = single_FileFullNames[currentSequence - 2];
            sequence_label.Text = Convert.ToString(currentSequence - 1);
        }

        //最后一个
        private void last_pictureBox_Click(object sender, EventArgs e)
        {
            if (single_FileFullNames.Count == 0) return;
            xml_path_textBox.Text = single_FileFullNames[single_FileFullNames.Count - 1];
            sequence_label.Text = single_FileFullNames.Count.ToString();
        }

        //第一个
        private void begin_pictureBox_Click(object sender, EventArgs e)
        {
            if (single_FileFullNames.Count == 0) return;
            xml_path_textBox.Text = single_FileFullNames[0];
            sequence_label.Text = "1";
        }

        //顺序号变化时
        private void sequence_label_TextChanged(object sender, EventArgs e)
        {
            int fileCount = single_FileFullNames.Count;
            if (sequence_label.Text == "1" && fileCount > 1)
            {
                begin_pictureBox.Enabled = false;
                begin_pictureBox.Visible = false;
                preview_pictureBox.Enabled = false;
                preview_pictureBox.Visible = false;
                next_pictureBox.Enabled = true;
                next_pictureBox.Visible = true;
                last_pictureBox.Enabled = true;
                last_pictureBox.Visible = true;
            }
            else if (sequence_label.Text == fileCount.ToString() && fileCount > 1)
            {
                begin_pictureBox.Enabled = true;
                begin_pictureBox.Visible = true;
                preview_pictureBox.Enabled = true;
                preview_pictureBox.Visible = true;
                next_pictureBox.Enabled = false;
                next_pictureBox.Visible = false;
                last_pictureBox.Enabled = false;
                last_pictureBox.Visible = false;
            }
            else if (string.IsNullOrEmpty(sequence_label.Text))
            {
                begin_pictureBox.Enabled = false;
                begin_pictureBox.Visible = false;
                preview_pictureBox.Enabled = false;
                preview_pictureBox.Visible = false;
                next_pictureBox.Enabled = false;
                next_pictureBox.Visible = false;
                last_pictureBox.Enabled = false;
                last_pictureBox.Visible = false;
            }
            else if (sequence_label.Text == fileCount.ToString() && fileCount == 1)
            {
                begin_pictureBox.Enabled = false;
                begin_pictureBox.Visible = false;
                preview_pictureBox.Enabled = false;
                preview_pictureBox.Visible = false;
                next_pictureBox.Enabled = false;
                next_pictureBox.Visible = false;
                last_pictureBox.Enabled = false;
                last_pictureBox.Visible = false;
            }
            else
            {
                begin_pictureBox.Enabled = true;
                begin_pictureBox.Visible = true;
                preview_pictureBox.Enabled = true;
                preview_pictureBox.Visible = true;
                next_pictureBox.Enabled = true;
                next_pictureBox.Visible = true;
                last_pictureBox.Enabled = true;
                last_pictureBox.Visible = true;
            }
        }


        //递归获取xml每个节点元素，并为每个元素创建对应的 TreeNode 对象
        private void AddNodes(TreeNodeCollection nodes, XElement element)
        {
            foreach (XElement child in element.Elements())
            {
                TreeNode treeNode = new TreeNode(GetNodeText(child));
                treeNode.Tag = child; // 存储XElement对象以便后续使用
                nodes.Add(treeNode);
                if (child.HasElements)
                {
                    AddNodes(treeNode.Nodes, child);
                }
            }
        }

        private string GetNodeText(XElement element)
        {
            string text = element.Name.LocalName;
            if (!string.IsNullOrEmpty(element.Value))
            {
                text += ": " + element.Value;
            }
            if (element.HasAttributes)
            {
                foreach (XAttribute attribute in element.Attributes())
                {
                    text += Environment.NewLine + "  " + attribute.Name.LocalName + " = " + attribute.Value;
                }
            }
            return text;
        }

        //文本框内容改变时
        private void xml_path_textBox_TextChanged(object sender, EventArgs e)
        {
            string path = xml_path_textBox.Text;
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                XDocument xmlDoc = XDocument.Load(path); // 加载你的XML文件

                // 假设XML文件的根元素是你想要的根节点
                XElement root = xmlDoc.Root;
                AddNodes(xml_treeView.Nodes, root);

                xml_treeView.ExpandAll(); // 展开所有节点

                // 将TreeView的滚动条设置到最顶端
                xml_treeView.SelectedNode = xml_treeView.Nodes[0]; // 选择第一个节点
                xml_treeView.TopNode = xml_treeView.SelectedNode; // 将TopNode设置为选中的节点
                xml_treeView.SelectedNode.EnsureVisible(); // 确保选中的节点是可见的
            }
            else
            {
                xml_treeView.Nodes.Clear();
                if (single_FileFullNames.Count == 0)
                {
                    sequence_label.Text = string.Empty;
                    begin_pictureBox.Enabled = false;
                    begin_pictureBox.Visible = false;
                    preview_pictureBox.Enabled = false;
                    preview_pictureBox.Visible = false;
                    next_pictureBox.Enabled = false;
                    next_pictureBox.Visible = false;
                    last_pictureBox.Enabled = false;
                    last_pictureBox.Visible = false;
                }
            }
        }

        List<string> single_FileFullNames = new List<string>();
        List<string> folder_FileFullNames = new List<string>();
        List<string> pdfSingle_FileFullNames = new List<string>();
        List<string> pdfFolder_FileFullNames = new List<string>();
        
        // PaddleOCR发票识别服务
        private InvoiceRecognitionService? _invoiceService;

        private void xml_path_textBox_DoubleClick(object sender, EventArgs e)
        {
            openFileDialog1.Title = "请选择要导入的电子发票";
            openFileDialog1.Filter = "电子发票文件(*.xml)|*.xml";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sequence_label.Text = string.Empty;
                single_FileFullNames.Clear();
                single_FileFullNames = openFileDialog1.FileNames.ToList();
                xml_path_textBox.Text = single_FileFullNames[0];
                sequence_label.Text = "1";
                single_result_label.Text = $"共{single_FileFullNames.Count}个XML文件";
                if (single_FileFullNames.Count > 1)
                {
                    int fileCount = single_FileFullNames.Count;
                    if (sequence_label.Text == "1")
                    {
                        begin_pictureBox.Enabled = false;
                        preview_pictureBox.Enabled = false;
                        next_pictureBox.Enabled = true;
                        last_pictureBox.Enabled = true;
                    }
                    else if (sequence_label.Text == fileCount.ToString())
                    {
                        begin_pictureBox.Enabled = true;
                        preview_pictureBox.Enabled = true;
                        next_pictureBox.Enabled = false;
                        last_pictureBox.Enabled = false;
                    }
                    else
                    {
                        begin_pictureBox.Enabled = true;
                        preview_pictureBox.Enabled = true;
                        next_pictureBox.Enabled = true;
                        last_pictureBox.Enabled = true;
                    }
                }
                else
                {
                    begin_pictureBox.Enabled = false;
                    preview_pictureBox.Enabled = false;
                    next_pictureBox.Enabled = false;
                    last_pictureBox.Enabled = false;
                }
            }
            else
            {
                return;
            }
        }

        private void run_button_Click(object sender, EventArgs e)
        {
            string xml_path = xml_path_textBox.Text;
            if (string.IsNullOrEmpty(xml_path))
            {
                single_result_label.Text = "文件路径文本框不能为空，请确认后再试";
                timer1.Enabled = true;
                return;
            }
            System.Data.DataTable dataTable = GetInvoiceDataFromXML(xml_path);
            if (dataTable.Rows.Count > 0)
            {
                WriteToExcel(dataTable);
            }
            single_result_label.Text = "XML电子发票导入已完成！";
            timer1.Enabled = true;
        }

        private void folder_path_textBox_DoubleClick(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "请选择XML电子发票所在文件夹";
            folderBrowserDialog1.ShowDialog();
            if (folderBrowserDialog1.SelectedPath != "")
            {
                folder_FileFullNames.Clear();
                folder_path_textBox.Text = folderBrowserDialog1.SelectedPath;
                DirectoryInfo folder = new DirectoryInfo(folder_path_textBox.Text);
                folder_FileFullNames = folder.GetFiles("*.xml", subfolder_checkBox.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly).Select(file => file.FullName).ToList();
            }
            else
            {
                folder_path_textBox.Text = "";
                batch_result_label.Text = "未选择文件夹";
            }
        }

        private void batch_run_button_Click(object sender, EventArgs e)
        {

            if (folder_FileFullNames.Count > 0 && Directory.Exists(folder_path_textBox.Text))
            {
                int o = 1;
                foreach (string file in folder_FileFullNames)
                {
                    batch_result_label.Text = $"共{folder_FileFullNames.Count}个XML文件，正在导入第{o}个......";
                    System.Data.DataTable dataTable = GetInvoiceDataFromXML(file);
                    if (dataTable.Rows.Count > 0)
                    {
                        WriteToExcel(dataTable);
                    }
                }
                batch_result_label.Text = "XML电子发票批量导入已完成！";
                timer1.Enabled = true;
            }
            else
            {
                batch_result_label.Text = "文件夹内没有XML电子发票文件，请核对！";
                timer1.Enabled = true;
            }
        }

        private void folder_path_textBox_TextChanged(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(folder_path_textBox.Text))
            {
                clear_pictureBox.Visible = true;
            }
            else
            {
                clear_pictureBox.Visible = false;
            }
        }

        private void clear_pictureBox_Click(object sender, EventArgs e)
        {
            folder_path_textBox.Text = "";
            folder_FileFullNames.Clear();
        }

        private System.Data.DataTable GetInvoiceDataFromXML(string xmlPath)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable();


            //// 加载XML文档
            //XElement xmlDocument = XElement.Load(xmlPath);

            // 定义DataTable的列
            dataTable.Columns.Add("发票号码", typeof(string));                                   //发票号码 InvoiceNumber
            dataTable.Columns.Add("开票日期", typeof(string));                                  //开票日期 IssueTime
            dataTable.Columns.Add("销售方纳税识别号", typeof(string));                         //销售方纳税识别号 SellerIdNum
            dataTable.Columns.Add("销售方名称", typeof(string));                              //销售方名称 SellerName
            dataTable.Columns.Add("销售方地址", typeof(string));                             //销售方地址 SellerAddr
            dataTable.Columns.Add("销售方电话号码", typeof(string));                        //销售方电话号码 SellerTelNum
            dataTable.Columns.Add("销售方开户银行", typeof(string));                       //销售方开户银行 SellerBankName
            dataTable.Columns.Add("销售方银行账号", typeof(string));                      //销售方银行账号 SellerBankAccNum
            dataTable.Columns.Add("购买方纳税识别号", typeof(string));                   //购买方纳税识别号 BuyerIdNum
            dataTable.Columns.Add("购买方名称", typeof(string));                        //购买方名称 BuyerAddr
            dataTable.Columns.Add("购买方地址", typeof(string));                       //购买方地址 BuyerName
            dataTable.Columns.Add("购买方电话号码", typeof(string));                  //购买方电话号码 BuyerTelNum
            dataTable.Columns.Add("购买方开户银行", typeof(string));                 //购买方开户银行 BuyerBankName
            dataTable.Columns.Add("购买方银行账号", typeof(string));                //购买方银行账号 BuyerBankAccNum
            dataTable.Columns.Add("不含税价格", typeof(string));                   //不含税价格 TotalAmWithoutTax
            dataTable.Columns.Add("税额", typeof(string));                        //税额 TotalTaxAm
            dataTable.Columns.Add("含税价格", typeof(string));                   //含税价格 TotalTax-includedAmount
            dataTable.Columns.Add("项目名称", typeof(string));                  //项目名称 ItemName
            dataTable.Columns.Add("发票类型", typeof(string));                 //发票类型 GeneralOrSpecialVATLabelName
            dataTable.Columns.Add("发票监制税务机关", typeof(string));        //发票监制税务机关 TaxBureauName
            dataTable.Columns.Add("电子发票文件路径", typeof(string));       //电子发票文件路径


            // 提取数据
            XElement eInvoice = XElement.Load(xmlPath);
            if (eInvoice != null)
            {
                XElement taxSupervisionInfo = eInvoice.Element("TaxSupervisionInfo");
                XElement eInvoiceData = eInvoice.Element("EInvoiceData");
                XElement header = eInvoice.Element("Header");
                XElement? generalOrSpecialVAT = header?.Element("InherentLabel")?.Element("GeneralOrSpecialVAT");

                // 添加行到DataTable
                dataTable.Rows.Add
                    (
                        taxSupervisionInfo.Element("InvoiceNumber")?.Value,
                        taxSupervisionInfo.Element("IssueTime")?.Value,
                        eInvoiceData.Element("SellerInformation").Element("SellerIdNum")?.Value,
                        eInvoiceData.Element("SellerInformation").Element("SellerName")?.Value,
                        eInvoiceData.Element("SellerInformation").Element("SellerAddr")?.Value,
                        eInvoiceData.Element("SellerInformation").Element("SellerTelNum")?.Value,
                        eInvoiceData.Element("SellerInformation").Element("SellerBankName")?.Value,
                        eInvoiceData.Element("SellerInformation").Element("SellerBankAccNum")?.Value,
                        eInvoiceData.Element("BuyerInformation").Element("BuyerIdNum")?.Value,
                        eInvoiceData.Element("BuyerInformation").Element("BuyerName")?.Value,
                        eInvoiceData.Element("BuyerInformation").Element("BuyerAddr")?.Value,
                        eInvoiceData.Element("BuyerInformation").Element("BuyerTelNum")?.Value,
                        eInvoiceData.Element("BuyerInformation").Element("BuyerBankName")?.Value,
                        eInvoiceData.Element("BuyerInformation").Element("BuyerBankAccNum")?.Value,
                        eInvoiceData.Element("BasicInformation").Element("TotalAmWithoutTax")?.Value,
                        eInvoiceData.Element("BasicInformation").Element("TotalTaxAm")?.Value,
                        eInvoiceData.Element("BasicInformation").Element("TotalTax-includedAmount")?.Value,
                        eInvoiceData.Element("IssuItemInformation").Element("ItemName")?.Value,
                        generalOrSpecialVAT?.Element("LabelName")?.Value,
                        taxSupervisionInfo.Element("TaxBureauName")?.Value,
                        xmlPath
                    );
            }
            return dataTable;
        }

        private void WriteToExcel(System.Data.DataTable dataTable)
        {
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;
            Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            bool isFaPiaoSheetExist = false;
            foreach (Excel.Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Name == "_FaPiao" && isFieldExist("电子发票文件路径") && isFieldExist("发票类型"))
                {
                    isFaPiaoSheetExist = true;
                }
                else if (sheet.Name == "_FaPiao")
                {
                    sheet.Name = "_FaPiao_Original";
                }
                else
                {
                    continue;
                }
            }
            if (!isFaPiaoSheetExist)
            {
                Excel.Worksheet addSheet = workbook.Worksheets.Add(Before: workbook.ActiveSheet);
                addSheet.Name = "_FaPiao";
                addSheet.Activate();
                for (int t = 0; t < dataTable.Columns.Count; t++)
                {
                    workbook.ActiveSheet.Cells[1, t + 1] = dataTable.Columns[t].ColumnName;
                }
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        workbook.ActiveSheet.Cells[i + 2, j + 1].NumberFormat = "@";
                        workbook.ActiveSheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j];
                        if (workbook.ActiveSheet.Cells[1, j + 1].Value == "电子发票文件路径")
                        {
                            Excel.Worksheet sht = workbook.ActiveSheet;
                            string str = Path.GetDirectoryName(sht.Cells[i + 2, j + 1].Value);
                            sht.Hyperlinks.Add(sht.Cells[i + 2, j + 1], str, Type.Missing, Type.Missing, str);
                        }
                    }
                }
            }
            else
            {
                workbook.Worksheets["_FaPiao"].Activate();
                long usedRow = workbook.ActiveSheet.UsedRange.Rows.Count;
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        workbook.ActiveSheet.Cells[usedRow + i + 1, j + 1].NumberFormat = "@";
                        workbook.ActiveSheet.Cells[usedRow + i + 1, j + 1] = dataTable.Rows[i][j];
                        if (workbook.ActiveSheet.Cells[1, j + 1].Value == "电子发票文件路径")
                        {
                            Excel.Worksheet sht = workbook.ActiveSheet;
                            string str = Path.GetDirectoryName(sht.Cells[usedRow + i + 1, j + 1].Value);
                            sht.Hyperlinks.Add(sht.Cells[usedRow + i + 1, j + 1], str, Type.Missing, Type.Missing, str);
                        }
                    }
                }
            }
            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ActiveWorkbook.RefreshAll();
        }

        private bool isFieldExist(string fieldName)
        {
            Excel.Worksheet activeSheet = ThisAddIn.app.ActiveSheet;
            foreach (Excel.Range cell in activeSheet.Range[activeSheet.Cells[1, 1], activeSheet.Cells[1, activeSheet.UsedRange.Columns.Count]])
            {
                if (cell.Value == fieldName)
                {
                    return true;
                }
            }
            return false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Invoke(new System.Action(() =>
            {
                batch_result_label.Text = "";
                single_result_label.Text = "";
            }));
            timer1.Enabled = false;
        }

        private void reset_button_Click(object sender, EventArgs e)
        {
            single_result_label.Text = "";
            xml_path_textBox.Text = "";
            begin_pictureBox.Visible = false;
            begin_pictureBox.Enabled = false;
            preview_pictureBox.Visible = false;
            preview_pictureBox.Enabled = false;
            next_pictureBox.Visible = false;
            next_pictureBox.Enabled = false;
            last_pictureBox.Visible = false;
            last_pictureBox.Enabled = false;
            sequence_label.Text = "";
            single_FileFullNames.Clear();
        }

        /// <summary>
        /// 以下为pdf电子发票读取功能模块
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>


        //下一个
        private void next_pictureBox1_Click(object sender, EventArgs e)
        {
            if (pdfSingle_FileFullNames.Count == 0) return;
            int currentSequence = int.Parse(sequence_label1.Text);
            pdf_path_textBox.Text = pdfSingle_FileFullNames[currentSequence];
            sequence_label1.Text = Convert.ToString(currentSequence + 1);
        }

        //上一个
        private void preview_pictureBox1_Click(object sender, EventArgs e)
        {
            if (pdfSingle_FileFullNames.Count == 0) return;
            int currentSequence = int.Parse(sequence_label1.Text);
            pdf_path_textBox.Text = pdfSingle_FileFullNames[currentSequence - 2];
            sequence_label1.Text = Convert.ToString(currentSequence - 1);
        }

        //最后一个
        private void last_pictureBox1_Click(object sender, EventArgs e)
        {
            if (pdfSingle_FileFullNames.Count == 0) return;
            pdf_path_textBox.Text = pdfSingle_FileFullNames[pdfSingle_FileFullNames.Count - 1];
            sequence_label1.Text = pdfSingle_FileFullNames.Count.ToString();
        }

        //第一个
        private void begin_pictureBox1_Click(object sender, EventArgs e)
        {
            if (pdfSingle_FileFullNames.Count == 0) return;
            pdf_path_textBox.Text = pdfSingle_FileFullNames[0];
            sequence_label1.Text = "1";
        }

        //顺序号变化时
        private void sequence_label1_TextChanged(object sender, EventArgs e)
        {
            int fileCount = pdfSingle_FileFullNames.Count;
            if (sequence_label1.Text == "1" && fileCount > 1)
            {
                begin_pictureBox1.Enabled = false;
                begin_pictureBox1.Visible = false;
                preview_pictureBox1.Enabled = false;
                preview_pictureBox1.Visible = false;
                next_pictureBox1.Enabled = true;
                next_pictureBox1.Visible = true;
                last_pictureBox1.Enabled = true;
                last_pictureBox1.Visible = true;
            }
            else if (sequence_label1.Text == fileCount.ToString() && fileCount > 1)
            {
                begin_pictureBox1.Enabled = true;
                begin_pictureBox1.Visible = true;
                preview_pictureBox1.Enabled = true;
                preview_pictureBox1.Visible = true;
                next_pictureBox1.Enabled = false;
                next_pictureBox1.Visible = false;
                last_pictureBox1.Enabled = false;
                last_pictureBox1.Visible = false;
            }
            else if (string.IsNullOrEmpty(sequence_label1.Text))
            {
                begin_pictureBox1.Enabled = false;
                begin_pictureBox1.Visible = false;
                preview_pictureBox1.Enabled = false;
                preview_pictureBox1.Visible = false;
                next_pictureBox1.Enabled = false;
                next_pictureBox1.Visible = false;
                last_pictureBox1.Enabled = false;
                last_pictureBox1.Visible = false;
            }
            else if (sequence_label1.Text == fileCount.ToString() && fileCount == 1)
            {
                begin_pictureBox1.Enabled = false;
                begin_pictureBox1.Visible = false;
                preview_pictureBox1.Enabled = false;
                preview_pictureBox1.Visible = false;
                next_pictureBox1.Enabled = false;
                next_pictureBox1.Visible = false;
                last_pictureBox1.Enabled = false;
                last_pictureBox1.Visible = false;
            }
            else
            {
                begin_pictureBox1.Enabled = true;
                begin_pictureBox1.Visible = true;
                preview_pictureBox1.Enabled = true;
                preview_pictureBox1.Visible = true;
                next_pictureBox1.Enabled = true;
                next_pictureBox1.Visible = true;
                last_pictureBox1.Enabled = true;
                last_pictureBox1.Visible = true;
            }
        }

        private void pdf_path_textBox_DoubleClick(object sender, EventArgs e)
        {
            openFileDialog1.Title = "请选择要导入的电子发票（PDF或图片）";
            openFileDialog1.Filter = "发票文件(*.pdf;*.jpg;*.jpeg;*.png;*.bmp)|*.pdf;*.jpg;*.jpeg;*.png;*.bmp|PDF文件(*.pdf)|*.pdf|图片文件(*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp";
            openFileDialog1.Multiselect = true;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                sequence_label1.Text = string.Empty;
                pdfSingle_FileFullNames.Clear();
                pdfSingle_FileFullNames = openFileDialog1.FileNames.ToList();
                pdf_path_textBox.Text = pdfSingle_FileFullNames[0];
                sequence_label1.Text = "1";
                single_result_label1.Text = $"共{pdfSingle_FileFullNames.Count}个文件";
                if (pdfSingle_FileFullNames.Count > 1)
                {
                    int fileCount = pdfSingle_FileFullNames.Count;
                    if (sequence_label1.Text == "1")
                    {
                        begin_pictureBox1.Enabled = false;
                        preview_pictureBox1.Enabled = false;
                        next_pictureBox1.Enabled = true;
                        last_pictureBox1.Enabled = true;
                    }
                    else if (sequence_label1.Text == fileCount.ToString())
                    {
                        begin_pictureBox1.Enabled = true;
                        preview_pictureBox1.Enabled = true;
                        next_pictureBox1.Enabled = false;
                        last_pictureBox1.Enabled = false;
                    }
                    else
                    {
                        begin_pictureBox1.Enabled = true;
                        preview_pictureBox1.Enabled = true;
                        next_pictureBox1.Enabled = true;
                        last_pictureBox1.Enabled = true;
                    }
                }
                else
                {
                    begin_pictureBox1.Enabled = false;
                    preview_pictureBox1.Enabled = false;
                    next_pictureBox1.Enabled = false;
                    last_pictureBox1.Enabled = false;
                }
            }
            else
            {
                return;
            }
        }

        public Bitmap RenderPdfToImage(string pdfPath, int pageIndex,
            int targetWidth = 1080, int targetHeight = 1920,
            bool removeTransparency = true)
        {
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF文件不存在", pdfPath);

            // 创建文档阅读器
            using var docReader = DocLib.Instance.GetDocReader(
                pdfPath,
                new PageDimensions(targetWidth, targetHeight));

            // 验证页面索引
            if (pageIndex < 0 || pageIndex >= docReader.GetPageCount())
                throw new ArgumentOutOfRangeException(nameof(pageIndex),
                    $"页面索引必须在0到{docReader.GetPageCount() - 1}之间");

            // 获取页面阅读器
            using var pageReader = docReader.GetPageReader(pageIndex);

            // 获取图像数据（可选择移除透明度）
            byte[] rawBytes = removeTransparency
                ? pageReader.GetImage(new NaiveTransparencyRemover(255, 255, 255)) // 白色背景
                : pageReader.GetImage();

            // 获取实际渲染尺寸
            int width = pageReader.GetPageWidth();
            int height = pageReader.GetPageHeight();

            // 创建并返回位图
            return CreateBitmapFromBytes(rawBytes, width, height);
        }

        private static Bitmap CreateBitmapFromBytes(byte[] rawBytes, int width, int height)
        {
            // 验证尺寸有效性
            if (width <= 0 || height <= 0)
                throw new InvalidOperationException($"无效的图像尺寸: {width}x{height}");

            // 创建位图对象
            var bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);

            // 锁定位图进行写入
            var rect = new Rectangle(0, 0, width, height);
            var bmpData = bmp.LockBits(rect, ImageLockMode.WriteOnly, bmp.PixelFormat);

            try
            {
                // 将字节数据复制到位图
                Marshal.Copy(rawBytes, 0, bmpData.Scan0, rawBytes.Length);
            }
            finally
            {
                // 确保始终解锁位图
                bmp.UnlockBits(bmpData);
            }

            return bmp;
        }


        private void pdf_path_textBox_TextChanged(object sender, EventArgs e)
        {
            string path = pdf_path_textBox.Text.Trim();

            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                try
                {
                    // 判断文件类型
                    string extension = Path.GetExtension(path).ToLower();

                    if (extension == ".pdf")
                    {
                        // PDF文件：渲染为图片
                        var bitmap = RenderPdfToImage(path, pageIndex: 0);
                    pictureBoxPdf.Image = bitmap;
                    }
                    else if (extension == ".jpg" || extension == ".jpeg" || extension == ".png" || extension == ".bmp")
                    {
                        // 图片文件：直接加载
                        // 释放旧图片
                        pictureBoxPdf.Image?.Dispose();
                        pictureBoxPdf.Image = Image.FromFile(path);
                    }
                    
                    pictureBoxPdf.SizeMode = PictureBoxSizeMode.Zoom; // 设置图片适应PictureBox大小
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"加载失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                // 清空 PictureBox
                pictureBoxPdf.Image?.Dispose();
                pictureBoxPdf.Image = null;

                // 其他控件的隐藏逻辑（与你的原始代码一致）
                if (single_FileFullNames.Count == 0)
                {
                    sequence_label1.Text = string.Empty;
                    begin_pictureBox1.Enabled = false;
                    begin_pictureBox1.Visible = false;
                    preview_pictureBox1.Enabled = false;
                    preview_pictureBox1.Visible = false;
                    next_pictureBox1.Enabled = false;
                    next_pictureBox1.Visible = false;
                    last_pictureBox1.Enabled = false;
                    last_pictureBox1.Visible = false;
                }
            }
        }

        private void pdfFolder_path_textBox_DoubleClick(object sender, EventArgs e)
        {
            folderBrowserDialog1.Description = "请选择发票文件（PDF或图片）所在文件夹";
            folderBrowserDialog1.ShowDialog();
            if (folderBrowserDialog1.SelectedPath != "")
            {
                folder_FileFullNames.Clear();
                pdfFolder_FileFullNames.Clear();
                string selectedFolder = folderBrowserDialog1.SelectedPath;
                pdfFolder_path_textBox.Text = selectedFolder;
                DirectoryInfo folder = new DirectoryInfo(selectedFolder);
                
                // 获取所有支持的文件类型
                SearchOption searchOption = pdfSubfolder_checkBox.Checked ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                var pdfFiles = folder.GetFiles("*.pdf", searchOption);
                var jpgFiles = folder.GetFiles("*.jpg", searchOption);
                var jpegFiles = folder.GetFiles("*.jpeg", searchOption);
                var pngFiles = folder.GetFiles("*.png", searchOption);
                var bmpFiles = folder.GetFiles("*.bmp", searchOption);
                
                // 合并所有文件并按名称排序
                pdfFolder_FileFullNames = pdfFiles.Concat(jpgFiles).Concat(jpegFiles).Concat(pngFiles).Concat(bmpFiles)
                    .OrderBy(f => f.Name)
                    .Select(file => file.FullName)
                    .ToList();
                    
                pdfBatch_result_label.Text = $"找到{pdfFolder_FileFullNames.Count}个发票文件";
            }
            else
            {
                pdfFolder_path_textBox.Text = "";
                pdfBatch_result_label.Text = "未选择文件夹";
            }
        }

        private void pdfBatch_run_button_Click(object sender, EventArgs e)
        {
            if (pdfFolder_FileFullNames.Count > 0 && Directory.Exists(pdfFolder_path_textBox.Text))
            {
                int o = 1;
                int successCount = 0;
                int failCount = 0;
                
                foreach (string file in pdfFolder_FileFullNames)
                {
                    string fileName = Path.GetFileName(file);
                    pdfBatch_result_label.Text = $"正在导入第{o}/{pdfFolder_FileFullNames.Count}个文件：{fileName}...";
                    Application.DoEvents(); // 强制更新UI
                    
                    try
                    {
                        System.Data.DataTable dataTable = GetInvoiceDataFromFile(file);
                    if (dataTable.Rows.Count > 0)
                    {
                        WriteToExcel(dataTable);
                            successCount++;
                        }
                        else
                        {
                            failCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        Console.WriteLine($"处理文件 {fileName} 失败: {ex.Message}");
                    }
                    
                    o++;
                }
                
                pdfBatch_result_label.Text = $"批量导入完成！共{pdfFolder_FileFullNames.Count}个文件，成功{successCount}个，失败{failCount}个";
                timer1.Enabled = true;
            }
            else
            {
                pdfBatch_result_label.Text = "文件夹内没有发票文件，请核对！";
                timer1.Enabled = true;
            }
        }

        // 单个文件导入按钮
        private void btnPdfRun_Click(object sender, EventArgs e)
        {
            string filePath = pdf_path_textBox.Text.Trim();
            
            System.Diagnostics.Debug.WriteLine($"【调试】开始识别发票: {filePath}");
            
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show("请先选择一个有效的发票文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            try
            {
                single_result_label1.Text = "正在识别发票...";
                Application.DoEvents();
                
                System.Diagnostics.Debug.WriteLine($"【调试】调用 GetInvoiceDataFromFile");
                System.Data.DataTable dataTable = GetInvoiceDataFromFile(filePath);
                System.Diagnostics.Debug.WriteLine($"【调试】返回的行数: {dataTable.Rows.Count}");
                
                if (dataTable.Rows.Count > 0)
                {
                    WriteToExcel(dataTable);
                    single_result_label1.Text = "发票识别并导入成功！";
                    timer1.Enabled = true;
                }
                else
                {
                    single_result_label1.Text = "识别失败，未获取到发票数据";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"【调试】异常: {ex.ToString()}");
                MessageBox.Show($"导入失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                single_result_label1.Text = "导入失败";
            }
        }
        
        // 重置按钮
        private void btnPdfReset_Click(object sender, EventArgs e)
        {
            pdf_path_textBox.Text = "";
            sequence_label1.Text = "";
            single_result_label1.Text = "";
            pdfSingle_FileFullNames.Clear();
            pictureBoxPdf.Image?.Dispose();
            pictureBoxPdf.Image = null;
            
            // 隐藏导航按钮
            begin_pictureBox1.Enabled = false;
            begin_pictureBox1.Visible = false;
            preview_pictureBox1.Enabled = false;
            preview_pictureBox1.Visible = false;
            next_pictureBox1.Enabled = false;
            next_pictureBox1.Visible = false;
            last_pictureBox1.Enabled = false;
            last_pictureBox1.Visible = false;
        }
        
        // 智能识别文件类型并调用相应的处理方法
        private System.Data.DataTable GetInvoiceDataFromFile(string filePath)
        {
            string extension = Path.GetExtension(filePath).ToLower();
            
            if (extension == ".pdf")
            {
                return GetInvoiceDataFromPDF(filePath);
            }
            else if (extension == ".jpg" || extension == ".jpeg" || extension == ".png" || extension == ".bmp")
            {
                return GetInvoiceDataFromImage(filePath);
            }
            else
            {
                // 不支持的文件类型，返回空表
                DataTable emptyTable = CreateTableSchema();
                return emptyTable;
            }
        }

        private System.Data.DataTable GetInvoiceDataFromPDF(string pdfPath)
        {
            System.Diagnostics.Debug.WriteLine($"【调试】GetInvoiceDataFromPDF 开始: {pdfPath}");
            DataTable dataTable = CreateTableSchema();
            DataRow row = dataTable.NewRow();
            row["电子发票文件路径"] = pdfPath;
            
            try
            {
                if (_invoiceService == null)
                {
                    System.Diagnostics.Debug.WriteLine("【调试】错误: _invoiceService 为 null");
                    MessageBox.Show("发票识别服务未初始化，无法处理PDF文件。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    dataTable.Rows.Add(row);
                    return dataTable;
                }
                
                System.Diagnostics.Debug.WriteLine("【调试】调用 RecognizePdfInvoiceFirstPage");
                // 使用PaddleOCR识别PDF发票
                var invoiceInfo = _invoiceService.RecognizePdfInvoiceFirstPage(pdfPath);
                System.Diagnostics.Debug.WriteLine("【调试】RecognizePdfInvoiceFirstPage 返回");
                
                if (invoiceInfo != null)
                {
                    // 调试：输出识别到的发票类型信息
                    System.Diagnostics.Debug.WriteLine($"【调试】识别到的发票类型: '{invoiceInfo.InvoiceType}'");
                    System.Diagnostics.Debug.WriteLine($"【调试】发票号码: '{invoiceInfo.InvoiceNumber}'");
                    
                    // 填充数据 - 使用新的InvoiceInfo结构
                    row["发票号码"] = invoiceInfo.InvoiceNumber ?? "";
                    row["开票日期"] = invoiceInfo.InvoiceDate ?? "";
                    row["销售方名称"] = invoiceInfo.Seller?.Name ?? "";
                    row["销售方纳税识别号"] = invoiceInfo.Seller?.TaxId ?? "";
                    row["购买方名称"] = invoiceInfo.Buyer?.Name ?? "";
                    row["购买方纳税识别号"] = invoiceInfo.Buyer?.TaxId ?? "";
                    row["不含税价格"] = invoiceInfo.Amount.ToString("F2");
                    row["税额"] = invoiceInfo.TaxAmount.ToString("F2");
                    row["含税价格"] = invoiceInfo.TotalAmount.ToString("F2");
                    row["发票类型"] = invoiceInfo.InvoiceType ?? "";
                    
                    System.Diagnostics.Debug.WriteLine($"【调试】填充后的发票类型: '{row["发票类型"]}'");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("【调试】警告: invoiceInfo 为 null");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"【调试】处理异常: {ex.ToString()}");
                MessageBox.Show($"识别PDF发票时出错: {ex.Message}\n\n详细:\n{ex.ToString()}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                dataTable.Rows.Add(row);
            }
            
            return dataTable;
        }
        
        // 处理图片发票
        private System.Data.DataTable GetInvoiceDataFromImage(string imagePath)
        {
            System.Diagnostics.Debug.WriteLine($"【调试】GetInvoiceDataFromImage 开始: {imagePath}");
            DataTable dataTable = CreateTableSchema();
            DataRow row = dataTable.NewRow();
            row["电子发票文件路径"] = imagePath;
            
            try
            {
                if (_invoiceService == null)
                {
                    System.Diagnostics.Debug.WriteLine("【调试】错误: _invoiceService 为 null");
                    MessageBox.Show("发票识别服务未初始化，无法处理图片文件。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    dataTable.Rows.Add(row);
                    return dataTable;
                }
                
                System.Diagnostics.Debug.WriteLine("【调试】调用 RecognizeImageInvoice");
                // 使用PaddleOCR识别图片发票
                var invoiceInfo = _invoiceService.RecognizeImageInvoice(imagePath);
                System.Diagnostics.Debug.WriteLine("【调试】RecognizeImageInvoice 返回");
                
                if (invoiceInfo != null)
                {
                    // 调试：输出识别到的发票类型信息
                    System.Diagnostics.Debug.WriteLine($"【调试】图片-识别到的发票类型: '{invoiceInfo.InvoiceType}'");
                    System.Diagnostics.Debug.WriteLine($"【调试】图片-发票号码: '{invoiceInfo.InvoiceNumber}'");
                    
                    // 填充数据 - 使用新的InvoiceInfo结构
                    row["发票号码"] = invoiceInfo.InvoiceNumber ?? "";
                    row["开票日期"] = invoiceInfo.InvoiceDate ?? "";
                    row["销售方名称"] = invoiceInfo.Seller?.Name ?? "";
                    row["销售方纳税识别号"] = invoiceInfo.Seller?.TaxId ?? "";
                    row["购买方名称"] = invoiceInfo.Buyer?.Name ?? "";
                    row["购买方纳税识别号"] = invoiceInfo.Buyer?.TaxId ?? "";
                    row["不含税价格"] = invoiceInfo.Amount.ToString("F2");
                    row["税额"] = invoiceInfo.TaxAmount.ToString("F2");
                    row["含税价格"] = invoiceInfo.TotalAmount.ToString("F2");
                    row["发票类型"] = invoiceInfo.InvoiceType ?? "";
                    
                    System.Diagnostics.Debug.WriteLine($"【调试】图片-填充后的发票类型: '{row["发票类型"]}'");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("【调试】警告: 图片-invoiceInfo 为 null");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"【调试】图片-处理异常: {ex.ToString()}");
                MessageBox.Show($"识别图片发票时出错: {ex.Message}\n\n详细:\n{ex.ToString()}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                dataTable.Rows.Add(row);
            }
            
            return dataTable;
        }

        public class WordInfo
        {
            public string? Text { get; set; }
            public double X0 { get; set; }   //距离顶部坐标
            public double Y0 { get; set; }   //距离左侧坐标
            public double X1 { get; set; }   //距离底部坐标
            public double Y1 { get; set; }   //距离右侧坐标
        }


        private DataTable CreateTableSchema()
        {
            var dt = new DataTable();
            dt.Columns.Add("发票号码", typeof(string));
            dt.Columns.Add("开票日期", typeof(string));
            dt.Columns.Add("销售方名称", typeof(string));
            dt.Columns.Add("销售方纳税识别号", typeof(string));
            dt.Columns.Add("购买方名称", typeof(string));
            dt.Columns.Add("购买方纳税识别号", typeof(string));
            dt.Columns.Add("不含税价格", typeof(string));
            dt.Columns.Add("税额", typeof(string));
            dt.Columns.Add("含税价格", typeof(string));
            dt.Columns.Add("发票类型", typeof(string));
            dt.Columns.Add("电子发票文件路径", typeof(string));
            return dt;
        }
    }
}
