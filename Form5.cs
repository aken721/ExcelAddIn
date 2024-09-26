using System;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel=Microsoft.Office.Interop.Excel;
using ZXing;
using Microsoft.Office.Tools.Excel;


namespace ExcelAddIn
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
            picture_radioButton.Checked = true;
            folder_path_label.Text = string.Empty;
            qr_image_listView.View = View.Details;
            qr_image_listView.FullRowSelect = true;
            qr_image_listView.Columns.Add("Name", 150);
            qr_image_listView.Columns.Add("QrCodeCount", 100);
            qr_image_listView.Clear();
        }


        private void quit_button_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        List<string> files_fullname = new List<string>();

        private void scan_button_Click(object sender, EventArgs e)
        {
            if(picture_radioButton.Checked)
            {
                openFileDialog1.Filter = "图片文件(*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp"; 
                openFileDialog1.Title = "请选择包含二维码的图片";
                openFileDialog1.Multiselect = true;
                openFileDialog1.ShowDialog();
                if (openFileDialog1.FileName != "") 
                {
                    files_fullname=openFileDialog1.FileNames.ToList();
                    folder_path_label.Text="已选文件夹："+Path.GetDirectoryName(openFileDialog1.FileName[0].ToString());
                }

                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("fold_path", typeof(string));
                dt.Columns.Add("file_name", typeof(string));
                dt.Columns.Add("qr_sequence", typeof(string));
                dt.Columns.Add("qr_content", typeof(string));

                foreach (string item in files_fullname)
                {
                    MessageBox.Show(item);
                    Bitmap bitmap = new Bitmap(item);
                    MessageBox.Show(bitmap.Width.ToString()+"*"+bitmap.Height.ToString());
                    if(CountQRCodes(bitmap)>0)
                    {
                        for (int i = 0; i < CountQRCodes(bitmap); i++)
                        {
                            string[] details=new string[4]
                            {
                                Path.GetDirectoryName(item),
                                Path.GetFileName(item),
                                i+1.ToString(),
                                ReadQRCode(bitmap)[i]
                            };
                            dt.Rows.Add(details);
                        }
                    }                    
                }
                if(dt.Rows.Count > 0)
                {
                    WriteToExcel(dt);
                }
            }
        }

        

        private List<string> ReadQRCode(Bitmap bitmap)
        {
            List<string> qrCodeContents = new List<string>();
            BarcodeReader reader = new BarcodeReader();
            var results = reader.DecodeMultiple(bitmap);
            if (results.Any())
            {
                foreach (var result in results)
                {
                    qrCodeContents.Add(result.Text); // 每个二维码的内容加入到列表
                }
                return qrCodeContents; // 返回二维码的数量
            }
            else
            {
                return qrCodeContents; // 没有找到二维码，返回空列表
            }
        }

        Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

        private void WriteToExcel(System.Data.DataTable data)
        {
            ThisAddIn.app.Application.ScreenUpdating=false;
            ThisAddIn.app.Application.DisplayAlerts=false;
            

            if (!ThisAddIn.Global.created_qr_sheet)
            {
                Excel.Worksheet worksheet = workbook.Worksheets.Add(Before: workbook.Worksheets[1]);
                foreach (Excel.Worksheet item in workbook.Worksheets)
                {
                    if (item.Name == "_QR_Scan")
                    {
                        item.Name = "_QR_Scan_备份";
                        break;
                    }
                }
                worksheet.Name = "_QR_Scan";
                worksheet.Activate();
                worksheet.Range["A1"].Value = "目录";
                worksheet.Range["B1"].Value = "文件名称";
                worksheet.Range["C1"].Value = "二维码编号";
                worksheet.Range["D1"].Value = "二维码内容";
                ThisAddIn.Global.created_qr_sheet = true;
            }
            else
            {
                workbook.Worksheets["_QR_Scan"].Activate();
            }

            Excel.Worksheet sheet = workbook.Worksheets["_QR_Scan"];
            try
            {
                int row = sheet.UsedRange.Rows.Count + 1;

                for (int i = 0; i < data.Rows.Count; i++)
                {
                    for (int j = 0; j < data.Columns.Count; j++)
                    {
                        string str = data.Rows[i][j].ToString();
                        sheet.Cells[row + i, j + 1].Value = str;
                        if (j == 1)
                        {
                            sheet.Hyperlinks.Add(sheet.Cells[row + i, j + 1], Path.Combine(data.Rows[i][0].ToString(), str), Type.Missing, Type.Missing, str);
                        }

                        if (str.StartsWith("http"))
                        {
                            sheet.Hyperlinks.Add(sheet.Cells[row + i, j + 1], str, Type.Missing, Type.Missing, str);
                        }
                    }
                }
            }
            catch(Exception ex) 
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ThisAddIn.app.Application.ScreenUpdating = true;
                ThisAddIn.app.Application.DisplayAlerts = true;
            }
        }

        private void folder_path_label_Click(object sender, EventArgs e)
        {

        }

        private void webcam_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (webcam_radioButton.Checked)
            {
                qr_image_listView.Visible = false;
                folder_path_label.Visible = false;
                qr_image_pictureBox.Visible = true;
            }
        }

        private void picture_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (picture_radioButton.Checked)
            {
                qr_image_listView.Visible = true;
                folder_path_label.Visible = true;
                qr_image_pictureBox.Visible = false;
            }
        }

        //判断一张图片中包含多少个二维码
        private int CountQRCodes(Bitmap bitmap)
        {
            // 检查 bitmap 是否为 null
            if (bitmap == null)
            {
                MessageBox.Show("Bitmap is null. Please provide a valid bitmap.");
                return -1;
            }

            try
            {
                BarcodeReader reader = new BarcodeReader();
                var results = reader.DecodeMultiple(bitmap);
                if (results.Any())
                {
                    foreach (var result in results)
                    {
                        Console.WriteLine(result.Text); // 打印出每个二维码的内容
                    }
                    return results.Count(); // 返回二维码的数量
                }
                else
                {
                    Console.WriteLine("No QR Code Found");
                    return 0; // 没有找到二维码，返回0
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return -1;
            }            
        }

        // 判断图片是否包含二维码
        private bool IsQR(Bitmap bitmap)
        {
            BarcodeReader reader = new BarcodeReader();
            Result result = reader.Decode(bitmap);
            if (result != null)
            {
                Console.WriteLine(result.Text);
                return true;
            }
            else
            {
                Console.WriteLine("No QR Code Found");
                return false;
            }
        }
    }
}
