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
using Excel=Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using ZXing;
using ZXing.Common;
using ZXing.QrCode.Internal;
using System.Windows.Media.Imaging;


namespace ExcelAddIn
{
    public partial class Form5 : Form
    {
        public Form5()
        {
            InitializeComponent();
            picture_radioButton.Checked = true;
            folder_path_label.Text = string.Empty;
        }


        private void quit_button_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        List<string> files_fullname = new List<string>();

        private async void scan_button_Click(object sender, EventArgs e)
        {
            if (picture_radioButton.Checked)
            {
                openFileDialog1.Filter = "图片文件(*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp";
                openFileDialog1.Title = "请选择包含二维码的图片";
                openFileDialog1.Multiselect = true;

                if (openFileDialog1.ShowDialog() == DialogResult.OK && openFileDialog1.FileNames.Length > 0)
                {
                    picture_radioButton.Enabled = false;
                    webcam_radioButton.Enabled = false;
                    quit_button.Enabled = false;
                    scan_button.Enabled = false;

                    files_fullname = openFileDialog1.FileNames.ToList();
                    string folderPath = Path.GetDirectoryName(openFileDialog1.FileNames[0]);

                    // 在主线程上更新 UI
                    this.Invoke((MethodInvoker)delegate
                    {
                        folder_path_label.Text = $"已选文件夹：{folderPath}。正在读取中......";
                    });

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt.Columns.Add("fold_path", typeof(string));
                    dt.Columns.Add("file_name", typeof(string));
                    dt.Columns.Add("qr_sequence", typeof(string));
                    dt.Columns.Add("qr_content", typeof(string));

                    await Task.Run(() =>
                    {
                        foreach (string item in files_fullname)
                        {
                            using (Bitmap bitmap = new Bitmap(item))
                            {
                                Bitmap gray_bitmap = PreprocessImage(bitmap);
                                
                                if (ReadQRCode(gray_bitmap).Count > 0)
                                {
                                    for (int i = 0; i < ReadQRCode(gray_bitmap).Count; i++)
                                    {
                                        string[] details = new string[4]
                                        {
                                            Path.GetDirectoryName(item),
                                            Path.GetFileName(item),
                                            (i + 1).ToString(),
                                            ReadQRCode(gray_bitmap)[i]
                                        };
                                        dt.Rows.Add(details);
                                    }
                                }
                            }
                        }
                    });

                    // 在主线程上更新 UI
                    this.Invoke((MethodInvoker)delegate
                    {
                        folder_path_label.Text = "读取完成！";
                        if (dt.Rows.Count > 0)
                        {
                            WriteToExcel(dt);
                        }
                    });

                    // 在主线程上恢复按钮状态
                    this.Invoke((MethodInvoker)delegate
                    {
                        picture_radioButton.Enabled = true;
                        webcam_radioButton.Enabled = true;
                        quit_button.Enabled = true;
                        scan_button.Enabled = true;
                    });
                }
            }
        }


        //ZXing 二维码读取
        private List<string> ReadQRCode(Bitmap bitmap)
        {
            List<string> qrCodeContents = new List<string>();

            // 创建专用的QRCodeReader
            BarcodeReader reader = new BarcodeReader()
            {
                Options = new DecodingOptions
                {
                    // 设置尝试解码的次数
                    TryHarder = true,
                    // 指定字符集
                    PossibleFormats = new List<BarcodeFormat> { BarcodeFormat.QR_CODE },
                    CharacterSet = null
                }
            };
            try
            {
                // 解码图像中的多个二维码
                var results = reader.DecodeMultiple(bitmap);
                foreach (var result in results)
                {
                    if (result != null && !string.IsNullOrEmpty(result.Text))
                    {
                        qrCodeContents.Add(result.Text);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return qrCodeContents;
        }

        //ZXing判断一张图片中包含多少个二维码，该方法在功能实现中未使用
        private int CountQRCodes(Bitmap bitmap)
        {
            // 检查 bitmap 是否为 null
            if (bitmap == null)
            {
                return -1;
            }

            try
            {
                BarcodeReader reader = new BarcodeReader()
                {
                    Options = new DecodingOptions
                    {
                        // 设置尝试解码的次数
                        TryHarder = true,
                        // 指定字符集
                        PossibleFormats = new List<BarcodeFormat> { BarcodeFormat.QR_CODE },
                        CharacterSet = null
                    }
                };

                var results = reader.DecodeMultiple(bitmap);
                if (results != null && results.Any())
                {
                    foreach (var result in results)
                    {
                        if (result != null && !string.IsNullOrEmpty(result.Text))
                        {
                            Console.WriteLine(result.Text); // 打印出每个二维码的内容
                        }
                    }
                    return results.Count(); // 返回二维码的数量
                }
                else
                {
                    return 0; // 没有找到二维码，返回0
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return -1;
            }
        }

        // ZXing判断图片是否包含二维码，该方法在功能实现中未使用
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

        // 转换为灰度图像
        private Bitmap PreprocessImage(Bitmap bitmap)
        {            
            Bitmap grayBitmap = new Bitmap(bitmap.Width, bitmap.Height);
            for (int y = 0; y < bitmap.Height; y++)
            {
                for (int x = 0; x < bitmap.Width; x++)
                {
                    Color pixel = bitmap.GetPixel(x, y);
                    int grayValue = (int)(pixel.R * 0.299 + pixel.G * 0.587 + pixel.B * 0.114);
                    grayBitmap.SetPixel(x, y, Color.FromArgb(grayValue, grayValue, grayValue));
                }
            }
            return grayBitmap;
        }


        /*
         * 写入当前打开的工作簿中的_QR_Scan工作表
         * 1. 判断工作簿中是否有_QR_Scan工作表。
         * 2. 如果是激活扫码功能前已经有，则重命名原有表，并新建一个，并激活。
         * 3. 如果原先没有，则直接新建一个并激活。
         * 4. 如果激活扫码功能后再次使用扫码功能，且已有表，则直接激活。
         */
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

        private void webcam_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (webcam_radioButton.Checked)
            {
                folder_path_label.Visible = false;
                qr_image_pictureBox.Visible = true;
            }
        }

        private void picture_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (picture_radioButton.Checked)
            {
                folder_path_label.Visible = true;
                qr_image_pictureBox.Visible = false;
            }
        }
    }
}
