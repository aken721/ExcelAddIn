using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel=Microsoft.Office.Interop.Excel;
using ZXing;
using ZXing.Common;
using AForge.Video.DirectShow;
using System.Media;
using ExcelAddIn.Properties;





namespace ExcelAddIn
{
    public partial class Form5 : Form
    {
        DataTable dt = new DataTable();
        public Form5()
        {
            InitializeComponent();
            picture_radioButton.Checked = true;
            folder_path_label.Text = string.Empty;
            
            dt.Columns.Add("fold_path", typeof(string));
            dt.Columns.Add("file_name", typeof(string));
            dt.Columns.Add("qr_sequence", typeof(string));
            dt.Columns.Add("qr_content", typeof(string));
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

                    await Task.Run(() =>
                    {
                        foreach (string item in files_fullname)
                        {
                            using (Bitmap bitmap = new Bitmap(item))
                            {
                                Bitmap gray_bitmap = PreprocessImage(bitmap);
                                
                                if (ReadQRCode(gray_bitmap)!=null && ReadQRCode(gray_bitmap).Count > 0)
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
                        folder_path_label.Text = $"读取完成！成功读取{dt.Rows.Count.ToString()}个二维码";
                        if (dt.Rows.Count > 0)
                        {
                            WriteToExcel(dt);
                            dt.Clear();
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
            else if (webcam_radioButton.Checked)
            {

                isProcessing = false;
                FilterInfoCollection videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
                if (videoDevices.Count > 0)
                {
                    frame = null;
                    picture_radioButton.Enabled = false;
                    webcam_radioButton.Enabled = false;
                    quit_button.Enabled = false;
                    scan_button.Enabled = false;
                    cancel_button.Enabled = true;
                    folder_path_label.Text = "开始扫描二维码......";
                    videoDevice = new VideoCaptureDevice(videoDevices[0].MonikerString);
                    videoDevice.VideoResolution =videoDevice.VideoCapabilities[0]; 
                    videoSourcePlayer1.VideoSource = videoDevice;
                    videoDevice.Start();
                    videoSourcePlayer1.Start();
                    //videoDevice.NewFrame += new NewFrameEventHandler(videoDevice_NewFrame);
                    timer1.Interval = 1000;
                    timer1.Enabled = true;
                }
                else
                {
                    MessageBox.Show("未检测到摄像头设备！请使用图片识别方式扫描二维码");
                }
            }
        }

        //private void videoDevice_NewFrame(object sender, NewFrameEventArgs eventArgs)
        //{
        //    frame = (Bitmap)eventArgs.Frame.Clone(); // 使用frame变量
        //}

        private VideoCaptureDevice videoDevice;   // 新增摄像头扫码时，用于存储当前设备的变量
        private bool isProcessing = false;      // 添加标志变量
        private Bitmap frame;                  // 新增摄像头扫码时，用于存储当前帧的变量
        private readonly object _lockObject = new object();    // 锁对象，用于确保dt写入时现成安全

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (isProcessing) return;
            isProcessing = true;
            
            Task.Run(() =>
            {
                frame = videoSourcePlayer1.GetCurrentVideoFrame();
                if (frame == null) 
                {
                    isProcessing = false;
                    return;
                }
                

                using (Bitmap gray_bitmap = PreprocessImage(frame))
                {
                    List<string> results = ReadQRCode(gray_bitmap);
                    if (results != null && results.Count > 0 ) // 检查标志变量
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
                            //videoDevice.NewFrame -= videoDevice_NewFrame;
                            using (SoundPlayer scan_player = new SoundPlayer(Resources.ScanSound))
                            {
                                scan_player.Play();
                            }
                            // 停止摄像头和视频播放器
                            videoDevice.SignalToStop();
                            videoDevice.WaitForStop();
                            videoSourcePlayer1.SignalToStop();
                            videoSourcePlayer1.WaitForStop();
                            timer1.Enabled = false;
                        });

                        foreach (var result in results)
                        {
                            lock (_lockObject) // Ensure thread safety when accessing shared resource 'dt'
                            {
                                string[] details = new string[4]
                            {
                                "webcam",
                                "webcam",
                                (results.IndexOf(result) + 1).ToString(),
                                result.ToString()
                            };
                                dt.Rows.Add(details);
                            }                            
                        }
                        if (dt.Rows.Count > 0)
                        {
                            WriteToExcel(dt);
                            dt.Clear();

                            // 写入完成后在主线程恢复按钮状态
                            this.Invoke((MethodInvoker)delegate
                            {
                                picture_radioButton.Enabled = true;
                                webcam_radioButton.Enabled = true;
                                quit_button.Enabled = true;
                                scan_button.Enabled = true;
                                cancel_button.Enabled = false;
                                folder_path_label.Text = "二维码读取完毕！";
                            });
                        }
                    }
                }
                isProcessing = false; // 处理完毕，重置标志
            });            
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
                if (results == null)
                {
                    return null;
                }
                foreach (var result in results)
                {
                    if (result != null && !string.IsNullOrEmpty(result.Text))
                    {
                        qrCodeContents.Add(result.Text);
                    }
                }
                return qrCodeContents;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

            
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


        ///<summary>
        /// 写入当前打开的工作簿中的_QR_Scan工作表
        /// 1. 判断工作簿中是否有_QR_Scan工作表。
        /// 2. 如果是激活扫码功能前已经有，则重命名原有表，并新建一个，并激活。
        /// 3. 如果原先没有，则直接新建一个并激活。
        /// 4. 如果激活扫码功能后再次使用扫码功能，且已有表，则直接激活。
        /// </summary>

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
                        sheet.Cells[row + i, j + 1].NumberFormat = "@";
                        sheet.Cells[row + i, j + 1].Value = str;
                        if (j == 1 && sheet.Cells[row + i, j + 1].Value!="webcam")
                        {
                            sheet.Hyperlinks.Add(sheet.Cells[row + i, j + 1], Path.Combine(data.Rows[i][0].ToString(), str), Type.Missing, Type.Missing, str);
                        }

                        if (str.StartsWith("http") || str.StartsWith("mailto:"))
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
                videoSourcePlayer1.Visible = true;
            }
        }

        private void picture_radioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (picture_radioButton.Checked)
            {
                videoSourcePlayer1.Visible = false;
                cancel_button.Enabled = false;
            }
        }

        private void cancel_button_Click(object sender, EventArgs e)
        {
            // 停止 VideoSourcePlayer
            if (videoSourcePlayer1.IsRunning)
            {
                videoSourcePlayer1.SignalToStop(); // 发出停止信号
                videoSourcePlayer1.WaitForStop();   // 等待线程停止
            }

            // 停止 VideoCaptureDevice
            if (videoDevice != null && videoDevice.IsRunning)
            {
                videoDevice.SignalToStop(); // 发出停止信号
                videoDevice.WaitForStop();   // 等待线程停止
            }

            // 恢复按钮状态
            scan_button.Enabled = true;
            quit_button.Enabled = true;
            cancel_button.Enabled = false;
            isProcessing=false;

            // 清空文件夹路径标签
            if (!string.IsNullOrEmpty(folder_path_label.Text))
            {
                folder_path_label.Text = "";
            }
        }


        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 停止 VideoSourcePlayer
            if (videoSourcePlayer1.IsRunning)
            {
                videoSourcePlayer1.SignalToStop(); // 发出停止信号
                videoSourcePlayer1.WaitForStop();   // 等待线程停止
                videoSourcePlayer1.VideoSource = null;
            }

            // 停止 VideoCaptureDevice
            if (videoDevice != null && videoDevice.IsRunning)
            {
                videoDevice.SignalToStop(); // 发出停止信号
                videoDevice.WaitForStop();   // 等待线程停止
            }
        }
    }
}
