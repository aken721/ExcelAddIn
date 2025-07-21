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

            // 视频框绘制事件
            videoSourcePlayer1.Paint += videoSourcePlayer1_Paint;
        }


        private void quit_button_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        List<string> files_fullname = new List<string>();

        private int currentCameraIndex = 0;             // 当前摄像头索引
        private FilterInfoCollection videoDevices;     // 摄像头设备集合
        private List<Rectangle> qrRectangles = new List<Rectangle>();
        private Pen qrBoxPen = new Pen(Color.LimeGreen, 3);
        private int frameWidth = 0;
        private int frameHeight = 0;
        private VideoCaptureDevice videoDevice;   // 新增摄像头扫码时，用于存储当前设备的变量
        private bool isProcessing = false;      // 添加标志变量
        private Bitmap frame;                  // 新增摄像头扫码时，用于存储当前帧的变量
        private readonly object _lockObject = new object();    // 锁对象，用于确保dt写入时现成安全

        // 添加摄像头切换方法
        private void switchCamera_button_Click(object sender, EventArgs e)
        {
            if (videoDevices == null || videoDevices.Count < 2) return;

            // 停止当前设备
            if (videoDevice != null && videoDevice.IsRunning)
            {
                videoDevice.SignalToStop();
                videoDevice.WaitForStop();
            }

            // 切换到下一个摄像头
            currentCameraIndex = (currentCameraIndex + 1) % videoDevices.Count;

            // 启动新设备
            videoDevice = new VideoCaptureDevice(videoDevices[currentCameraIndex].MonikerString);
            videoDevice.VideoResolution = videoDevice.VideoCapabilities[0];
            videoSourcePlayer1.VideoSource = videoDevice;
            videoDevice.Start();
            videoSourcePlayer1.Start();

            // 更新标签显示
            folder_path_label.Text = $"已切换摄像头: {videoDevices[currentCameraIndex].Name}";
        }

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

                                // 使用无矩形版本的方法（图片模式不需要矩形）
                                List<string> results = ReadQRCode(gray_bitmap);
                                if (results != null && results.Count > 0)
                                {
                                    for (int i = 0; i < results.Count; i++)
                                    {
                                        string[] details = new string[4]
                                        {
                                    Path.GetDirectoryName(item),
                                    Path.GetFileName(item),
                                    (i + 1).ToString(),
                                    results[i]
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
                videoDevices = new FilterInfoCollection(FilterCategory.VideoInputDevice);
                if (videoDevices.Count > 0)
                {
                    // 显示切换摄像头按钮
                    switchCamera_button.Visible = videoDevices.Count > 1;

                    // 初始化当前摄像头索引
                    currentCameraIndex = 0;

                    frame = null;
                    picture_radioButton.Enabled = false;
                    webcam_radioButton.Enabled = false;
                    quit_button.Enabled = false;
                    scan_button.Enabled = false;
                    cancel_button.Enabled = true;
                    folder_path_label.Text = "开始扫描二维码......";
                    videoDevice = new VideoCaptureDevice(videoDevices[0].MonikerString);
                    videoDevice.VideoResolution = videoDevice.VideoCapabilities[0];
                    videoSourcePlayer1.VideoSource = videoDevice;
                    videoDevice.Start();
                    videoSourcePlayer1.Start();
                    timer1.Interval = 100; // 加快扫描频率
                    timer1.Enabled = true;
                }
                else
                {
                    MessageBox.Show("未检测到摄像头设备！请使用图片识别方式扫描二维码");
                }
            }
        }

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
                    List<Rectangle> rects;
                    List<string> results = ReadQRCode(gray_bitmap, out rects);

                    // 更新矩形和帧尺寸
                    this.Invoke((MethodInvoker)delegate {
                        qrRectangles = rects;
                        if (frame != null)
                        {
                            frameWidth = frame.Width;
                            frameHeight = frame.Height;
                        }
                        videoSourcePlayer1.Invalidate(); // 请求重绘
                    });

                    if (results != null && results.Count > 0)
                    {
                        this.Invoke((MethodInvoker)delegate
                        {
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
                            lock (_lockObject)
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
                                qrRectangles.Clear(); // 清空矩形
                            });
                        }
                    }
                }
                isProcessing = false;
            });
        }


        //ZXing 二维码读取
        // 重载ReadQRCode方法 - 用于图片模式
        private List<string> ReadQRCode(Bitmap bitmap)
        {
            List<Rectangle> dummy;
            return ReadQRCode(bitmap, out dummy);
        }

        // 重载ReadQRCode方法 - 用于摄像头模式
        private List<string> ReadQRCode(Bitmap bitmap, out List<Rectangle> rectangles)
        {
            rectangles = new List<Rectangle>();
            List<string> qrCodeContents = new List<string>();

            BarcodeReader reader = new BarcodeReader()
            {
                Options = new DecodingOptions
                {
                    TryHarder = true,
                    PossibleFormats = new List<BarcodeFormat> { BarcodeFormat.QR_CODE },
                    CharacterSet = null
                }
            };

            try
            {
                var results = reader.DecodeMultiple(bitmap);
                if (results == null) return null;

                foreach (var result in results)
                {
                    if (result != null && !string.IsNullOrEmpty(result.Text))
                    {
                        qrCodeContents.Add(result.Text);

                        // 计算二维码边界框
                        if (result.ResultPoints != null && result.ResultPoints.Length >= 3)
                        {
                            float minX = float.MaxValue;
                            float minY = float.MaxValue;
                            float maxX = float.MinValue;
                            float maxY = float.MinValue;

                            foreach (var point in result.ResultPoints)
                            {
                                minX = Math.Min(minX, point.X);
                                minY = Math.Min(minY, point.Y);
                                maxX = Math.Max(maxX, point.X);
                                maxY = Math.Max(maxY, point.Y);
                            }

                            rectangles.Add(new Rectangle(
                                (int)minX, (int)minY,
                                (int)(maxX - minX), (int)(maxY - minY)));
                        }
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

        // 添加视频播放器的绘制事件
        private void videoSourcePlayer1_Paint(object sender, PaintEventArgs e)
        {
            if (qrRectangles == null || qrRectangles.Count == 0 || frameWidth == 0 || frameHeight == 0)
                return;

            float scaleX = (float)videoSourcePlayer1.Width / frameWidth;
            float scaleY = (float)videoSourcePlayer1.Height / frameHeight;

            foreach (var rect in qrRectangles)
            {
                // 计算缩放后的矩形
                Rectangle scaledRect = new Rectangle(
                    (int)(rect.X * scaleX),
                    (int)(rect.Y * scaleY),
                    (int)(rect.Width * scaleX),
                    (int)(rect.Height * scaleY)
                );

                // 绘制矩形框
                e.Graphics.DrawRectangle(qrBoxPen, scaledRect);

                // 添加聚焦效果
                using (var brush = new SolidBrush(Color.FromArgb(50, Color.LimeGreen)))
                {
                    e.Graphics.FillRectangle(brush, scaledRect);
                }
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
            // 停止设备
            if (videoSourcePlayer1.IsRunning)
            {
                videoSourcePlayer1.SignalToStop();
                videoSourcePlayer1.WaitForStop();
            }
            if (videoDevice != null && videoDevice.IsRunning)
            {
                videoDevice.SignalToStop();
                videoDevice.WaitForStop();
            }

            // 重置状态
            scan_button.Enabled = true;
            quit_button.Enabled = true;
            cancel_button.Enabled = false;
            isProcessing = false;
            qrRectangles.Clear();
            folder_path_label.Text = "";
        }


        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (videoSourcePlayer1.IsRunning)
            {
                videoSourcePlayer1.SignalToStop();
                videoSourcePlayer1.WaitForStop();
                videoSourcePlayer1.VideoSource = null;
            }
            if (videoDevice != null && videoDevice.IsRunning)
            {
                videoDevice.SignalToStop();
                videoDevice.WaitForStop();
            }

            // 释放资源
            if (qrBoxPen != null) qrBoxPen.Dispose();
        }
    }
}
