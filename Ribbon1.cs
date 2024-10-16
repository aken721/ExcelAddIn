using Microsoft.Office.Tools.Ribbon;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class Ribbon1
    {       

        //文件还是目录判断标识变量
        public static string runcommand = "";

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Select_f_or_d.Checked = false;
            Select_f_or_d.Label = "改文件名";
            Select_f_or_d.ShowLabel = false;
            switch_FD_label.Label = "文件名";
            ThisAddIn.Global.readFile = 0;
            confirm_spotlight.Checked = false;
            playbackMode = PlaybackMode.Sequential;
            Mode_button.Label = "顺序播放";
            Mode_button.Image = Properties.Resources.order_play;
            currentPlayState = PlaybackState.Stopped;
            Select_mp3_button.Image = Properties.Resources.no_open_fold;
            musicFiles.Clear();
            page_orientation_comboBox.Text = "横向";
            paper_size_comboBox.Text = "A4";
            page_zoom_comboBox.Text = "无缩放";
            sheet_export_comboBox.Text = "当前表";
            export_type_comboBox.Text = "多表单文件";
            export_type_comboBox.Visible=false;
        }


        //表操作按钮
        private void Excel_extend_Click(object sender, RibbonControlEventArgs e)
        {
            Form1 form1 = new Form1();
            form1.ShowDialog();
        }

        //邮件群发按钮
        private void Send_mail_Click(object sender, RibbonControlEventArgs e)
        {
            Form2 form2 = new Form2();
            form2.ShowDialog();
        }

        //MP3批量改名按钮
        private void Rename_mp3_Click(object sender, RibbonControlEventArgs e)
        {
            Form3 form3 = new Form3();
            form3.ShowDialog();
        }

        //文件删除或移动按钮
        private void delandmove_button_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;

            if (ThisAddIn.Global.readFile == 1 && IsSheetExist(workbook, "_rename"))
            {
                if (Select_f_or_d.Checked == false)
                {
                    runcommand = "file";
                }
                else
                {
                    runcommand = "folder";
                }
                Form form4 = new Form4();
                form4.FormClosed += new FormClosedEventHandler(form4_FormClosed);
                form4.Show();
            }
        }

        //窗体4关闭事件
        private async void form4_FormClosed(object sender, FormClosedEventArgs e)
        {
            await Task.Run(() =>
            {
                int runClick_form4 = Form4.runButtonClicked;
                int resetClick_form4 = Form4.resetButtonClicked;
                if (runClick_form4 > 0 || resetClick_form4 > 0)
                {
                    RefreshRenameTable();
                    MessageBox.Show("删除或移动文件窗口已关闭，_rename表已更新，如不再需要进行重命名操作，可手工删除_rename表即可");
                }
                else
                {
                    return;
                }
            }); 
        }

        //指定字段名所处的列
        private int GetUsedRangeColumn(string targetColumn)
        {
            for (int n = 1; n <= ThisAddIn.app.ActiveSheet.UsedRange.Columns.Count; n++)
            {
                string targetValue = ThisAddIn.app.ActiveSheet.Cells[1, n].Value.ToString();
                if (targetValue == targetColumn) return n;
            }
            return 0;
        }


        //批读文件名和批改文件名选择路径
        string get_directory_path;

        //批读文件名
        private void Files_read_Click(object sender, RibbonControlEventArgs e)
        {
            if (ThisAddIn.Global.readFile == 1 && get_directory_path.Length > 0)
            {
                ThisAddIn.app.Application.StatusBar = "正在刷新_rename表";
                RefreshRenameTable();
                ThisAddIn.app.Application.StatusBar = false; 
            }
            else
            {
                Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;
                ThisAddIn.app.DisplayAlerts = false;
                ThisAddIn.app.ScreenUpdating = false;

                folderBrowserDialog1.Description = "请选择文件所在文件夹";
                if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                {
                    get_directory_path = folderBrowserDialog1.SelectedPath;
                }
                else
                {

                    ThisAddIn.app.DisplayAlerts = true;
                    ThisAddIn.app.ScreenUpdating = true;
                    return;
                }

                if (!string.IsNullOrEmpty(get_directory_path))
                {
                ThisAddIn.app.Application.StatusBar = "_rename表正在读取生成中，请稍等......";

                    foreach (Excel.Worksheet sheet in workbook.Worksheets)
                    {
                        if (sheet.Name == "_rename")
                        {
                            sheet.Name = "_rename_备份";
                        }
                    }
                    Excel.Worksheet worksheet = workbook.Worksheets.Add();
                    worksheet.Name = "_rename";
                    //worksheet.Activate();
                    try
                    {
                        switch (Select_f_or_d.Checked)
                        {
                            case false:
                                worksheet.Cells[1, 1] = "路径";
                                worksheet.Cells[1, 2] = "旧文件名";
                                worksheet.Cells[1, 3] = "新文件名";
                                List<string> files = new List<string>(Directory.GetFiles(get_directory_path, "*.*", SearchOption.AllDirectories));
                                files.RemoveAll(file => (File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden);
                                if (files.Count > 0)
                                {
                                    for (int i = 1; i <= files.Count; i++)
                                    {
                                        string file_name = Path.GetFileName(files[i - 1]);
                                        string file_path = Path.GetDirectoryName(files[i - 1]);
                                        worksheet.Cells[i + 1, 1] = file_path;
                                        worksheet.Hyperlinks.Add(worksheet.Cells[i + 1, 1], file_path, Type.Missing, Type.Missing, file_path);
                                        worksheet.Cells[i + 1, 2] = file_name;
                                        worksheet.Hyperlinks.Add(worksheet.Cells[i + 1, 2], file_path + "\\" + file_name, Type.Missing, Type.Missing, file_name);
                                        worksheet.Cells[i + 1, 3] = file_name;
                                    }
                                }
                                worksheet.Range["C2"].Select();
                                break;
                            case true:
                                worksheet.Cells[1, 1] = "文件夹路径";
                                worksheet.Cells[1, 2] = "旧文件夹名";
                                worksheet.Cells[1, 3] = "新文件夹名";
                                string[] directorys = Directory.GetDirectories(get_directory_path, "*", SearchOption.AllDirectories);
                                if (directorys.Length > 0)
                                {
                                    for (int i = 1; i <= directorys.Length; i++)
                                    {
                                        string[] directory = directorys[i - 1].Split('\\');
                                        string directory_name = directory[directory.Length - 1];
                                        Array.Resize(ref directory, directory.Length - 1);
                                        string directory_path = string.Join("\\", directory);
                                        worksheet.Cells[i + 1, 1] = directory_path;
                                        worksheet.Hyperlinks.Add(worksheet.Cells[i + 1, 1], directory_path, Type.Missing, Type.Missing, directory_path);
                                        worksheet.Cells[i + 1, 2] = directory_name;
                                        worksheet.Hyperlinks.Add(worksheet.Cells[i + 1, 2], directory_path + "\\" + directory_name, Type.Missing, Type.Missing, directory_name);
                                        worksheet.Cells[i + 1, 3] = directory_name;
                                    }
                                    worksheet.Range["C2"].Select();
                                }
                                break;
                        }
                        ThisAddIn.Global.readFile = 1;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        ThisAddIn.app.DisplayAlerts = true;
                        ThisAddIn.app.ScreenUpdating = true;
                        ThisAddIn.app.Application.StatusBar = false;
                    }
                }
                else
                {
                    MessageBox.Show("未选择文件夹");
                    ThisAddIn.app.DisplayAlerts = true;
                    ThisAddIn.app.ScreenUpdating = true;
                }              

                workbook.Worksheets["_rename"].Activate();
                workbook.RefreshAll();
                ThisAddIn.app.Application.StatusBar = false;
            }           
        }

        //批量重命名
        private void File_rename_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;
            Excel.Worksheet worksheet = workbook.Worksheets["_rename"];

            if (!string.IsNullOrEmpty(get_directory_path) && ThisAddIn.Global.readFile == 1 && IsSheetExist(workbook, "_rename"))
            {
                //调用file.move或direction.move修改名
                ThisAddIn.app.Application.StatusBar="文件正在重命名中，请稍后...";

                for (int i = 2; i <= workbook.ActiveSheet.UsedRange.Rows.Count; i++)
                {
                    string cell1 = worksheet.Cells[i, 1].Value;
                    string cell2 = worksheet.Cells[i, 2].Value;
                    string cell3 = worksheet.Cells[i, 3].Value;
                    if (string.IsNullOrEmpty(cell1) || string.IsNullOrEmpty(cell2) || string.IsNullOrEmpty(cell3))
                    {
                        MessageBox.Show($"第{i}行有空格，请检查！该行对应文件将不被改名");
                        continue;
                    }
                    else
                    {
                        string old_name = Path.Combine(cell1, cell2);
                        string new_name = Path.Combine(cell1, cell3);
                        switch (Select_f_or_d.Checked == true)
                        {
                            case false:
                                int exist_file = 0;
                                if (old_name != new_name)
                                {
                                    while (File.Exists(new_name))
                                    {
                                        exist_file++;
                                        new_name = Path.Combine(cell1, Path.GetFileNameWithoutExtension(new_name) + "(" + exist_file.ToString() + ")" + Path.GetExtension(new_name));
                                    }
                                    File.Move(old_name, new_name);
                                }
                                break;
                            case true:
                                int exist_fold = 0;
                                if (old_name != new_name)
                                {
                                    while (Directory.Exists(new_name))
                                    {
                                        exist_fold++;
                                        new_name = Path.Combine(cell1, cell3 + "(" + exist_fold.ToString() + ")");
                                    }
                                    Directory.Move(old_name, new_name);
                                }
                                break;
                        }
                    }
                }


                //删除_rename表，并显示完成结果
                workbook.Worksheets["_rename"].Delete();
                if (IsSheetExist(workbook, "_rename_备份"))
                {
                    workbook.Worksheets["_rename_备份"].Name = "_rename";
                }
                workbook.RefreshAll();
                MessageBox.Show("文件名修改完毕");
                if (Directory.Exists(get_directory_path))
                {
                    Process.Start(get_directory_path);
                }
                ThisAddIn.Global.readFile = 0;
            }
            else
            {
                MessageBox.Show("没有选择文件夹，请先使用批读文件名功能后再使用该功能");
                ThisAddIn.Global.readFile = 0;
            }
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
            ThisAddIn.app.Application.StatusBar=false;
        }
  

        //文件目录选项
        private void Select_f_or_d_Click(object sender, RibbonControlEventArgs e)
        {
            if (Select_f_or_d.Checked == true)
            {
                Select_f_or_d.Image = ExcelAddIn.Properties.Resources.Radio_Button_on;
                Select_f_or_d.Label = "改文件夹名";
                Select_f_or_d.ShowLabel = false;
                switch_FD_label.Label = "目录名";
            }
            else
            {
                Select_f_or_d.Image = ExcelAddIn.Properties.Resources.Radio_Button_off;
                Select_f_or_d.Label = "改文件名";
                Select_f_or_d.ShowLabel = false;
                switch_FD_label.Label = "文件名";
            }
        }


        //判断指定工作簿中指定工作表名是否存在
        public static bool IsSheetExist(Excel.Workbook workbook, string sheetName)
        {
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                if (worksheet.Name == sheetName)
                {
                    return true;
                }
            }
            return false;
        }

        //音乐播放模式
        private enum PlaybackMode
        {
            Sequential,
            SingleLoop,
            AllLoop
        }

        //音乐播放状态
        private enum PlaybackState
        {
            Stopped,
            Playing,
            Paused
        }

        //初始化音乐播放模式
        private PlaybackMode playbackMode = PlaybackMode.Sequential;
        //初始化音乐播放状态
        private PlaybackState currentPlayState = PlaybackState.Stopped;

        //实例化WaveOutEvent对象
        private WaveOutEvent waveOutEvent;
        //实例化AudioFileReader对象
        private AudioFileReader audioFile = null;

        //音乐播放列表
        private readonly List<string> musicFiles = new List<string>();
        //当前播放歌曲序号
        private int currentSongIndex = -1;

        //选择音乐文件夹
        private void Select_mp3_button_Click(object sender, RibbonControlEventArgs e)
        {
            musicFiles.Clear();
            DisposeWavePlayer();
            folderBrowserDialog1.Description = "请选择音乐文件所在文件夹";
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                LoadMusicFiles(folderBrowserDialog1.SelectedPath);
                if (musicFiles.Count == 0)
                {
                    Select_mp3_button.Image = Properties.Resources.no_open_fold;
                    ThisAddIn.app.Application.StatusBar = "选择的文件夹中没有支持格式的音乐";
                }
                else
                {
                    currentSongIndex = 0;
                    Select_mp3_button.Image = Properties.Resources.fold;
                    ThisAddIn.app.Application.StatusBar = $"当前未播放音乐，共{musicFiles.Count}首音乐，第1首：{Path.GetFileName(musicFiles[currentSongIndex])}";
                }
            }
            else
            {
                return;
            }
        }

        //写入播放列表方法
        private void LoadMusicFiles(string folderPath)
        {
            musicFiles.AddRange(Directory.GetFiles(folderPath, "*.mp3", SearchOption.AllDirectories));
            musicFiles.AddRange(Directory.GetFiles(folderPath, "*.wav", SearchOption.AllDirectories));
            musicFiles.AddRange(Directory.GetFiles(folderPath, "*.flac", SearchOption.AllDirectories));
            musicFiles.AddRange(Directory.GetFiles(folderPath, "*.aiff", SearchOption.AllDirectories));
            musicFiles.AddRange(Directory.GetFiles(folderPath, "*.wma", SearchOption.AllDirectories));
            musicFiles.AddRange(Directory.GetFiles(folderPath, "*.aac", SearchOption.AllDirectories));
            musicFiles.AddRange(Directory.GetFiles(folderPath, "*.g711", SearchOption.AllDirectories));
            musicFiles.AddRange(Directory.GetFiles(folderPath, "*.mp4", SearchOption.AllDirectories));
        }

        //播放模式选择按钮
        private void Mode_button_Click(object sender, RibbonControlEventArgs e)
        {
            switch (playbackMode)
            {
                case PlaybackMode.Sequential:
                    playbackMode = PlaybackMode.AllLoop;
                    Mode_button.Label = "全部循环";
                    Mode_button.ScreenTip = "全部循环";
                    Mode_button.Image = Properties.Resources.all_cycle;
                    break;
                case PlaybackMode.AllLoop:
                    playbackMode = PlaybackMode.SingleLoop;
                    Mode_button.Label = "单曲循环";
                    Mode_button.ScreenTip = "单曲循环";
                    Mode_button.Image = Properties.Resources.single_cycle;
                    break;
                case PlaybackMode.SingleLoop:
                    playbackMode = PlaybackMode.Sequential;
                    Mode_button.Label = "顺序播放";
                    Mode_button.ScreenTip = "顺序播放";
                    Mode_button.Image = Properties.Resources.order_play;
                    break;
            }
        }

        private bool isMusicEnded = false;

        //播放按钮单击事件
        private async void Play_button_Click(object sender, RibbonControlEventArgs e)
        {
            //timer1.Interval = 200;


            if (musicFiles.Count == 0)
                return;

            isAutoContinue = true;

            if (currentPlayState == PlaybackState.Stopped)
            {
                //timer1.Enabled = true;
                await PlayMusic();
            }
            else if (currentPlayState == PlaybackState.Paused)
            {
                waveOutEvent?.Play();
                //timer1.Enabled = true;
                currentPlayState = PlaybackState.Playing;
                UpdateTrackInfo();
                Play_button.Image = Properties.Resources.pause;
                Play_button.ScreenTip = "暂停";
            }
            else
            {
                waveOutEvent?.Pause();
                //timer1.Enabled = false;
                currentPlayState = PlaybackState.Paused;
                ThisAddIn.app.Application.StatusBar = "播放暂停中......";
                Play_button.Image = Properties.Resources.play;
                Play_button.ScreenTip = "播放";
            }
        }

        //播放/暂停（方法）
        private async Task PlayMusic()
        {

            isMusicEnded = false;
            if (currentPlayState == PlaybackState.Stopped)
            {
                currentPlayState = PlaybackState.Playing;
                Play_button.Image = Properties.Resources.pause;
                Play_button.ScreenTip = "暂停";
                DisposeWavePlayer();
                if (waveOutEvent == null)
                {
                    waveOutEvent = new WaveOutEvent();
                    waveOutEvent.PlaybackStopped += OnPlaybackStopped;
                }
                if (audioFile == null)
                {
                    audioFile = new AudioFileReader(musicFiles[currentSongIndex]);
                    waveOutEvent?.Init(audioFile);
                }
                UpdateTrackInfo();
                waveOutEvent?.Play();

                await Task.Delay(200);

                // 等待播放完成
                while (!isMusicEnded)
                {
                    await Task.Delay(100);
                }
            }
        }


        //播放完毕事件触发
        private async void OnPlaybackStopped(object sender, StoppedEventArgs args)
        {
            DisposeWavePlayer();
            isMusicEnded = true;
            currentPlayState = PlaybackState.Stopped;

            if (isAutoContinue == false)
            {
                ThisAddIn.app.Application.StatusBar = false;
                currentSongIndex = 0;
                return;
            }
            else
            {
                // 播放下一首歌曲
                currentSongIndex++;
                switch (playbackMode)
                {

                    case PlaybackMode.Sequential:
                        if (currentSongIndex < musicFiles.Count)
                        {
                            await Task.Delay(1000);
                            await PlayMusic();
                        }
                        else
                        {
                            currentSongIndex = 0;
                            StopMusic();

                        }
                        break;

                    case PlaybackMode.AllLoop:
                        if (currentSongIndex == musicFiles.Count)
                        {
                            currentSongIndex = 0;

                        }
                        await Task.Delay(1000);
                        await PlayMusic();
                        break;
                    case PlaybackMode.SingleLoop:
                        currentSongIndex--;
                        await Task.Delay(1000);
                        await PlayMusic();
                        break;
                }
            }
        }

        private bool isAutoContinue = true;

        //停止按钮单击事件
        private void Stop_button_Click(object sender, RibbonControlEventArgs e)
        {
            StopMusic();
        }

        //停止播放
        private void StopMusic()
        {
            if (waveOutEvent != null)
            {
                waveOutEvent?.Stop();
                waveOutEvent.PlaybackStopped += OnPlaybackStopped;

            }
            if (musicFiles.Count > 0)
            {
                ThisAddIn.app.Application.StatusBar = $"播放已停止，共{musicFiles.Count}首音乐，可从第1首重新播放";
            }
            else
            {
                ThisAddIn.app.Application.StatusBar = false;
            }
            isAutoContinue = false;
            //timer1.Enabled = false;
            Play_button.Image = Properties.Resources.play;
            Play_button.Label = "播放";
        }

        //下一首曲目
        private async void Next_button_Click(object sender, RibbonControlEventArgs e)
        {
            if (musicFiles.Count != 0)
            {
                if (currentPlayState == PlaybackState.Stopped)
                {
                    currentSongIndex = (currentSongIndex + 1) % musicFiles.Count;
                    ThisAddIn.app.Application.StatusBar = $"当前未播放音乐，共{musicFiles.Count}首音乐，第{currentSongIndex + 1}首：{Path.GetFileName(musicFiles[currentSongIndex])}";
                }
                else
                {
                    if (playbackMode == PlaybackMode.SingleLoop)
                    {
                        waveOutEvent?.Stop();
                        currentSongIndex = (currentSongIndex + 1) % musicFiles.Count;
                        await PlayMusic();
                    }
                    else
                    {
                        waveOutEvent?.Stop();
                        currentSongIndex = (currentSongIndex) % musicFiles.Count;
                        await PlayMusic();
                    }
                }
            }
        }

        //上一首曲目
        private async void Previous_button_Click(object sender, RibbonControlEventArgs e)
        {
            if (musicFiles.Count != 0)
            {
                if (currentPlayState == PlaybackState.Stopped)
                {
                    currentSongIndex = (currentSongIndex - 1 + musicFiles.Count) % musicFiles.Count;
                    ThisAddIn.app.Application.StatusBar = $"当前未播放音乐，共{musicFiles.Count}首音乐，第{currentSongIndex + 1}首：{Path.GetFileName(musicFiles[currentSongIndex])}";
                }
                else
                {
                    if (playbackMode==PlaybackMode.SingleLoop)
                    {
                        waveOutEvent?.Stop();
                        currentSongIndex = (currentSongIndex - 1 + musicFiles.Count) % musicFiles.Count;
                        await PlayMusic();
                    }
                    else
                    {
                        waveOutEvent?.Stop();
                        currentSongIndex = (currentSongIndex - 2 + musicFiles.Count) % musicFiles.Count;
                        await PlayMusic();
                    }                    
                }
            }
        }

        //清理wavePlayer
        private void DisposeWavePlayer()
        {
            waveOutEvent?.Dispose();
            waveOutEvent = null;
            audioFile?.Dispose();
            audioFile = null;
        }

        //当前播放歌曲信息
        private async void UpdateTrackInfo()
        {
            if (audioFile != null && currentSongIndex < musicFiles.Count)
            {
                string trackInfo = await Task.Run(() =>
                {
                    return $"正在播放第{currentSongIndex + 1}首：{Path.GetFileName(musicFiles[currentSongIndex])}，" +
                        $"时长：{audioFile.TotalTime:mm\\:ss}";
                    //return $"正在播放第{currentSongIndex + 1}首：{Path.GetFileName(musicFiles[currentSongIndex])}，" +
                    //    $"已播放时长：{audioFile.CurrentTime.ToString(@"mm\:ss")}，" + $"歌曲时长：{audioFile.TotalTime.ToString(@"mm\:ss")}";
                });
                try
                {
                    ThisAddIn.app.Application.StatusBar = trackInfo;
                }
                catch
                {
                    if (currentPlayState == PlaybackState.Playing)
                    {
                        waveOutEvent?.Pause();
                        //timer1.Enabled = false;
                        currentPlayState = PlaybackState.Paused;
                        //ThisAddIn.app.Application.StatusBar = "播放暂停中......";
                        Play_button.Image = Properties.Resources.play;
                        Play_button.ScreenTip = "播放";
                    }
                }
            }
            else
            {
                ThisAddIn.app.Application.StatusBar = false;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            UpdateTrackInfo();
        }

        //记录聚光灯功能打开前的彩色单元格位置和填充颜色
        Dictionary<string, int> cellColor = new Dictionary<string, int>();  

        //聚光灯功能按钮
        private void confirm_spotlight_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet currentWorksheet = ThisAddIn.app.ActiveSheet;
            Excel.Range usedRange = currentWorksheet.UsedRange;
            if (confirm_spotlight.Checked==true)
            {
                cellColor = GetColorDictionary(usedRange);
                confirm_spotlight.Label = "关闭聚光灯";
                confirm_spotlight.Image = ExcelAddIn.Properties.Resources.spotlight_open;
                ThisAddIn.Global.spotlight = 1;
            }
            else
            {
                confirm_spotlight.Label = "打开聚光灯";
                confirm_spotlight.Image = ExcelAddIn.Properties.Resources.spotlight_close;
                ThisAddIn.Global.spotlight = 0;
                Excel.Range selectCell= ThisAddIn.app.ActiveCell;
                ThisAddIn.app.ScreenUpdating = false;
                selectCell.EntireRow.Interior.ColorIndex = 0;
                selectCell.EntireColumn.Interior.ColorIndex = 0;
                if (cellColor.Count>0)
                {
                    foreach (var cellColorEntry in cellColor)
                    {
                        string cellAddress = cellColorEntry.Key;
                        int cellColorIndex=cellColorEntry.Value;
                        Excel.Range cell= currentWorksheet.Range[cellAddress];
                        cell.Interior.ColorIndex = cellColorIndex;
                    }
                }
                cellColor.Clear();
                ThisAddIn.app.ScreenUpdating = true;
            }
        }

        //获取已有彩色单元格位置和颜色索引的字典变量的方法
        private Dictionary<string,int> GetColorDictionary(Excel.Range usedRange)
        {
            Dictionary<string,int> cellColorDict=new Dictionary<string,int>();
            foreach (Excel.Range cell in usedRange)
            {
                if (cell.Interior.ColorIndex>0)
                {
                    string cellAddress = cell.Address;
                    int cellColorIndex = cell.Interior.ColorIndex;
                    cellColorDict.Add(cellAddress,cellColorIndex);
                }
            }
            return cellColorDict;
        }

        //刷新_rename表
        private void RefreshRenameTable()
        {
            Excel.Worksheet worksheet = ThisAddIn.app.ActiveWorkbook.Worksheets["_rename"];
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;
            worksheet.Rows.Clear();

            try
            {
                switch (Select_f_or_d.Checked)
                {
                    case false:
                        worksheet.Cells[1, 1] = "路径";
                        worksheet.Cells[1, 2] = "旧文件名";
                        worksheet.Cells[1, 3] = "新文件名";
                        List<string> files = new List<string>(Directory.GetFiles(get_directory_path, "*.*", SearchOption.AllDirectories));
                        files.RemoveAll(file => (File.GetAttributes(file) & FileAttributes.Hidden) == FileAttributes.Hidden);
                        if (files.Count > 0)
                        {
                            for (int i = 1; i <= files.Count; i++)
                            {
                                string file_name = Path.GetFileName(files[i - 1]);
                                string file_path = Path.GetDirectoryName(files[i - 1]);
                                worksheet.Cells[i + 1, 1] = file_path;
                                worksheet.Hyperlinks.Add(worksheet.Cells[i + 1, 1], file_path, Type.Missing, Type.Missing, file_path);
                                worksheet.Cells[i + 1, 2] = file_name;
                                worksheet.Hyperlinks.Add(worksheet.Cells[i + 1, 2], file_path + "\\" + file_name, Type.Missing, Type.Missing, file_name);
                                worksheet.Cells[i + 1, 3] = file_name;
                            }
                        }
                        worksheet.Range["C2"].Select();                        
                        break;
                    case true:
                        worksheet.Cells[1, 1] = "文件夹路径";
                        worksheet.Cells[1, 2] = "旧文件夹名";
                        worksheet.Cells[1, 3] = "新文件夹名";
                        string[] directorys = Directory.GetDirectories(get_directory_path, "*", SearchOption.AllDirectories);
                        if (directorys.Length > 0)
                        {
                            for (int i = 1; i <= directorys.Length; i++)
                            {
                                string[] directory = directorys[i - 1].Split('\\');
                                string directory_name = directory[directory.Length - 1];
                                Array.Resize(ref directory, directory.Length - 1);
                                string directory_path = string.Join("\\", directory);
                                worksheet.Cells[i + 1, 1] = directory_path;
                                worksheet.Hyperlinks.Add(worksheet.Cells[i + 1, 1], directory_path, Type.Missing, Type.Missing, directory_path);
                                worksheet.Cells[i + 1, 2] = directory_name;
                                worksheet.Hyperlinks.Add(worksheet.Cells[i + 1, 2], directory_path + "\\" + directory_name, Type.Missing, Type.Missing, directory_name);
                                worksheet.Cells[i + 1, 3] = directory_name;
                            }
                            worksheet.Range["C2"].Select();
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating=true;
                ThisAddIn.app.ActiveWorkbook.RefreshAll();
                
            }
        }

        private void to_pdf_button_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;

            switch (sheet_export_comboBox.Text)
            {
                case "当前表":
                    ThisAddIn.app.ScreenUpdating = false;
                    ThisAddIn.app.DisplayAlerts= false;
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    string sheetName = worksheet.Name;
                    saveFileDialog1.Filter = "PDF Files|*.pdf";
                    saveFileDialog1.Title = "保存PDF文件";
                    saveFileDialog1.FileName = $"{sheetName}.pdf";
                    saveFileDialog1.AddExtension = true;
                    DialogResult result = saveFileDialog1.ShowDialog();
                    Excel.PageSetup pageSetup = worksheet.PageSetup;

                    if (result == DialogResult.OK)
                    {
                        // 设置页面
                        pageSetup.PrintArea = worksheet.UsedRange.Address; //设置打印区域
                        pageSetup.LeftMargin = 0.8; // 左边距，单位为英寸
                        pageSetup.RightMargin = 0.8; // 右边距
                        pageSetup.TopMargin = 0.8; // 上边距
                        pageSetup.BottomMargin = 0.8; // 下边距
                        pageSetup.CenterHorizontally = true;
                        pageSetup.CenterVertically = true;

                        switch (page_orientation_comboBox.Text)
                        {
                            case "纵向":
                                worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                                break;
                            case "横向":
                                worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                                break;
                            default:
                                worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                                break;
                        }
                        switch (paper_size_comboBox.Text)
                        {
                            case "A4":
                                worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                                break;
                            case "A3":
                                worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA3;
                                break;
                            case "B5":
                                worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperB5;
                                break;
                            default:
                                worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                                break;
                        }
                        switch (page_zoom_comboBox.Text)
                        {
                            case "无缩放":
                                pageSetup.Zoom = 100;
                                break;
                            case "表自适应":
                                pageSetup.Zoom = false;
                                pageSetup.FitToPagesWide = 1;
                                pageSetup.FitToPagesTall = 1;
                                break;
                            case "行自适应":
                                pageSetup.Zoom = false;
                                pageSetup.FitToPagesTall = 1;
                                break;
                            case "列自适应":
                                pageSetup.Zoom = false;
                                pageSetup.FitToPagesTall = 1;
                                break;
                        }
                        worksheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, saveFileDialog1.FileName, Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false);
                        MessageBox.Show("PDF已转换完成！");
                    }
                    ThisAddIn.app.ScreenUpdating = true;
                    ThisAddIn.app.DisplayAlerts = true;
                    break;
                case "全部表":
                    ThisAddIn.app.ScreenUpdating = false;
                    ThisAddIn.app.DisplayAlerts = false;
                    switch (export_type_comboBox.Text)
                    {
                        case "多表单文件":
                            saveFileDialog1.Filter = "PDF Files|*.pdf";
                            saveFileDialog1.Title = "保存PDF文件";
                            saveFileDialog1.FileName = $"{Path.GetFileNameWithoutExtension(workbook.Name)}.pdf";
                            saveFileDialog1.AddExtension = true;
                            DialogResult res = saveFileDialog1.ShowDialog();
                            if (res == DialogResult.OK) 
                            {
                                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                                {
                                    // 设置页面
                                    Excel.PageSetup pagesSetup = sheet.PageSetup;
                                    pagesSetup.PrintArea = sheet.UsedRange.Address; //设置打印区域
                                    pagesSetup.LeftMargin = 0.8; // 左边距，单位为英寸
                                    pagesSetup.RightMargin = 0.8; // 右边距
                                    pagesSetup.TopMargin = 0.8; // 上边距
                                    pagesSetup.BottomMargin = 0.8; // 下边距
                                    pagesSetup.CenterHorizontally = true;
                                    pagesSetup.CenterVertically = true;

                                    switch (page_orientation_comboBox.Text)
                                    {
                                        case "纵向":
                                            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                                            break;
                                        case "横向":
                                            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                                            break;
                                        default:
                                            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                                            break;
                                    }
                                    switch (paper_size_comboBox.Text)
                                    {
                                        case "A4":
                                            sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                                            break;
                                        case "A3":
                                            sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA3;
                                            break;
                                        case "B5":
                                            sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperB5;
                                            break;
                                        default:
                                            sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                                            break;
                                    }
                                    switch (page_zoom_comboBox.Text)
                                    {
                                        case "无缩放":
                                            pagesSetup.Zoom = 100;
                                            break;
                                        case "表自适应":
                                            pagesSetup.Zoom = false;
                                            pagesSetup.FitToPagesWide = 1;
                                            pagesSetup.FitToPagesTall = 1;
                                            break;
                                        case "行自适应":
                                            pagesSetup.Zoom = false;
                                            pagesSetup.FitToPagesTall = 1;
                                            break;
                                        case "列自适应":
                                            pagesSetup.Zoom = false;
                                            pagesSetup.FitToPagesTall = 1;
                                            break;
                                    }
                                }                                
                            }
                            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, saveFileDialog1.FileName, Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false);
                            MessageBox.Show("PDF已转换完成！");
                            break;
                        case "多表多文件":
                            ThisAddIn.app.ScreenUpdating = false;
                            ThisAddIn.app.DisplayAlerts = false;
                            string get_pdf_path;
                            string output_pdf_name;
                            string output_pdf_path;
                            folderBrowserDialog1.Description = "请选择保存PDF的文件夹";
                            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                            {
                                get_pdf_path = folderBrowserDialog1.SelectedPath;
                            }
                            else
                            {
                                return;
                            }
                            foreach (Excel.Worksheet sheet in workbook.Worksheets)
                            {
                                // 设置页面
                                Excel.PageSetup pagesSetup = sheet.PageSetup;
                                pagesSetup.PrintArea = sheet.UsedRange.Address; //设置打印区域
                                pagesSetup.LeftMargin = 0.8; // 左边距，单位为英寸
                                pagesSetup.RightMargin = 0.8; // 右边距
                                pagesSetup.TopMargin = 0.8; // 上边距
                                pagesSetup.BottomMargin = 0.8; // 下边距
                                pagesSetup.CenterHorizontally = true;
                                pagesSetup.CenterVertically = true;

                                switch (page_orientation_comboBox.Text)
                                {
                                    case "纵向":
                                        sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                                        break;
                                    case "横向":
                                        sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                                        break;
                                    default:
                                        sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
                                        break;
                                }
                                switch (paper_size_comboBox.Text)
                                {
                                    case "A4":
                                        sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                                        break;
                                    case "A3":
                                        sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA3;
                                        break;
                                    case "B5":
                                        sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperB5;
                                        break;
                                    default:
                                        sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                                        break;
                                }
                                switch (page_zoom_comboBox.Text)
                                {
                                    case "无缩放":
                                        pagesSetup.Zoom = 100;
                                        break;
                                    case "表自适应":
                                        pagesSetup.Zoom = false;
                                        pagesSetup.FitToPagesWide = 1;
                                        pagesSetup.FitToPagesTall = 1;
                                        break;
                                    case "行自适应":
                                        pagesSetup.Zoom = false;
                                        pagesSetup.FitToPagesTall = 1;
                                        break;
                                    case "列自适应":
                                        pagesSetup.Zoom = false;
                                        pagesSetup.FitToPagesTall = 1;
                                        break;
                                }
                                output_pdf_name = sheet.Name;
                                output_pdf_path = Path.Combine(get_pdf_path, output_pdf_name + ".pdf");
                                // 如果目标路径已经存在同名文件，则重命名目标文件
                                int i = 1;
                                string originalFileName = Path.GetFileNameWithoutExtension(output_pdf_path);
                                while (File.Exists(output_pdf_path))
                                {
                                    string newFileName = $"{originalFileName}({i}){Path.GetExtension(output_pdf_path)}";
                                    output_pdf_path = Path.Combine(get_pdf_path, newFileName);
                                    i++;
                                }
                                sheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, output_pdf_path, Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false);
                            }
                            MessageBox.Show("PDF已转换完成！");
                            break;
                    }
                    ThisAddIn.app.ScreenUpdating = true;
                    ThisAddIn.app.DisplayAlerts = true;
                    break;
            }            
        }

        /// <summary>
        /// 以下代码均为RibbonComboBox中的限制手工输入逻辑。
        /// 因RibbonComboBox不支持readonly属性，所以在判断手工输入内容为非可选内容时，自动会改为默认可选内容
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void sheet_export_comboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (sheet_export_comboBox.Text == "全部表")
            {
                export_type_comboBox.Visible = true;
            }
            else if(sheet_export_comboBox.Text =="当前表")
            {
                export_type_comboBox.Visible = false;
            }
            else
            {       

                if (MessageBox.Show("该选项只能是“当前表”或“全部表”")==DialogResult.OK)
                {
                    export_type_comboBox.Visible = false;
                    sheet_export_comboBox.Text = "当前表";
                }
            }
        }

        private void export_type_comboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if(export_type_comboBox.Text =="多表多文件" || export_type_comboBox.Text == "多表单文件")
            {
                return;
            }
            else
            {
                if (MessageBox.Show("该选项只能是“多表单文件”或“多表多文件”") == DialogResult.OK)
                {
                    export_type_comboBox.Text = "多表单文件";
                }
            }
        }

        private void page_orientation_comboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (page_orientation_comboBox.Text == "纵向" || page_orientation_comboBox.Text == "横向")
            {
                return;
            }
            else
            {
                if (MessageBox.Show("该选项只能是“纵向”或“横向”") == DialogResult.OK)
                {
                    page_orientation_comboBox.Text = "横向";
                }
            }
        }

        private void paper_size_comboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (paper_size_comboBox.Text == "A4" || paper_size_comboBox.Text == "A3" || paper_size_comboBox.Text == "A5" || paper_size_comboBox.Text == "B5")
            {
                return;
            }
            else
            {
                if (MessageBox.Show("该选项只能下拉选择") == DialogResult.OK)
                {
                    paper_size_comboBox.Text = "A4";
                }
            }
        }

        private void page_zoom_comboBox_TextChanged(object sender, RibbonControlEventArgs e)
        {
            if (page_zoom_comboBox.Text == "无缩放" || page_zoom_comboBox.Text == "表自适应" || page_zoom_comboBox.Text == "行自适应" || page_zoom_comboBox.Text == "列自适应")
            {
                return;
            }
            else
            {
                if (MessageBox.Show("该选项只能下拉选择") == DialogResult.OK)
                {
                    page_zoom_comboBox.Text = "无缩放";
                }
            }
        }

        private void scan_button_Click(object sender, RibbonControlEventArgs e)
        {
            Form form5 = new Form5();
            form5.Show();
        }

        private void fapiao_button_Click(object sender, RibbonControlEventArgs e)
        {
            Form form6 = new Form6();
            form6.Show();
        }
    }
}

