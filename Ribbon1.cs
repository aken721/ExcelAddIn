using Microsoft.Office.Tools.Ribbon;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn
{
    public partial class Ribbon1
    {

        //是否执行读文件名功能标识
        private int readFile;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            Select_f_or_d.Checked = false;
            Select_f_or_d.Label = "改文件名";
            Select_f_or_d.ShowLabel = false;
            switch_FD_label.Label = "文件名";
            readFile = 0;
            playbackMode = PlaybackMode.Sequential;
            Mode_btuuon.Label = "顺序播放";
            Mode_btuuon.Image = Properties.Resources.order_play;
            currentPlayState = PlaybackState.Stopped;
            Select_mp3_button.Image = Properties.Resources.no_open_fold;
            musicFiles.Clear();
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
                return;
            }

            if (!string.IsNullOrEmpty(get_directory_path))
            {
                string bat_name = get_directory_path + "\\run.bat";
                FileInfo run_file = new FileInfo(bat_name);
                if (run_file.Exists)
                {
                    run_file.Delete();
                }
                foreach (Excel.Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == "_rename")
                    {
                        sheet.Name = "_rename_备份";
                    }
                }
                Excel.Worksheet worksheet = workbook.Worksheets.Add();
                worksheet.Name = "_rename";
                worksheet.Activate();
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
                                workbook.ActiveSheet.Cells[i + 1, 1] = file_path;
                                workbook.ActiveSheet.Cells[i + 1, 2] = file_name;
                                workbook.ActiveSheet.Cells[i + 1, 3] = file_name;
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
                                workbook.ActiveSheet.Cells[i + 1, 1] = directory_path;
                                workbook.ActiveSheet.Cells[i + 1, 2] = directory_name;
                                workbook.ActiveSheet.Cells[i + 1, 3] = directory_name;
                            }
                            worksheet.Range["C2"].Select();
                        }
                        break;
                }
            }
            else
            {
                MessageBox.Show("未选择文件夹");
            }
            readFile = 1;
            ThisAddIn.app.DisplayAlerts = true;
            ThisAddIn.app.ScreenUpdating = true;
        }

        //批量重命名
        private void File_rename_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Workbook workbook = ThisAddIn.app.ActiveWorkbook;
            ThisAddIn.app.ScreenUpdating = false;
            ThisAddIn.app.DisplayAlerts = false;

            if (!string.IsNullOrEmpty(get_directory_path) && readFile == 1 && IsSheetExist(workbook, "_rename"))
            {
                //调用file.move或direction.move修改名
                for (int i = 2; i <= workbook.ActiveSheet.UsedRange.Rows.Count; i++)
                {
                    string cell1 = workbook.Worksheets["_rename"].Cells[i, 1].Value;
                    string cell2 = workbook.Worksheets["_rename"].Cells[i, 2].Value;
                    string cell3 = workbook.Worksheets["_rename"].Cells[i, 3].Value;
                    if (string.IsNullOrEmpty(cell1) || string.IsNullOrEmpty(cell2) || string.IsNullOrEmpty(cell3))
                    {
                        MessageBox.Show($"第{i}行有空格，请检查！该行对应文件将不被改名");
                        continue;
                    }
                    else
                    {
                        string full_path = cell1;
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
                MessageBox.Show("文件名修改完毕");
                Process.Start(get_directory_path);
                readFile = 0;
                ThisAddIn.app.DisplayAlerts = true;
                ThisAddIn.app.ScreenUpdating = true;

            }
            else
            {
                MessageBox.Show("没有选择文件夹，请先使用批读文件名功能后再使用该功能");
                readFile = 0;
            }
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

        private void Rename_mp3_Click(object sender, RibbonControlEventArgs e)
        {
            Form3 form3 = new Form3();
            form3.ShowDialog();
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
        private List<string> musicFiles = new List<string>();
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
        private void Mode_btuuon_Click(object sender, RibbonControlEventArgs e)
        {
            switch (playbackMode)
            {
                case PlaybackMode.Sequential:
                    playbackMode = PlaybackMode.AllLoop;
                    Mode_btuuon.Label = "全部循环";
                    Mode_btuuon.ScreenTip = "全部循环";
                    Mode_btuuon.Image = Properties.Resources.all_cycle;
                    break;
                case PlaybackMode.AllLoop:
                    playbackMode = PlaybackMode.SingleLoop;
                    Mode_btuuon.Label = "单曲循环";
                    Mode_btuuon.ScreenTip = "单曲循环";
                    Mode_btuuon.Image = Properties.Resources.single_cycle;
                    break;
                case PlaybackMode.SingleLoop:
                    playbackMode = PlaybackMode.Sequential;
                    Mode_btuuon.Label = "顺序播放";
                    Mode_btuuon.ScreenTip = "顺序播放";
                    Mode_btuuon.Image = Properties.Resources.order_play;
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

            syncContext = new SynchronizationContext();

            isAutoContinue = true;

            if (currentPlayState == PlaybackState.Stopped)
            {
                //timer1.Enabled = true;
                await PlayMusic();
            }
            else if (currentPlayState == PlaybackState.Paused)
            {
                waveOutEvent.Play();
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
                waveOutEvent.Play();

                await Task.Delay(200);

                // 等待播放完成
                while (!isMusicEnded)
                {
                    await Task.Delay(100);
                }
            }
        }

        private SynchronizationContext syncContext = SynchronizationContext.Current;

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
                //// 在UI线程上更新UI
                //syncContext.Send(state =>
                //{
                //    UpdateTrackInfo();
                //}, null);

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
                waveOutEvent.Stop();
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
                    waveOutEvent.Stop();
                    currentSongIndex = (currentSongIndex + 1) % musicFiles.Count;
                    await PlayMusic();
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
                    waveOutEvent.Stop();
                    currentSongIndex = (currentSongIndex - 1 + musicFiles.Count) % musicFiles.Count;
                    await PlayMusic();
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
                        $"时长：{audioFile.TotalTime.ToString(@"mm\:ss")}";
                    //return $"正在播放第{currentSongIndex + 1}首：{Path.GetFileName(musicFiles[currentSongIndex])}，" +
                    //    $"已播放时长：{audioFile.CurrentTime.ToString(@"mm\:ss")}，" + $"歌曲时长：{audioFile.TotalTime.ToString(@"mm\:ss")}";
                });
                try
                {
                    ThisAddIn.app.Application.StatusBar = trackInfo;
                }
                catch(System.Runtime.InteropServices.COMException ex)
                {
                    // 捕获COM异常，判断是否是锁定焦点的交互框
                    if (ex.Message.Contains("locked for editing"))
                    {
                        Debug.WriteLine("Excel交互框已锁定焦点");
                    }
                    else
                    {
                        Debug.WriteLine("发生其他COM异常:"+ex.Message);
                    }
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
    }
}

