using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using NAudio.Wave;

namespace TableMagic.Skills
{
    public class ExcelMusicSkill : ISkill
    {
        public string Name => "ExcelMusic";
        public string Description => "音乐播放技能，支持播放本地音乐文件和播放列表管理";

        public List<SkillTool> GetTools()
        {
            return new List<SkillTool>
            {
                new SkillTool
                {
                    Name = "play_music",
                    Description = "播放音乐文件。当用户要求播放音乐时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "filePath", new { type = "string", description = "音乐文件路径" } },
                                { "volume", new { type = "integer", description = "音量（0-100，默认50）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "filePath" }
                },
                new SkillTool
                {
                    Name = "play_playlist",
                    Description = "播放文件夹中的所有音乐文件。当用户要求播放播放列表、播放文件夹音乐时使用此工具。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "folderPath", new { type = "string", description = "音乐文件夹路径" } },
                                { "shuffle", new { type = "boolean", description = "是否随机播放（默认false）" } },
                                { "volume", new { type = "integer", description = "音量（0-100，默认50）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "folderPath" }
                },
                new SkillTool
                {
                    Name = "pause_music",
                    Description = "暂停/继续播放音乐。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "stop_music",
                    Description = "停止播放音乐。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "set_volume",
                    Description = "设置音量。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>
                            {
                                { "volume", new { type = "integer", description = "音量（0-100）" } }
                            }
                        }
                    },
                    RequiredParameters = new List<string> { "volume" }
                },
                new SkillTool
                {
                    Name = "next_track",
                    Description = "播放下一首。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "previous_track",
                    Description = "播放上一首。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                },
                new SkillTool
                {
                    Name = "get_player_status",
                    Description = "获取播放器状态。",
                    Parameters = new Dictionary<string, object>
                    {
                        { "type", "object" },
                        { "properties", new Dictionary<string, object>() }
                    },
                    RequiredParameters = new List<string>()
                }
            };
        }

        public async Task<SkillResult> ExecuteToolAsync(string toolName, Dictionary<string, object> arguments)
        {
            try
            {
                switch (toolName)
                {
                    case "play_music":
                        return PlayMusic(arguments);
                    case "play_playlist":
                        return PlayPlaylist(arguments);
                    case "pause_music":
                        return PauseMusic();
                    case "stop_music":
                        return StopMusic();
                    case "set_volume":
                        return SetVolume(arguments);
                    case "next_track":
                        return NextTrack();
                    case "previous_track":
                        return PreviousTrack();
                    case "get_player_status":
                        return GetPlayerStatus();
                    default:
                        return new SkillResult { Success = false, Error = $"Tool {toolName} not implemented in ExcelMusicSkill" };
                }
            }
            catch (Exception ex)
            {
                return new SkillResult { Success = false, Error = ex.Message };
            }
        }

        private static WaveOutEvent _waveOut;
        private static AudioFileReader _audioFile;
        private static List<string> _playlist = new List<string>();
        private static int _currentTrackIndex = 0;
        private static bool _shuffle = false;
        private static PlaybackState _playbackState = PlaybackState.Stopped;
        private static float _volume = 0.5f;

        private enum PlaybackState
        {
            Stopped,
            Playing,
            Paused
        }

        private SkillResult PlayMusic(Dictionary<string, object> arguments)
        {
            var filePath = arguments["filePath"].ToString();
            var volume = arguments.ContainsKey("volume") ? Convert.ToInt32(arguments["volume"]) : 50;

            if (!File.Exists(filePath))
                return new SkillResult { Success = false, Error = $"文件不存在: {filePath}" };

            DisposeWavePlayer();

            _audioFile = new AudioFileReader(filePath);
            _waveOut = new WaveOutEvent();
            _waveOut.Init(_audioFile);
            _waveOut.Volume = volume / 100f;
            _volume = volume / 100f;
            _waveOut.PlaybackStopped += WaveOut_PlaybackStopped;
            _waveOut.Play();

            _playlist = new List<string> { filePath };
            _currentTrackIndex = 0;
            _playbackState = PlaybackState.Playing;

            return new SkillResult { Success = true, Content = $"正在播放: {Path.GetFileName(filePath)}" };
        }

        private SkillResult PlayPlaylist(Dictionary<string, object> arguments)
        {
            var folderPath = arguments["folderPath"].ToString();
            var shuffle = arguments.ContainsKey("shuffle") && Convert.ToBoolean(arguments["shuffle"]);
            var volume = arguments.ContainsKey("volume") ? Convert.ToInt32(arguments["volume"]) : 50;

            if (!Directory.Exists(folderPath))
                return new SkillResult { Success = false, Error = $"文件夹不存在: {folderPath}" };

            var extensions = new[] { "*.mp3", "*.wav", "*.wma", "*.m4a", "*.flac", "*.aac", "*.aiff" };
            var files = new List<string>();

            foreach (var ext in extensions)
            {
                files.AddRange(Directory.GetFiles(folderPath, ext, SearchOption.AllDirectories));
            }

            if (files.Count == 0)
                return new SkillResult { Success = false, Error = "文件夹中没有音乐文件" };

            _playlist = files;
            _shuffle = shuffle;

            if (shuffle)
            {
                var random = new Random();
                _playlist = _playlist.OrderBy(x => random.Next()).ToList();
            }

            _currentTrackIndex = 0;

            DisposeWavePlayer();

            _audioFile = new AudioFileReader(_playlist[_currentTrackIndex]);
            _waveOut = new WaveOutEvent();
            _waveOut.Init(_audioFile);
            _waveOut.Volume = volume / 100f;
            _volume = volume / 100f;
            _waveOut.PlaybackStopped += WaveOut_PlaybackStopped;
            _waveOut.Play();

            _playbackState = PlaybackState.Playing;

            return new SkillResult { Success = true, Content = $"正在播放播放列表，共 {_playlist.Count} 首歌曲\n当前: {Path.GetFileName(_playlist[_currentTrackIndex])}" };
        }

        private SkillResult PauseMusic()
        {
            if (_waveOut == null)
                return new SkillResult { Success = false, Error = "没有正在播放的音乐" };

            if (_playbackState == PlaybackState.Playing)
            {
                _waveOut.Pause();
                _playbackState = PlaybackState.Paused;
                return new SkillResult { Success = true, Content = "音乐已暂停" };
            }
            else if (_playbackState == PlaybackState.Paused)
            {
                _waveOut.Play();
                _playbackState = PlaybackState.Playing;
                return new SkillResult { Success = true, Content = "音乐继续播放" };
            }
            else
            {
                return new SkillResult { Success = false, Error = "当前状态无法暂停" };
            }
        }

        private SkillResult StopMusic()
        {
            if (_waveOut == null)
                return new SkillResult { Success = true, Content = "没有正在播放的音乐" };

            _waveOut.Stop();
            _playbackState = PlaybackState.Stopped;
            return new SkillResult { Success = true, Content = "音乐已停止" };
        }

        private SkillResult SetVolume(Dictionary<string, object> arguments)
        {
            var volume = Convert.ToInt32(arguments["volume"]);

            if (_waveOut == null)
            {
                _volume = Math.Max(0, Math.Min(100, volume)) / 100f;
                return new SkillResult { Success = true, Content = $"音量已设置为: {volume}" };
            }

            _waveOut.Volume = Math.Max(0, Math.Min(100, volume)) / 100f;
            _volume = _waveOut.Volume;
            return new SkillResult { Success = true, Content = $"音量已设置为: {volume}" };
        }

        private SkillResult NextTrack()
        {
            if (_playlist.Count == 0)
                return new SkillResult { Success = false, Error = "播放列表为空" };

            _currentTrackIndex = (_currentTrackIndex + 1) % _playlist.Count;

            DisposeWavePlayer();

            _audioFile = new AudioFileReader(_playlist[_currentTrackIndex]);
            _waveOut = new WaveOutEvent();
            _waveOut.Init(_audioFile);
            _waveOut.Volume = _volume;
            _waveOut.PlaybackStopped += WaveOut_PlaybackStopped;
            _waveOut.Play();

            _playbackState = PlaybackState.Playing;

            return new SkillResult { Success = true, Content = $"正在播放: {Path.GetFileName(_playlist[_currentTrackIndex])}" };
        }

        private SkillResult PreviousTrack()
        {
            if (_playlist.Count == 0)
                return new SkillResult { Success = false, Error = "播放列表为空" };

            _currentTrackIndex = (_currentTrackIndex - 1 + _playlist.Count) % _playlist.Count;

            DisposeWavePlayer();

            _audioFile = new AudioFileReader(_playlist[_currentTrackIndex]);
            _waveOut = new WaveOutEvent();
            _waveOut.Init(_audioFile);
            _waveOut.Volume = _volume;
            _waveOut.PlaybackStopped += WaveOut_PlaybackStopped;
            _waveOut.Play();

            _playbackState = PlaybackState.Playing;

            return new SkillResult { Success = true, Content = $"正在播放: {Path.GetFileName(_playlist[_currentTrackIndex])}" };
        }

        private SkillResult GetPlayerStatus()
        {
            if (_waveOut == null)
                return new SkillResult { Success = true, Content = "播放器未初始化" };

            var state = _playbackState switch
            {
                PlaybackState.Stopped => "已停止",
                PlaybackState.Playing => "正在播放",
                PlaybackState.Paused => "已暂停",
                _ => "未知状态"
            };

            var currentTrack = _playlist.Count > 0 && _currentTrackIndex < _playlist.Count
                ? Path.GetFileName(_playlist[_currentTrackIndex])
                : "无";

            var currentTime = _audioFile?.CurrentTime.ToString(@"mm\:ss") ?? "00:00";
            var totalTime = _audioFile?.TotalTime.ToString(@"mm\:ss") ?? "00:00";

            return new SkillResult 
            { 
                Success = true, 
                Content = $"播放器状态: {state}\n当前曲目: {currentTrack}\n播放进度: {currentTime}/{totalTime}\n音量: {(int)(_volume * 100)}\n播放列表: {_playlist.Count} 首\n随机播放: {(_shuffle ? "是" : "否")}" 
            };
        }

        private void WaveOut_PlaybackStopped(object sender, StoppedEventArgs e)
        {
            if (_playlist.Count > 1)
            {
                _currentTrackIndex = (_currentTrackIndex + 1) % _playlist.Count;

                try
                {
                    DisposeWavePlayer();

                    _audioFile = new AudioFileReader(_playlist[_currentTrackIndex]);
                    _waveOut = new WaveOutEvent();
                    _waveOut.Init(_audioFile);
                    _waveOut.Volume = _volume;
                    _waveOut.PlaybackStopped += WaveOut_PlaybackStopped;
                    _waveOut.Play();

                    _playbackState = PlaybackState.Playing;
                }
                catch { }
            }
        }

        private void DisposeWavePlayer()
        {
            if (_waveOut != null)
            {
                _waveOut.PlaybackStopped -= WaveOut_PlaybackStopped;
                _waveOut.Stop();
                _waveOut.Dispose();
                _waveOut = null;
            }

            if (_audioFile != null)
            {
                _audioFile.Dispose();
                _audioFile = null;
            }
        }
    }
}
