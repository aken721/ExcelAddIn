using System;
using System.Drawing;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Http.Json;
using System.IO;
using System.Security.Cryptography;
using System.Threading;


namespace ExcelAddIn
{
    public partial class Form8 : Form
    {
        int i = 0;
        public Form8()
        {
            InitializeComponent();
            pbxPassword.BackgroundImage = Properties.Resources.eye_hide;
            txbKey.UseSystemPasswordChar = true;
            txbUrl.Text = @"https://api.deepseek.com/v1/chat/completions";
            txbUrl.ReadOnly = true;
        }

        private void Form8_Load(object sender, EventArgs e)
        {
            DecodeConfig();
            i = 0;
        }
        //重绘选项页布局
        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            //调整选项卡文字方向
            SolidBrush _Brush = new SolidBrush(Color.Black);//单色画刷
            RectangleF _TabTextArea = (RectangleF)tabControl1.GetTabRect(e.Index);//绘制区域
            StringFormat _sf = new StringFormat();//封装文本布局格式信息
            _sf.LineAlignment = StringAlignment.Center;
            _sf.Alignment = StringAlignment.Center;
            e.Graphics.DrawString(tabControl1.Controls[e.Index].Text, SystemInformation.MenuFont, _Brush, _TabTextArea, _sf);
        }

        private async Task<string> GetDeepSeekResponse(string apiKey,string model)
        {
            string apiUrl = txbUrl.Text;

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization =
                    new AuthenticationHeaderValue("Bearer", apiKey);

                // 构造请求体
                var requestBody = new
                {
                    model = model, // 选择的模型
                    messages = new[]
                    {
                        new { role = "user", content = "Hello, DeepSeek!" }
                    }
                };

                try
                {
                    var response = await client.PostAsJsonAsync(apiUrl, requestBody);
                    var responseContent = await response.Content.ReadAsStringAsync();

                    if (response.IsSuccessStatusCode)
                    {
                        return "连接成功";
                    }
                    else
                    {
                        // 返回错误信息和错误代码
                        return $"测试失败，错误代码：{response.StatusCode}，错误信息：{responseContent}";
                    }
                }
                catch(HttpRequestException ex)
                {
                    // 返回异常信息
                    return $"测试失败，发生异常：{ex.Message}";
                }                         
            }
        }


        private async void btnTest_Click(object sender, EventArgs e)
        {
            lblResult.Text = "开始检测......";
            txbKey.ReadOnly = true;
            cbbModel.Enabled = false;
            btnTest.Enabled = false;
            btnSave.Enabled = false;
            btnQuit.Enabled = false;
            this.Refresh();
            var result= await GetDeepSeekResponse(txbKey.Text,cbbModel.Text);
            if(!string.IsNullOrEmpty(result.ToString()))
            {
                lblResult.Text = result.ToString();
            }
            txbKey.ReadOnly = false;
            btnTest.Enabled = true;
            cbbModel.Enabled = true;
            btnSave.Enabled = true;
            btnQuit.Enabled = true;
        }
        private const string KeyFilePath = "encryption.key"; // 保存密钥和IV的文件路径
        private const string ConfigFilePath = "config.encrypted"; // 保存加密配置信息的文件路径

        private void btnSave_Click(object sender, EventArgs e)
        {
            string apiKey = txbKey.Text.Trim();
            string model = cbbModel.Text.Trim();
            string apiUrl = txbUrl.Text.Trim();

            if (string.IsNullOrEmpty(apiKey))
            {
                lblResult.Text = "API KEY不能为空";
                return;
            }
            if (string.IsNullOrEmpty(apiUrl))
            {
                lblResult.Text = "API地址不能为空";
                return;
            }

            string content = $"api-key^{apiKey};model^{model};api-url^{apiUrl}";

            try
            {
                // 获取或生成密钥和IV
                (byte[] key, byte[] iv) = GetOrCreateEncryptionKey();

                // 加密文本内容
                string encryptedContent = EncryptString(content, key, iv);

                // 将加密内容写入文件
                File.WriteAllText(ConfigFilePath, encryptedContent);

                lblResult.Text = "保存成功";
                this.Refresh();
                Thread.Sleep(1000);
                this.Dispose();
            }
            catch (Exception ex)
            {
                lblResult.Text = $"发生错误：{ex.Message}";
            }
        }

        //读取配置信息
        private void DecodeConfig()
        {
            if (!File.Exists(ConfigFilePath))
            {
                lblResult.Text = "未初始化配置";
                txbKey.Text = "";
                cbbModel.SelectedIndex = 0;
                return;
            }

            try
            {
                // 获取密钥和IV
                (byte[] key, byte[] iv) = GetEncryptionKey();

                // 读取加密内容
                string encryptedContent = File.ReadAllText(ConfigFilePath);

                // 解密文本内容
                string decryptedContent = DecryptString(encryptedContent, key, iv);

                // 解析配置信息
                txbKey.Text = decryptedContent.Split(';')[0].Split('^')[1];
                cbbModel.Text = decryptedContent.Split(';')[1].Split('^')[1];
                txbUrl.Text = decryptedContent.Split(';')[2].Split('^')[1];
            }
            catch (Exception ex)
            {
                lblResult.Text = $"发生错误：{ex.Message}";
            }
        }


        private void btnQuit_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        // 获取或生成密钥和IV
        private (byte[], byte[]) GetOrCreateEncryptionKey()
        {
            if (File.Exists(KeyFilePath))
            {
                // 从文件中读取密钥和IV
                string[] lines = File.ReadAllLines(KeyFilePath);
                return (Convert.FromBase64String(lines[0]), Convert.FromBase64String(lines[1]));
            }
            else
            {
                // 自动生成密钥和IV
                using (Aes aesAlg = Aes.Create())
                {
                    aesAlg.GenerateKey();
                    aesAlg.GenerateIV();

                    byte[] key = aesAlg.Key;
                    byte[] iv = aesAlg.IV;

                    // 将密钥和IV保存到文件
                    File.WriteAllLines(KeyFilePath, new[] { Convert.ToBase64String(key), Convert.ToBase64String(iv) });

                    return (key, iv);
                }
            }
        }

        // 从文件中获取密钥和IV
        private (byte[], byte[]) GetEncryptionKey()
        {
            if (!File.Exists(KeyFilePath))
            {
                throw new FileNotFoundException("密钥文件不存在，请先进行加密操作。");
            }

            string[] lines = File.ReadAllLines(KeyFilePath);
            return (Convert.FromBase64String(lines[0]), Convert.FromBase64String(lines[1]));
        }

        // 加密字符串
        private string EncryptString(string plainText, byte[] key, byte[] iv)
        {
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = key;
                aesAlg.IV = iv;

                ICryptoTransform encryptor = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV);
                byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);
                byte[] encryptedBytes = encryptor.TransformFinalBlock(plainTextBytes, 0, plainTextBytes.Length);

                return Convert.ToBase64String(encryptedBytes);
            }
        }

        // 解密字符串
        private string DecryptString(string cipherText, byte[] key, byte[] iv)
        {
            using (Aes aesAlg = Aes.Create())
            {
                aesAlg.Key = key;
                aesAlg.IV = iv;

                ICryptoTransform decryptor = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV);
                byte[] cipherTextBytes = Convert.FromBase64String(cipherText);
                byte[] decryptedBytes = decryptor.TransformFinalBlock(cipherTextBytes, 0, cipherTextBytes.Length);

                return Encoding.UTF8.GetString(decryptedBytes);
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            if (File.Exists("encryption.key")) File.Delete("encryption.key");

            if (File.Exists("config.encrypted")) File.Delete("config.encrypted");

            txbKey.Text = "";
            cbbModel.SelectedIndex = 0;
            lblResult.Text = "";
            txbUrl.Text = @"https://api.deepseek.com/v1/chat/completions";
            txbUrl.ReadOnly = true;
        }

        private void pbxPassword_Click(object sender, EventArgs e)
        {
            if (txbKey.UseSystemPasswordChar) 
            {
                pbxPassword.BackgroundImage = Properties.Resources.eye_open;
                txbKey.UseSystemPasswordChar = false;
            }
            else
            {
                pbxPassword.BackgroundImage=Properties.Resources.eye_hide;
                txbKey.UseSystemPasswordChar=true;
            }
        }

        //apiUrl文本框保护机制，双击3次，激活文本框可编辑
        private void txbUrl_DoubleClick(object sender, EventArgs e)
        {
            if (i < 2)
            {
                i++;
                lblResult.Text = $"警告{i}：请不要轻易修改api地址";                
            }
            else if(i==2)
            {
                i++;
                lblResult.Text = $"警告{i}：修改api地址可能导致访问失败，除非你知道必须修改它了，那就再双击我1次试试";
                
            }
            else
            { 
            txbUrl.ReadOnly = false;
            }
        }
    }
}
