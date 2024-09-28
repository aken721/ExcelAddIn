namespace ExcelAddIn
{
    partial class Form5
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form5));
            this.picture_radioButton = new System.Windows.Forms.RadioButton();
            this.webcam_radioButton = new System.Windows.Forms.RadioButton();
            this.scan_button = new System.Windows.Forms.Button();
            this.quit_button = new System.Windows.Forms.Button();
            this.qr_image_pictureBox = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.folder_path_label = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.qr_image_pictureBox)).BeginInit();
            this.SuspendLayout();
            // 
            // picture_radioButton
            // 
            this.picture_radioButton.AutoSize = true;
            this.picture_radioButton.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.picture_radioButton.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.picture_radioButton.Location = new System.Drawing.Point(159, 68);
            this.picture_radioButton.Name = "picture_radioButton";
            this.picture_radioButton.Size = new System.Drawing.Size(83, 24);
            this.picture_radioButton.TabIndex = 0;
            this.picture_radioButton.TabStop = true;
            this.picture_radioButton.Text = "图片识别";
            this.picture_radioButton.UseVisualStyleBackColor = true;
            this.picture_radioButton.CheckedChanged += new System.EventHandler(this.picture_radioButton_CheckedChanged);
            // 
            // webcam_radioButton
            // 
            this.webcam_radioButton.AutoSize = true;
            this.webcam_radioButton.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.webcam_radioButton.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.webcam_radioButton.Location = new System.Drawing.Point(364, 68);
            this.webcam_radioButton.Name = "webcam_radioButton";
            this.webcam_radioButton.Size = new System.Drawing.Size(83, 24);
            this.webcam_radioButton.TabIndex = 1;
            this.webcam_radioButton.TabStop = true;
            this.webcam_radioButton.Text = "扫描识别";
            this.webcam_radioButton.UseVisualStyleBackColor = true;
            this.webcam_radioButton.CheckedChanged += new System.EventHandler(this.webcam_radioButton_CheckedChanged);
            // 
            // scan_button
            // 
            this.scan_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.scan_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.scan_button.Location = new System.Drawing.Point(143, 125);
            this.scan_button.Name = "scan_button";
            this.scan_button.Size = new System.Drawing.Size(76, 32);
            this.scan_button.TabIndex = 3;
            this.scan_button.Text = "识别";
            this.scan_button.UseVisualStyleBackColor = true;
            this.scan_button.Click += new System.EventHandler(this.scan_button_Click);
            // 
            // quit_button
            // 
            this.quit_button.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.quit_button.ForeColor = System.Drawing.Color.SeaGreen;
            this.quit_button.Location = new System.Drawing.Point(381, 125);
            this.quit_button.Name = "quit_button";
            this.quit_button.Size = new System.Drawing.Size(76, 32);
            this.quit_button.TabIndex = 4;
            this.quit_button.Text = "退出";
            this.quit_button.UseVisualStyleBackColor = true;
            this.quit_button.Click += new System.EventHandler(this.quit_button_Click);
            // 
            // qr_image_pictureBox
            // 
            this.qr_image_pictureBox.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.qr_image_pictureBox.Location = new System.Drawing.Point(147, 216);
            this.qr_image_pictureBox.Name = "qr_image_pictureBox";
            this.qr_image_pictureBox.Size = new System.Drawing.Size(310, 150);
            this.qr_image_pictureBox.TabIndex = 5;
            this.qr_image_pictureBox.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("微软雅黑", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.DarkSlateBlue;
            this.label1.Location = new System.Drawing.Point(205, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(192, 27);
            this.label1.TabIndex = 6;
            this.label1.Text = "二维码批量识别工具";
            // 
            // folder_path_label
            // 
            this.folder_path_label.AutoSize = true;
            this.folder_path_label.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.folder_path_label.ForeColor = System.Drawing.Color.SaddleBrown;
            this.folder_path_label.Location = new System.Drawing.Point(144, 183);
            this.folder_path_label.Name = "folder_path_label";
            this.folder_path_label.Size = new System.Drawing.Size(0, 17);
            this.folder_path_label.TabIndex = 7;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form5
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(632, 375);
            this.Controls.Add(this.folder_path_label);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.qr_image_pictureBox);
            this.Controls.Add(this.quit_button);
            this.Controls.Add(this.scan_button);
            this.Controls.Add(this.webcam_radioButton);
            this.Controls.Add(this.picture_radioButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form5";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form5";
            ((System.ComponentModel.ISupportInitialize)(this.qr_image_pictureBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton picture_radioButton;
        private System.Windows.Forms.RadioButton webcam_radioButton;
        private System.Windows.Forms.Button scan_button;
        private System.Windows.Forms.Button quit_button;
        private System.Windows.Forms.PictureBox qr_image_pictureBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label folder_path_label;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}