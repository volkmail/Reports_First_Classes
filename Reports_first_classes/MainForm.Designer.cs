namespace Reports_first_classes
{
    partial class MainForm
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.Button_OpenFileDialog = new System.Windows.Forms.Button();
            this.FilePath_TextBox = new System.Windows.Forms.TextBox();
            this.progress_bar_read = new System.Windows.Forms.ProgressBar();
            this.progress_label = new System.Windows.Forms.Label();
            this.Button_ReadExcelFile = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Button_WriteExcelFile = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Button_OpenFileDialog
            // 
            this.Button_OpenFileDialog.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Button_OpenFileDialog.Location = new System.Drawing.Point(11, 60);
            this.Button_OpenFileDialog.Margin = new System.Windows.Forms.Padding(2);
            this.Button_OpenFileDialog.Name = "Button_OpenFileDialog";
            this.Button_OpenFileDialog.Size = new System.Drawing.Size(143, 30);
            this.Button_OpenFileDialog.TabIndex = 0;
            this.Button_OpenFileDialog.Text = "Открыть файлы...";
            this.Button_OpenFileDialog.UseVisualStyleBackColor = true;
            this.Button_OpenFileDialog.Click += new System.EventHandler(this.Button_OpenFileDialog_Click);
            // 
            // FilePath_TextBox
            // 
            this.FilePath_TextBox.BackColor = System.Drawing.SystemColors.HighlightText;
            this.FilePath_TextBox.Location = new System.Drawing.Point(12, 30);
            this.FilePath_TextBox.Name = "FilePath_TextBox";
            this.FilePath_TextBox.Size = new System.Drawing.Size(487, 25);
            this.FilePath_TextBox.TabIndex = 10;
            // 
            // progress_bar_read
            // 
            this.progress_bar_read.ForeColor = System.Drawing.Color.Lime;
            this.progress_bar_read.Location = new System.Drawing.Point(12, 129);
            this.progress_bar_read.Name = "progress_bar_read";
            this.progress_bar_read.Size = new System.Drawing.Size(487, 30);
            this.progress_bar_read.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            this.progress_bar_read.TabIndex = 2;
            // 
            // progress_label
            // 
            this.progress_label.AutoSize = true;
            this.progress_label.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.progress_label.Location = new System.Drawing.Point(160, 108);
            this.progress_label.Name = "progress_label";
            this.progress_label.Size = new System.Drawing.Size(93, 17);
            this.progress_label.TabIndex = 3;
            this.progress_label.Text = "progress_label";
            // 
            // Button_ReadExcelFile
            // 
            this.Button_ReadExcelFile.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Button_ReadExcelFile.Location = new System.Drawing.Point(12, 95);
            this.Button_ReadExcelFile.Name = "Button_ReadExcelFile";
            this.Button_ReadExcelFile.Size = new System.Drawing.Size(142, 30);
            this.Button_ReadExcelFile.TabIndex = 4;
            this.Button_ReadExcelFile.Text = "Считать файл(ы)";
            this.Button_ReadExcelFile.UseVisualStyleBackColor = true;
            this.Button_ReadExcelFile.Click += new System.EventHandler(this.Button_ReadExcelFile_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Enabled = false;
            this.label1.Location = new System.Drawing.Point(290, 65);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 17);
            this.label1.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(12, 10);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(128, 17);
            this.label2.TabIndex = 11;
            this.label2.Text = "Выбранные файлы:";
            // 
            // Button_WriteExcelFile
            // 
            this.Button_WriteExcelFile.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Button_WriteExcelFile.Location = new System.Drawing.Point(163, 61);
            this.Button_WriteExcelFile.Name = "Button_WriteExcelFile";
            this.Button_WriteExcelFile.Size = new System.Drawing.Size(142, 30);
            this.Button_WriteExcelFile.TabIndex = 12;
            this.Button_WriteExcelFile.Text = "Записать файлы";
            this.Button_WriteExcelFile.UseVisualStyleBackColor = true;
            this.Button_WriteExcelFile.Click += new System.EventHandler(this.Button_WriteExcelFile_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(511, 163);
            this.Controls.Add(this.Button_WriteExcelFile);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Button_ReadExcelFile);
            this.Controls.Add(this.progress_label);
            this.Controls.Add(this.progress_bar_read);
            this.Controls.Add(this.FilePath_TextBox);
            this.Controls.Add(this.Button_OpenFileDialog);
            this.Font = new System.Drawing.Font("Times New Roman", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.ForeColor = System.Drawing.SystemColors.ControlText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Отчет по первым классам";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Button_OpenFileDialog;
        private System.Windows.Forms.TextBox FilePath_TextBox;
        private System.Windows.Forms.ProgressBar progress_bar_read;
        private System.Windows.Forms.Label progress_label;
        private System.Windows.Forms.Button Button_ReadExcelFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button Button_WriteExcelFile;
    }
}

