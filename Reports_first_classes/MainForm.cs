using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;

namespace Reports_first_classes
{
    public partial class MainForm : Form
    {
        ExcelManager excel_manager;
        BackgroundWorker backround_worker;
        string[] file_paths;
        string file_path_template;

        public MainForm()
        {
            InitializeComponent();
            IntializeFormElements();

            excel_manager = new ExcelManager();
            excel_manager.OnProgress_up += ProgressBar_up;
        }

        private void IntializeFormElements()
        {
            Button_ReadExcelFile.Visible = false;
            Button_ReadExcelFile.Enabled = false;
            Button_WriteExcelFile.Enabled = false;
            Button_WriteExcelFile.Visible = false;
            FilePath_TextBox.ReadOnly = true;
            progress_label.Visible = false;
            progress_bar_read.BackColor = Color.LightGreen;
            progress_bar_read.Minimum = 0;
            progress_bar_read.Maximum = 100;
            progress_bar_read.Step = 5;
            progress_bar_read.Visible = false;
            progress_bar_read.UseWaitCursor = true;
        }

        private void Button_OpenFileDialog_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 0;
            openFileDialog.Multiselect = true;
            openFileDialog.RestoreDirectory = true;
            file_paths = new string[3]; // Небольшой костыль, потому что пока всего 3 предмета нужно считать

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string text = string.Empty;
                foreach (string name in openFileDialog.FileNames)
                    text += $"\"{Path.GetFileName(name)}\", ";

                FilePath_TextBox.Text = $"{text.Remove(text.LastIndexOf(','))}.";

                if (openFileDialog.FileNames.Length == 3)
                {
                    file_paths = openFileDialog.FileNames;
                    Button_ReadExcelFile.Enabled = true;
                    Button_ReadExcelFile.Visible = true;
                    Button_WriteExcelFile.Enabled = false;
                    Button_WriteExcelFile.Visible = false;
                }
                else if (openFileDialog.FileNames.Length == 1 && Regex.IsMatch(Path.GetFileName(openFileDialog.FileName), @"^XXXX"))
                {
                    file_path_template = openFileDialog.FileName;
                    Button_ReadExcelFile.Enabled = false;
                    Button_ReadExcelFile.Visible = false;
                    Button_WriteExcelFile.Enabled = true;
                    Button_WriteExcelFile.Visible = true;
                }
                else
                {
                    FilePath_TextBox.Text = String.Empty;
                    Button_ReadExcelFile.Enabled = false;
                    Button_ReadExcelFile.Visible = false;
                    Button_WriteExcelFile.Enabled = false;
                    Button_WriteExcelFile.Visible = false;
                    MessageBox.Show("Выберете либо 3 файла для считывания, \nлибо 1 файл шаблона для записи !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            openFileDialog.Dispose();
            openFileDialog = null;
            GC.Collect();
        }

        private void Button_ReadExcelFile_Click(object sender, EventArgs e)
        {
            progress_bar_read.Visible = true;
            Button_OpenFileDialog.Enabled = false;
            Button_ReadExcelFile.Enabled = false;
            progress_label.Visible = true;
            progress_label.Text = String.Empty;

            backround_worker = new BackgroundWorker();
            backround_worker.WorkerReportsProgress = true;
            backround_worker.WorkerSupportsCancellation = true;

            backround_worker.DoWork += (obj, ev) =>
            {
                excel_manager.ExcelReader(file_paths);
                if (file_paths.Length == 3)
                    ev.Result = "Считывание завершено успешно";
                else
                    ev.Result = "Принудительное завершение чтения";
            };

            backround_worker.ProgressChanged += (obj, ev) =>
            {
                progress_label.Text = $"Считывание \"{Path.GetFileName(ev.UserState.ToString())}\" {ev.ProgressPercentage}%";
                progress_bar_read.Value = ev.ProgressPercentage;
            };

            backround_worker.RunWorkerCompleted += (obj, ev) =>
            {
                if (MessageBox.Show($"{ev.Result}", "Завершение чтения",
                    MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    Button_OpenFileDialog.Enabled = true;
                    Button_OpenFileDialog.Text = "Открыть шаблон";
                    Button_ReadExcelFile.Text = "Записать файлы";
                    Button_ReadExcelFile.Enabled = false;
                    Button_ReadExcelFile.Visible = false;
                    FilePath_TextBox.Text = String.Empty;
                    progress_bar_read.Value = 0;
                    progress_bar_read.Visible = false;
                    progress_label.Visible = false;
                }
            };

            backround_worker.RunWorkerAsync();
        }

        private void ProgressBar_up(int check, string text_for_show_on_label)
        {
            if (backround_worker != null)
                backround_worker.ReportProgress(check, text_for_show_on_label);
        }

        private void Button_WriteExcelFile_Click(object sender, EventArgs e)
        {
            if (excel_manager.ChekReadyToWrite())
            {
                Button_OpenFileDialog.Enabled = false;
                Button_WriteExcelFile.Enabled = false;
                progress_label.Text = String.Empty;
                progress_bar_read.Visible = true;
                progress_label.Visible = true;

                backround_worker = new BackgroundWorker();
                backround_worker.WorkerSupportsCancellation = true;
                backround_worker.WorkerReportsProgress = true;

                backround_worker.DoWork += (obj, ev) =>
                {
                    excel_manager.ExcelWriter(file_path_template);
                };

                backround_worker.ProgressChanged += (obj, ev) =>
                {
                    progress_label.Text = $"Запись в лист \"{ev.UserState.ToString()}\" {ev.ProgressPercentage}%";
                    progress_bar_read.Value = ev.ProgressPercentage;
                };

                backround_worker.RunWorkerCompleted += (obj, ev) =>
                {
                    FilePath_TextBox.Text = String.Empty;
                    Button_OpenFileDialog.Enabled = true;
                    Button_WriteExcelFile.Enabled = false;
                    Button_WriteExcelFile.Visible = false;
                    MessageBox.Show("Запись в шаблон завершена !");
                };

                backround_worker.RunWorkerAsync();
            }
        }
    }
}
