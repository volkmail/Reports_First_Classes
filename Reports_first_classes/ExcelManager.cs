using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Windows.Forms;

namespace Reports_first_classes
{
    class ExcelManager
    {
        public delegate void UpProgressBar(int check, string text_for_show_on_label);
        public event UpProgressBar OnProgress_up;
        string file_path { get; set; } // Путь до файла для считывания
        public List<Record> records_maths;// Список считываемых записей математики
        public List<Record> records_rus;// Список считываемых записей русского
        public List<Record> records_reading; // Список считываемых записей чтения
        List<string> temp_values; // Промежуточные значения полей для создания записи (Без заданий и баллов за них)
        List<Tuple<double, double>> temp_task_results; // Item1 - номер задания, Item2 - Баллы
        List<Tuple<string, string>> columnName_columnIndex; //Список заголовков столбцов и их индекса (буква ссылки яйчеки в Excel)
        DocumentFormat.OpenXml.UInt32Value cell_style_fio;
        DocumentFormat.OpenXml.UInt32Value cell_style_variant;
        DocumentFormat.OpenXml.UInt32Value cell_style_tasks;

        struct DataForFill // Структура для простого заполнения первых трех листов в шаблоне  
        {
            public string student_FIO { get; private set; }
            public string variant { get; private set; }
            public List<Tuple<double, double>> task_results { get; private set; }

            public DataForFill(string student_FIO, string variant, List<Tuple<double, double>> task_results)
            {
                this.student_FIO = student_FIO;
                this.variant = variant;
                this.task_results = task_results;
            }
        }

        struct DataSortedByRegion
        {
            public string region { get; private set; }
            public int schools_count { get; private set; }
            public int clasess_count { get; private set; }
            public int student_count { get; private set; }
            public int student_write_count { get; private set; }
            public int low_lvl_score_count { get; private set; }
            public int upper_low_lvl_score_count { get; private set; }
            public int base_lvl_score_count { get; private set; }
            public int upper_base_lvl_score_count { get; private set; }
            public int high_lvl_score_count { get; private set; }
            public int total_score { get; private set; }
            public int base_lvl_count { get; private set; }
            public int high_lvl_count { get; private set; }
            public double min_total_score { get; private set; }
            public double max_total_score { get; private set; }

            public DataSortedByRegion(string region, int schools_count, int clasess_count, int student_count, int student_write_count,
                int low_lvl_score_count, int upper_low_lvl_score_count, int base_lvl_score_count,
                int upper_base_lvl_score_count, int high_lvl_score_count, int total_score, int base_lvl_count, int high_lvl_count, double min_total_score, double max_total_score)
            {
                this.region = region;
                this.schools_count = schools_count;
                this.clasess_count = clasess_count;
                this.student_count = student_count;
                this.student_write_count = student_write_count;
                this.low_lvl_score_count = low_lvl_score_count;
                this.upper_low_lvl_score_count = upper_low_lvl_score_count;
                this.base_lvl_score_count = base_lvl_score_count;
                this.upper_base_lvl_score_count = upper_base_lvl_score_count;
                this.high_lvl_score_count = high_lvl_score_count;
                this.total_score = total_score;
                this.base_lvl_count = base_lvl_count;
                this.high_lvl_count = high_lvl_count;
                this.min_total_score = min_total_score;
                this.max_total_score = max_total_score;
            }
        }

        //public ExcelManager()
        //{
        //    records_maths = null;
        //    records_rus = null;
        //    records_reading = null;
        //    temp_values = null;
        //    temp_task_results = null;
        //    columnName_columnIndex = null;
        //}

        public void ExcelWriter(string template_file_path)
        {
            try
            {
                using (SpreadsheetDocument excel_doc = SpreadsheetDocument.Open(template_file_path, true))
                {
                    SharedStringTable sharedString;
                    if (excel_doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                        sharedString = excel_doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault().SharedStringTable;
                    else
                        sharedString = excel_doc.WorkbookPart.AddNewPart<SharedStringTablePart>().SharedStringTable;

                    foreach (Sheet sheet in excel_doc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().ToList())
                    {
                        if (Regex.IsMatch(sheet.Name.ToString(), @"^МА"))
                            WriteToSheet((excel_doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart).Worksheet, sharedString, records_maths, "fill", sheet.Name.ToString());
                        if (Regex.IsMatch(sheet.Name.ToString(), @"^РУ"))
                            WriteToSheet((excel_doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart).Worksheet, sharedString, records_rus, "fill", sheet.Name.ToString());
                        if (Regex.IsMatch(sheet.Name.ToString(), @"^ЧТ"))
                            WriteToSheet((excel_doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart).Worksheet, sharedString, records_reading, "fill", sheet.Name.ToString());
                        if (Regex.IsMatch(sheet.Name.ToString(), @"свод_МА$"))
                            WriteToSheet((excel_doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart).Worksheet, sharedString, records_maths, "calculate_and_fill", sheet.Name.ToString());
                        if (Regex.IsMatch(sheet.Name.ToString(), @"свод_РУ$"))
                            WriteToSheet((excel_doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart).Worksheet, sharedString, records_rus, "calculate_and_fill", sheet.Name.ToString());
                        if (Regex.IsMatch(sheet.Name.ToString(), @"свод_ЧТ$"))
                            WriteToSheet((excel_doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart).Worksheet, sharedString, records_reading, "calculate_and_fill", sheet.Name.ToString());
                    }
                    excel_doc.Close();
                }
            }
            catch
            {
                MessageBox.Show($"Файл {Path.GetFileName(template_file_path)} уже открыт в другой программе.\nПожалуйста, закройте его и начните процедуру записи заново.",
                    "Ошибка",MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void WriteToSheet(Worksheet worksheet, SharedStringTable sharedString, List<Record> records, string action, string sheet_name)
        {
            switch (action)
            {
                case "fill":
                    {
                        List<DataForFill> data_for_fill = new List<DataForFill>(); // Создаем список структур, в которых будут содержаться только интересующие нас данные

                        Row task_result_row = worksheet.Descendants<Row>().Where(r => r.RowIndex == 11).FirstOrDefault(); // т.к. начинаем писать в файл с 11 строки

                        foreach (Cell cell in task_result_row) // Собираем буквенные ссылки на ячейки с заданиями 
                        {
                            if (cell.CellValue != null && !String.IsNullOrWhiteSpace(cell.CellValue.InnerText) && cell.DataType?.Value != CellValues.Error)
                            {
                                if (cell.DataType != null)
                                {
                                    if (cell.DataType == CellValues.SharedString)
                                    {
                                        if (int.TryParse(sharedString.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText, out int result_int)
                                            || double.TryParse(sharedString.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText,
                                            NumberStyles.Any, CultureInfo.InvariantCulture.NumberFormat, out double result_double))
                                        {
                                            columnName_columnIndex.Add(new Tuple<string, string>(sharedString.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText,
                                              Regex.Match(cell.CellReference, "[A-Z]{1,}").Value));
                                        }
                                    }
                                }
                                else
                                {
                                    if (int.TryParse(sharedString.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText, out int result_int)
                                        || double.TryParse(cell.CellValue.InnerText, NumberStyles.Any, CultureInfo.InvariantCulture.NumberFormat, out double result_double))
                                    {
                                        columnName_columnIndex.Add(new Tuple<string, string>(cell.CellValue.InnerText,
                                          Regex.Match(cell.CellReference, "[A-Z]{1,}").Value));
                                    }
                                }
                            }
                        } // Собираем буквенные ссылки на ячейки с заданиями

                        foreach (Record item in records) // Заполняем список data
                            data_for_fill.Add(new DataForFill(item.student_FIO, item.variant, item.task_results));

                        Row sample_row = worksheet.Descendants<Row>().Where(i => i.RowIndex == 25).FirstOrDefault();

                        foreach (Cell sample_cell in sample_row)
                        {
                            if (sample_cell.CellReference == "D25")
                                cell_style_fio = sample_cell.StyleIndex;
                            if (sample_cell.CellReference == "E25")
                                cell_style_variant = sample_cell.StyleIndex;
                            if (sample_cell.CellReference == "F25")
                                cell_style_tasks = sample_cell.StyleIndex;
                        }
                              
                        foreach (DataForFill data in data_for_fill) // Бежим по нашей структуре и заполняем ячейки
                        {
                            // Insert the text into the SharedStringTablePart.
                            int index_FIO = InsertSharedStringItem(data.student_FIO, sharedString);
                            int index_variant = InsertSharedStringItem(data.variant, sharedString);

                            Cell cellFIO = InsertCellInWorksheet("D", (uint)data_for_fill.IndexOf(data) + 25, worksheet);
                            Cell cellVariant = InsertCellInWorksheet("E", (uint)data_for_fill.IndexOf(data) + 25, worksheet);

                            // Задаем значения ячеек.
                            cellFIO.CellValue = new CellValue(index_FIO.ToString());
                            cellVariant.CellValue = new CellValue(index_variant.ToString());

                            cellFIO.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.SharedString);
                            cellVariant.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.SharedString);

                            cellFIO.StyleIndex = cell_style_fio;
                            cellVariant.StyleIndex = cell_style_variant;

                            foreach (Tuple<double, double> temp_task in data.task_results) // Заполяем ячейки результатов
                            {
                                foreach (Tuple<string, string> task_ref in columnName_columnIndex)
                                {
                                    if (temp_task.Item1 == double.Parse(task_ref.Item1, NumberStyles.Any, CultureInfo.InvariantCulture.NumberFormat))
                                    {
                                        int c_row_index = (int)data_for_fill.IndexOf(data) + 25;
                                        //if (c_row_index > 165)
                                        //    MessageBox.Show("Watch");
                                        Cell cell_task_result = InsertCellInWorksheet(task_ref.Item2, (uint)data_for_fill.IndexOf(data) + 25, worksheet);
                                        cell_task_result.CellValue = new CellValue(temp_task.Item2.ToString());
                                        cell_task_result.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                                        cell_task_result.StyleIndex = cell_style_tasks;
                                    }
                                }
                            }

                            // Save the new worksheet.
                            worksheet.Save();
                            double count_record = data_for_fill.IndexOf(data);
                            double progress = (((double)data_for_fill.IndexOf(data) + 1) / (double)(data_for_fill.Count - 1)) * 100;
                            OnProgress_up((int)progress, sheet_name);
                            progress = 0;
                        }

                        break;
                    }
                case "calculate_and_fill": 
                    {
                        var region_records = records.GroupBy(r => r.region).ToList();

                        List<DataSortedByRegion> region_data_to_fill = new List<DataSortedByRegion>(region_records.Count);

                        foreach (IGrouping<string, Record> record in region_records) // Заполняем список данных, собранных с кадого региона отдельно
                        {
                            string region = record.Key;
                            int schools_count = record.GroupBy(s => s.school_id).ToList().Count;
                            int clasess_count = record.GroupBy(s => s.class_id).ToList().Count;
                            int student_count = record.GroupBy(s => s.student_FIO).ToList().Count;
                            int student_write_count = record.Where(s => s.variant != "N").ToList().Count;
                            int low_lvl_score_count = record.Where(l => Regex.IsMatch(l.understand_lvl, "Низкий", RegexOptions.IgnoreCase)).ToList().Count;
                            int upper_low_lvl_score_count = record.Where(l => Regex.IsMatch(l.understand_lvl, "Пониженный", RegexOptions.IgnoreCase)).ToList().Count;
                            int base_lvl_score_count = record.Where(l => Regex.IsMatch(l.understand_lvl, "Базовый", RegexOptions.IgnoreCase)).ToList().Count;
                            int upper_base_lvl_score_count = record.Where(l => Regex.IsMatch(l.understand_lvl, "Повышенный", RegexOptions.IgnoreCase)).ToList().Count;
                            int high_lvl_score_count = record.Where(l => Regex.IsMatch(l.understand_lvl, "Высокий", RegexOptions.IgnoreCase)).ToList().Count;
                            double min_total_score = record.Where(v => v.variant != "N").Min(t => t.total_score);
                            double max_total_score = record.Where(v => v.variant != "N").Max(t => t.total_score);
                            double total_score = 0;
                            double base_lvl_count = 0;
                            double high_lvl_count = 0;

                            foreach (var data in record)
                            {
                                total_score += data.total_score;
                                base_lvl_count += data.base_lvl;
                                high_lvl_count += data.high_lvl;
                            }

                            region_data_to_fill.Add(new DataSortedByRegion(region, schools_count, clasess_count, student_count, student_write_count,
                                low_lvl_score_count, upper_low_lvl_score_count, base_lvl_score_count,
                                upper_base_lvl_score_count, high_lvl_score_count, (int)Math.Truncate(total_score), (int)Math.Truncate(base_lvl_count),
                                (int)Math.Truncate(high_lvl_count), min_total_score, max_total_score));
                        }

                        foreach (DataSortedByRegion data in region_data_to_fill)
                        {
                            int index_region = InsertSharedStringItem(data.region, sharedString);
                            //int index_schools_count = InsertSharedStringItem(data.schools_count.ToString(), sharedString);
                            //int index_clasess_count = InsertSharedStringItem(data.clasess_count.ToString(), sharedString);
                            //int index_student_count = InsertSharedStringItem(data.student_count.ToString(), sharedString);
                            //int index_student_write_count = InsertSharedStringItem(data.student_write_count.ToString(), sharedString);
                            //int index_low_lvl_score_count = InsertSharedStringItem(data.low_lvl_score_count.ToString(), sharedString);
                            //int index_upper_low_lvl_score_count = InsertSharedStringItem(data.upper_low_lvl_score_count.ToString(), sharedString);
                            //int index_base_lvl_score_count = InsertSharedStringItem(data.base_lvl_score_count.ToString(), sharedString);
                            //int index_upper_base_lvl_score_count = InsertSharedStringItem(data.upper_base_lvl_score_count.ToString(), sharedString);
                            //int index_high_lvl_score_count = InsertSharedStringItem(data.high_lvl_score_count.ToString(), sharedString);
                            //int index_total_score = InsertSharedStringItem(data.total_score.ToString(), sharedString);
                            //int index_base_lvl_count = InsertSharedStringItem(data.base_lvl_count.ToString(), sharedString);
                            //int index_high_lvl_count = InsertSharedStringItem(data.high_lvl_count.ToString(), sharedString);
                            //int min_total_score = InsertSharedStringItem(data.min_total_score.ToString(), sharedString);
                            //int max_total_score = InsertSharedStringItem(data.max_total_score.ToString(), sharedString);

                            Cell cell_region = InsertCellInWorksheet("B", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_schools_count = InsertCellInWorksheet("D", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_clasess_count = InsertCellInWorksheet("E", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_student_count = InsertCellInWorksheet("F", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_student_write_count = InsertCellInWorksheet("H", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_low_lvl_score_count = InsertCellInWorksheet("M", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_upper_low_lvl_score_count = InsertCellInWorksheet("Q", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_base_lvl_score_count = InsertCellInWorksheet("U", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_upper_base_lvl_score_count = InsertCellInWorksheet("Y", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_high_lvl_score_count = InsertCellInWorksheet("AC", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_total_score = InsertCellInWorksheet("AG", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_base_lvl_count = InsertCellInWorksheet("AK", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_high_lvl_count = InsertCellInWorksheet("AN", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_min_total_score = InsertCellInWorksheet("AS", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);
                            Cell cell_max_total_score = InsertCellInWorksheet("AU", (uint)region_data_to_fill.IndexOf(data) + 7, worksheet);

                            cell_region.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.SharedString);
                            cell_schools_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_clasess_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_student_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_student_write_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_low_lvl_score_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_upper_low_lvl_score_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_base_lvl_score_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_upper_base_lvl_score_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_high_lvl_score_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_total_score.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_base_lvl_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_high_lvl_count.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_min_total_score.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                            cell_max_total_score.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);

                            // Set the value of cell A1.
                            cell_region.CellValue = new CellValue(index_region.ToString());
                            cell_schools_count.CellValue = new CellValue(data.schools_count.ToString());
                            cell_clasess_count.CellValue = new CellValue(data.clasess_count.ToString());
                            cell_student_count.CellValue = new CellValue(data.student_count.ToString());
                            cell_student_write_count.CellValue = new CellValue(data.student_write_count.ToString());
                            cell_low_lvl_score_count.CellValue = new CellValue(data.low_lvl_score_count.ToString());
                            cell_upper_low_lvl_score_count.CellValue = new CellValue(data.upper_low_lvl_score_count.ToString());
                            cell_base_lvl_score_count.CellValue = new CellValue(data.base_lvl_score_count.ToString());
                            cell_upper_base_lvl_score_count.CellValue = new CellValue(data.upper_base_lvl_score_count.ToString());
                            cell_high_lvl_score_count.CellValue = new CellValue(data.high_lvl_score_count.ToString());
                            cell_total_score.CellValue = new CellValue(data.total_score.ToString());
                            cell_base_lvl_count.CellValue = new CellValue(data.base_lvl_count.ToString());
                            cell_high_lvl_count.CellValue = new CellValue(data.high_lvl_count.ToString());
                            cell_min_total_score.CellValue = new CellValue(data.min_total_score.ToString());
                            cell_max_total_score.CellValue = new CellValue(data.max_total_score.ToString());                           

                            // Save the new worksheet.
                            worksheet.Save();

                            double progress = (((double)region_data_to_fill.IndexOf(data) + 1) / (double)(region_data_to_fill.Count)) * 100;
                            OnProgress_up((int)progress, sheet_name);
                            progress = 0;
                        }

                        break;
                    }
            }
        }

        private static int InsertSharedStringItem(string text, SharedStringTable sharedString)
        {
            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in sharedString.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            sharedString.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            sharedString.Save();

            return i;
        }

        private Cell InsertCellInWorksheet(string columnName, uint rowIndex, Worksheet worksheet)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                //if (rowIndex > 165)
                //    MessageBox.Show("Смотри");
                return newCell;
            }
        }

        public void ExcelReader(string[] file_paths) // Менеджер для чтения
        {
            if (file_paths.Length == 3)
                foreach (string file_path in file_paths)
                {
                    if (Regex.IsMatch(file_path, @".xlsx$", RegexOptions.IgnoreCase))
                    {
                        if (Regex.IsMatch(file_path, "Матем", RegexOptions.IgnoreCase))
                            ReadExcelFile(file_path, ref records_maths);
                        else if (Regex.IsMatch(file_path, "Рус", RegexOptions.IgnoreCase))
                            ReadExcelFile(file_path, ref records_rus);
                        else if (Regex.IsMatch(file_path, "Чтен", RegexOptions.IgnoreCase))
                            ReadExcelFile(file_path, ref records_reading);
                        else
                            MessageBox.Show("Выбран неверный файл, нажмите \"ОК\", чтобы продолжить", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                        MessageBox.Show("Выбран не Excel файл !", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            else
                MessageBox.Show("Выберете 3 файла для трех предметов, путем выделения трех файлов формата .xlsx", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
       }

        private void ReadExcelFile(string file_path, ref List<Record> records)
        {
            try
            {
                using (SpreadsheetDocument excel_doc = SpreadsheetDocument.Open(file_path, false))
                {
                    Sheet sheet = excel_doc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().FirstOrDefault();
                    if (sheet != null)
                    {
                        SharedStringTable sharedString = excel_doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault().SharedStringTable;
                        Worksheet worksheet = (excel_doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart).Worksheet;

                        List<Row> rows = worksheet.Descendants<Row>().ToList();

                        FindHeadRow(excel_doc, rows, sharedString); // Находим заголовки столбцов и их буквенные индексы, вынесенно в отдельный метод

                        records = new List<Record>(rows.Count);

                        foreach (Row row in rows)
                        {
                            if (row.RowIndex == 1) // Потому что первый мы считываем для заголовков столбов в FindHeadRow
                                continue;

                            temp_values = new List<string>();
                            temp_task_results = new List<Tuple<double, double>>();
                            foreach (Cell cell in row)
                            {
                                if (cell.CellValue != null && !String.IsNullOrWhiteSpace(cell.CellValue.InnerText) && cell.DataType?.Value != CellValues.Error)
                                {
                                    var value = cell.CellValue.InnerText;

                                    if (columnName_columnIndex != null)
                                    {
                                        var temp_columnName_columnIndex = columnName_columnIndex.Where(t => t.Item2 == Regex.Match(cell.CellReference, "[A-Z]{1,}", RegexOptions.IgnoreCase).Value).FirstOrDefault();

                                        if (cell.DataType != null && temp_columnName_columnIndex != null)
                                        {
                                            if (cell.DataType == CellValues.SharedString)
                                            {
                                                if (double.TryParse(temp_columnName_columnIndex.Item1, out double result)
                                                    || Regex.IsMatch(temp_columnName_columnIndex.Item1, "сумма", RegexOptions.IgnoreCase))
                                                {
                                                    if (Regex.IsMatch(temp_columnName_columnIndex.Item1, "сумма", RegexOptions.IgnoreCase))
                                                        temp_task_results.Add(new Tuple<double, double>((int)temp_task_results[temp_task_results.Count - 1].Item1 + 1,
                                                            double.Parse(sharedString.ElementAt(int.Parse(value)).InnerText, CultureInfo.InvariantCulture.NumberFormat)));
                                                    else
                                                        temp_task_results.Add(new Tuple<double, double>(result, double.Parse(cell.CellValue.InnerText, CultureInfo.InvariantCulture.NumberFormat)));
                                                }
                                                else
                                                    temp_values.Add(sharedString.ElementAt(int.Parse(value)).InnerText);
                                            }
                                        }
                                        else if (temp_columnName_columnIndex != null)
                                        {
                                            if (double.TryParse(temp_columnName_columnIndex.Item1, NumberStyles.Any, CultureInfo.InvariantCulture.NumberFormat, out double result)
                                                || Regex.IsMatch(temp_columnName_columnIndex.Item1, "сумма", RegexOptions.IgnoreCase))
                                            {
                                                if (Regex.IsMatch(temp_columnName_columnIndex.Item1, "сумма", RegexOptions.IgnoreCase))
                                                    temp_task_results.Add(new Tuple<double, double>((int)temp_task_results[temp_task_results.Count - 1].Item1 + 1,
                                                        double.Parse(cell.CellValue.InnerText, CultureInfo.InvariantCulture.NumberFormat)));
                                                else
                                                    temp_task_results.Add(new Tuple<double, double>(result, double.Parse(cell.CellValue.InnerText, CultureInfo.InvariantCulture.NumberFormat)));
                                            }
                                            else
                                                temp_values.Add(cell.CellValue.InnerText);
                                        }
                                    }
                                    else
                                        MessageBox.Show("Список заголовков пуст (Этап считывания)", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            records.Add(new Record(temp_values, temp_task_results));

                            double progress = ((double)records.Count / (double)(rows.Count - 1)) * 100;
                            OnProgress_up((int)progress, file_path);
                            progress = 0;
                        }
                    }
                }
                if (columnName_columnIndex != null)
                    columnName_columnIndex.Clear();
            }
            catch
            {
                MessageBox.Show($"Файл {Path.GetFileName(file_path)} уже открыт в другой программе.\nПожалуйста, закройте его и начните процедуру считывания заново.",
                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }            
        }

        private void FindHeadRow(SpreadsheetDocument doc, IEnumerable<Row> rows, SharedStringTable sharedStringTable) // Считываем заголовки столбцов и их буквенные инедксы
        {
            Row head_row = rows.Where(r => r.RowIndex == 1).FirstOrDefault();

            if (columnName_columnIndex == null)
                columnName_columnIndex = new List<Tuple<string, string>>();

            if (columnName_columnIndex != null)
            {
                foreach (Cell cell in head_row)
                {
                    if (cell.CellValue != null && !String.IsNullOrWhiteSpace(cell.CellValue.InnerText) && cell.DataType?.Value != CellValues.Error)
                    {
                        if (cell.DataType != null)
                        {
                            if (cell.DataType == CellValues.SharedString)
                            {
                                columnName_columnIndex.Add(new Tuple<string, string>(sharedStringTable.ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText,
                                      Regex.Match(cell.CellReference, "[A-Z]{1,}").Value));
                            }
                        }
                        else
                        {
                            columnName_columnIndex.Add(new Tuple<string, string>(cell.CellValue.InnerText,
                                Regex.Match(cell.CellReference, "[A-Z]{1,}").Value));
                        }
                    }
                }
                RemoveNeedless();
            }
        }

        private void RemoveNeedless() // Скорее всего здесь придется подкручивать, если количество подзаданий изменится. Как-то это нужно решить.
        {
            List<Tuple<string, string>> remove_tuple = new List<Tuple<string, string>>(); // Темповый список тьюплов, которые нужно удалить из columnName_columnIndex
            if (columnName_columnIndex != null)
            {
                foreach (Tuple<string, string> remove_it in columnName_columnIndex) // Ищем то, что нужно удалить
                {
                    if (Regex.IsMatch(remove_it.Item1, "сумма", RegexOptions.IgnoreCase) && columnName_columnIndex.IndexOf(remove_it) >= 2)
                    {
                        remove_tuple.AddRange(new List<Tuple<string, string>> {columnName_columnIndex[columnName_columnIndex.IndexOf(remove_it) - 1],
                            columnName_columnIndex[columnName_columnIndex.IndexOf(remove_it) - 2] });
                    }
                }
                foreach (Tuple<string, string> remove_it_now in remove_tuple) // Удаляем необходимое
                {
                    columnName_columnIndex.Remove(remove_it_now);
                }
            }
        }

        public bool ChekReadyToWrite()
        {
            if (records_maths != null && records_rus != null && records_reading != null)
                return true;
            else
                return false;
        }
    }
}
