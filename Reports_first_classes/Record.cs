using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Windows.Forms;
using System.IO;

namespace Reports_first_classes
{
    class Record
    {
        public string main_id { get; private set; } // Код класса
        public string school_id { get; private set; } // Код класса
        public string class_id { get; private set; } // Код класса
        public string school_name { get; private set; } // Наименование школы
        public string class_name { get; private set; } // Класс
        public string student_FIO { get; private set; } // ФИО ученика
        public string region { get; private set; } // Регион 
        public string variant { get; private set; }
        public List<Tuple<double, double>> task_results { get; private set; } // Номер задания и количество баллов, полученных за задание
        public double total_score { get; private set; } // Сумма баллов за все задания
        public double t_score { get; private set; } // t-баллы
        public string understand_lvl { get; private set; } // Уровень подготовки
        public double base_lvl { get; private set; } // Базовый уровень
        public double high_lvl { get; private set; } // Повышенный уровень

        public Record(List<string> data, List<Tuple<double, double>> task_results) //TODO: Реализовать работу со списком task_results. Добавить второй входной лист данных для номера задания и результата.
        {
            if (data.Count == 9)
            {
                data.Insert(4, "None");
                data.Insert(5, "None");
            }
            else if (data.Count == 10)
            {
                if (Regex.IsMatch(data[4], @"\d"))
                {
                    data.Insert(4, "None");
                }
                else if (Regex.IsMatch(data[4], @"\w+", RegexOptions.IgnoreCase))
                {
                    data.Insert(5, "None");
                }
            }

            if (data.Count == 11)
            {
                main_id = data[0];
                school_id = data[0].Substring(4, 6); 
                class_id = data[0].Substring(10, 4); 
                school_name = data[1];
                class_name = Regex.IsMatch(data[2], @"\d$") ? Regex.IsMatch(data[2], @"[0-9]-[0-9]$") ? data[2]
                    : data[2].Length > 3 ? data[2].Remove(3) : data[2] : data[2];
                student_FIO = data[3];
                region = data[4] == "None"? school_name : data[4];
                variant = Regex.IsMatch(data[5], "N", RegexOptions.IgnoreCase) ? "N" : data[5];
                total_score = double.Parse(data[6], CultureInfo.InvariantCulture.NumberFormat);
                t_score = Regex.IsMatch(data[7], "N", RegexOptions.IgnoreCase) ? 0 : double.Parse(data[7], CultureInfo.InvariantCulture.NumberFormat);
                understand_lvl = data[8];
                base_lvl = double.Parse(data[9], CultureInfo.InvariantCulture.NumberFormat);
                high_lvl = double.Parse(data[10], CultureInfo.InvariantCulture.NumberFormat);
            }
            else
            {
                MessageBox.Show($" Не совпадает количество данных");
            }

            this.task_results = task_results;
        }

        public void ShowInfo(int record_number, string file_path) // Для теста
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"\nЗапись № {record_number + 2} в Excel файле \"{Path.GetFileName(file_path)}\"");
            Console.ForegroundColor = ConsoleColor.Black;
            Console.WriteLine($"Код класса: {class_id}\n" +
                $"Наименование школы: {school_name}\n" +
                $"Наименование класса: {class_name}\n" +
                $"ФИО студента: {student_FIO}\n" +
                $"Регион: {region}\n" +
                $"Вариант: {variant}\n" +
                $"Итоговый результат: {total_score}\n" +
                $"t-балл: {t_score}\n" +
                $"Уровень усвоения: {understand_lvl}\n" +
                $"Базовый уровень: {base_lvl}\n" +
                $"Повышенный уровень: {high_lvl}");
            if (task_results != null)
            {
                Console.WriteLine("Задания и результаты: ");
                foreach (Tuple<double, double> res in task_results)
                    Console.WriteLine($"{res.Item1}) {res.Item2}");
            }
            else
                Console.WriteLine("Список заданий пуст");
        }
    }
}
