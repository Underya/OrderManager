using System;
using System.Windows.Forms;
using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;

namespace FromExel.Resourse
{
    /// <summary>
    /// Все используемые форматы
    /// </summary>
    class Format
    {
        /// <summary>
        /// Массив всех форматов, с которыми происходит работа
        /// </summary>
        static string[] allFormat = null;

        /// <summary>
        /// Загрузка всех форм 
        /// </summary>
        /// <param name="path"></param>
        public static void inital(string path)
        {
            //Получение книги со всеми форматами
            ExelBook b = ExelManager.getBook(path);
            //Выбор листа с форматами
            b.SetActiveSheet("1");
            //Создание вспогоательного списка для его превращения в массиве
            List<string> list = new List<string>();
            //Чтение из файла всех строк до пустой
            for(int i = 1; i < 2500; i++)
            {
                //Чтение текста формата из ячейки
                string s = b.Read(i, "A");
                //если в ячейки не текста, конец цикла
                if (s == "") break;
                //Добавлегние формата в список
                list.Add(s);
            }
            //Закрытие книги
            b.CloseNotsave();
            //Создание массива с форматами
            allFormat = new string[list.Count];
            //Перенос всех форматов в массив
            for(int i = 0, count = list.Count; i < count; i++)
            {
                allFormat[i] = list[i];
            }
        }

        /// <summary>
        /// Добавлеение в ячейка комбо бокса всех возможных форматов
        /// </summary>
        /// <param name="cell">Ячейка, в которую происходит добавление</param>
        public static void SetFormatToCell(DataGridViewComboBoxCell cell)
        {
            //Перенос форматов в ячейку таблицы

            foreach(string s in allFormat)
            {
                cell.Items.Add(s);
                cell.Value = "";
            }
        }
    }
}
