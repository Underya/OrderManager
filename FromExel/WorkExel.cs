using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;


namespace FromExel
{
    /// <summary>
    /// Функции для работы с exel 
    /// </summary>
    class WorkExel
    {
        /// <summary>
        /// Полное имя Exel файла шаблона заказа
        /// </summary>
        string TempName = "";

        /// <summary>
        /// Полное имя Exel файла с атрибутами
        /// </summary>
        string AttName = "";

        /// <summary>
        /// Имя текущего открытого EXEL файла
        /// </summary>
        protected string CurrName = "Без имени";

        /// <summary>
        /// Путь, куда сохраняется файл
        /// </summary>
        string SavePath = "";

        /// <summary>
        /// Номер заказа
        /// </summary>
        int CurrNumber = 0;

        /// <summary>
        /// Книга для записи
        /// </summary>
        ExelBook book;

        /// <summary>
        /// Имя последнего файла
        /// </summary>
        public string LastFileName;

        /// <summary>
        /// Номер последнего сохранённого файла
        /// </summary>
        public int LastFileNumber;

        /// <summary>
        /// Создание файла для работы с EXEL документами с параметрами из 2 файлов
        /// </summary>
        /// <param name="_TempName">Полное имя EXEL файла шаблона заказа</param>
        /// <param name="_AttName">Полное имя Exel файла аттрибутов заказа</param>
        protected WorkExel(string _TempName, string _AttName, string _SavePath)
        {
            //Сохранение параметров
            TempName = _TempName;
            AttName = _AttName;
            SavePath = _SavePath;
            //Проверка, существуют ли указанные файлы
            try
            {

                if (!System.IO.File.Exists(TempName) || !System.IO.File.Exists(AttName)) throw new Exception();

            } catch
            {
                throw new ExceptionError(101, "Нет файла шаблона", "Class:WorkExel;");

            }

        }

        /// <summary>
        /// Создание нового файла заказа  по щаблону, и установка его как активного
        /// </summary>
        protected void OpenNewFile()
        {
            //Получение номера из файла аттриубтов
            
            //Открытия книжки шаблона
            book = ExelManager.getBook(AttName);
            //Установка активного листа
            book.SetActiveSheet("0");
            //Чтение номера
            CurrNumber = int.Parse(read(1, "A"));
            //Запись нового номера
            write(1, "A", (++CurrNumber).ToString());
            //Сохранение нового номера как текущего
            LastFileNumber = CurrNumber;
            //Сохранить изменения и закрыть книжку
            book.CloseSave();
            //Откыть книгу с шаблоном для заказа
            book = ExelManager.getBook(TempName);
            //Установка активной книги
            book.SetActiveSheet("0");
        }

        /// <summary>
        /// Запись в указаную ячейку в активном листе
        /// </summary>
        /// <param name="num">номер строки</param>
        /// <param name="column">номер столбца</param>
        /// <param name="value">значение</param>
        protected void write(int num, string column, string value)
        {
            if (book == null) throw new ExceptionError(40, "Не выбран активный лист", "WorkExel");
            book.Write(num, column, value);
        }

        /// <summary>
        /// чтение из указаной ячейки в активном листе
        /// </summary>
        /// <param name="num">номер строки</param>
        /// <param name="column">номер столбца</param>
        protected string read(int num, string column)
        {
            if (book == null) throw new ExceptionError(40, "Не выбран активный лист", "WorkExel");
            return book.Read(num, column);
        }

        /// <summary>
        /// Сохранение заказ по указному в классе пути и имени
        /// </summary>
        protected void saveOrder()
        {
            if (book == null) throw new ExceptionError(41, "Не выбрана активная книга", "workExel:saveOrder");
            LastFileName = CurrName;
            book.SaveAs(SavePath + LastFileNumber.ToString() + ' ' + LastFileName + ".xlsx");
            //Закрытие книжки без сохранения
            book = null;
        }

        /// <summary>
        /// Пример функций и ТД
        /// </summary>
        void PrimerRaboty()
        {
            //Открытия приложения EXEL
            var ExelApp = new Application();
            //Для теста, вывод приложения
            ExelApp.Visible = true;
            //Открытия книжки шаблона
            var tempBook = ExelApp.Workbooks.Open(@"C:\project\FromExel\FromExel\t.xlsx");
            //Введение измнений, допустим
            //получение 1 листа
            _Worksheet sheet = tempBook.Worksheets.Item["0"];
            sheet.Cells[1, "A"] = "adwdawda";
            tempBook.SaveAs(@"C:\project\FromExel\FromExel\t2.xlsx");
            tempBook.Close(false);
            ExelApp.Quit();
            
        }

        /// <summary>
        /// Закрыть работу без сохранения
        /// </summary>
        protected void quitNoSave()
        {
            //Если есть книжка
            if(book != null)
            {
                book.CloseNotsave();
            }
        }
    }
}
