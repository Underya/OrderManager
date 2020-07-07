using System;
using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;

namespace FromExel
{
    /// <summary>
    /// Открытый EXEL файл с функциями работы с ним
    /// </summary>
    class ExelBook
    {
        /// <summary>
        /// Идентификатор открытой книги
        /// </summary>
        int id;

        /// <summary>
        /// Текущая выбранная книга
        /// </summary>
        _Workbook book = null;

        /// <summary>
        /// Текущий активный лист
        /// </summary>
        _Worksheet activ = null;

        /// <summary>
        /// Создание новокого класс для работы с книжкой
        /// </summary>
        /// <param name="_book">Указатель на открытую книгу</param>
        /// <param name="_id">id новой книжки</param>
        public ExelBook(_Workbook _book, int _id)
        {
            //сохранение парамаетров
            book = _book;
            id = _id;

        }

        /// <summary>
        /// Закрытие книжки с сохраннеием данных
        /// </summary>
        public void CloseSave()
        {
            //Если книжка уже закрыта
            if (book == null) return;
            // Сохранение данных в книжке
            book.Save();
            //Закрытие книги
            book.Close(false);
            //Удаление ссылки на книжку
            book = null;
            // Вызов функции для закрытия книжки из списка открытых
            ExelManager.CloseBooks(id);
        }

        /// <summary>
        /// Закрытия книжки без сохранения данных
        /// </summary>
        public void CloseNotsave()
        {
            //Если книжка уже закрыта
            if (book == null) return;
            //Закрытие книжки
            book.Close(false);
            //Удаление ссылку на книжку
            book = null;
            //Вызов метода для удаления из списка открытых
            ExelManager.CloseBooks(id);
        }

        /// <summary>
        /// Установка активного листа
        /// </summary>
        /// <param name="NameSheet">Имя листа</param>
        public void SetActiveSheet(string NameSheet)
        {
            activ = book.Worksheets[NameSheet];
        }

        /// <summary>
        /// Установка активного листа
        /// </summary>
        /// <param name="Number">Номер листа</param>
        public void SetActiveSheet(int Number)
        {
            activ = book.Worksheets[Number];
        }

        /// <summary>
        /// Запись текста в укзанную ячейку для активного листа
        /// </summary>
        /// <param name="row">Строка, в которую надо записать</param>
        /// <param name="column">Столбец, в который надо записать</param>
        /// <param name="str">Текст, который будет записан в ячейку</param>
        public void Write(int row, string column, string str)
        {
            //Если не выбран лист, выбросить ошибку
            if(activ == null)
            {
                throw new ExceptionError(40, "Не выбран активный лист", "Location:ExelBook");
            }

            //Запись в ячейку
            activ.Cells[row, column] = str;
        }

        /// <summary>
        /// Чтение из указаной ячейки в активном листе
        /// </summary>
        /// <param name="row">Строка ячейки</param>
        /// <param name="column">Столбец ячейки</param>
        /// <returns></returns>
        public string Read(int row, string column)
        {
            //Если не выбран лист, выбросить ошибку
            if (activ == null) throw new ExceptionError(40, "Не выбран активный лист", "Location:ExelBook");
            //Получение текста из ячейки
            return activ.Cells[row, column].Text.ToString();
        }

        /// <summary>
        /// Чтение числом текста из указнной ячейки в активном листе
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        public int ReadNum(int row, string column)
        {
            //Если не выбран лист, выбросить ошибку
            if (activ == null) throw new ExceptionError(40, "Не выбран активный лист", "Location:ExelBook");
            //Получение числа из указной ячейки
            return int.Parse(activ.Cells[row, column].Text.ToString());
        }

        /// <summary>
        /// Сохранение книжки по новому имени
        /// </summary>
        /// <param name="_path">Полное имя новой книжки</param>
        public void SaveAs(string _path)
        {
            //Если книжка уже сохранена
            if (book == null) return;
            book.SaveAs(_path);
            //Закрытие книжки
            book.Close(false);
            //Удаление ссылки на книжку
            book = null;
            //Удаление из спика открытых
            ExelManager.CloseBooks(id);
        }

        /// <summary>
        /// Число листов в книге
        /// </summary>
        public int SheetCount
        {
            get { return book.Worksheets.Count; }
        }

        /// <summary>
        /// Установка ширина столба
        /// </summary>
        /// <param name="width">Новая ширина столбца</param>
        /// <param name="number">Номер столбца</param>
        public void SetColumnWidth(int number, int width)
        {
            //Установка ширины столбца
            activ.Columns[number].ColumnWidth = width;
        }

        /// <summary>
        /// Добавление нового листа в конец списка
        /// </summary>
        /// <returns>Число листов в книге вместе с новым</returns>
        public int addSheetLast()
        {
            //Получение листа
            _Worksheet last = book.Worksheets[book.Worksheets.Count];
            //Добавление нового листа после старого последнего
            book.Worksheets.Add(After: last);
            //Возвращение нового числа листов
            return book.Worksheets.Count;
        }

        /// <summary>
        /// Установка имени активного листа
        /// </summary>
        /// <param name="name"></param>
        public void SetNameSheet(string name)
        {
            activ.Name = (SheetCount).ToString();
        }

        /// <summary>
        /// Деструктор, на случай, если элемент не был закрыт
        /// </summary>
        ~ExelBook()
        {
            //Выход без сохранения
            CloseNotsave();
        }
    }
}
