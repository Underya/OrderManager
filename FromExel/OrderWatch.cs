using System;
using Microsoft.Office.Interop.Excel;
using System.IO;


namespace FromExel
{
    /// <summary>
    /// Существует файл, в котором хранятся краткая информация о заказах
    /// В виде "Номер, дата, имя"
    /// Данный класс предоставляет метод для записи в этот файл новых записей о заказах
    /// </summary>
    class OrderWatch
    {
        /// <summary>
        /// Путь к EXEL файлу
        /// </summary>
        string _path;

        /// <summary>
        /// Создание нового класса для работы с файлами
        /// Класс привязывается к файлу с данными о зазах
        /// </summary>
        /// <param name="path">Пусть к exel файлу для отслеживания</param>
        public OrderWatch(string path)
        {
            //Проверка, существует ли файл
            if (!File.Exists(path)) throw new ExceptionError(102, "Нет файла с заказами", "Location:OrdeWatch");
            //Сохранени пути
            _path = path;
        }

        /// <summary>
        /// Добавление в файл записи о новой работе
        /// </summary>
        /// <param name="number">Номер заказа</param>
        /// <param name="filename">Имя заказа</param>
        /// <param name="Exists">Подисани ли работа отвественм или нет</param>
        public void addOrder(int number, string filename, bool ExistsSignature = false) 
        {
            try
            {
                //Получение книги
                ExelBook book = ExelManager.getBook(_path);
                //Номер последнего листа
                int SheetCount = book.SheetCount;
                //Установка активным последнего листа
                book.SetActiveSheet(SheetCount);

                //Поиск свободного места в файла
                //Индекс строки
                int i = 1;
                //В цикле происходит поиск свободной ячейки
                while (true)
                {
                    
                    if (book.Read(i, "A") == "") break;
                    i++;
                }

                //Если нашли строку, заполнение
                book.Write(i, "A",  number.ToString());
                book.Write(i, "B", DateTime.Now.ToShortDateString());
                book.Write(i, "C", filename);
                if (ExistsSignature) book.Write(i, "D", "1"); else book.Write(i, "D", "0");

                //Если один лист заполнен
                if (i >= 50)
                {
                    //Добавление нового листа в конец
                    book.addSheetLast();
                    //Добавление нового листа
                    //Установка листа активным
                    book.SetActiveSheet(book.SheetCount);
                    book.SetColumnWidth(2, 20);
                    book.SetColumnWidth(3, 20);
                    
                }

                //Закрытие книги
                book.CloseSave();
            }
            catch (Exception error)
            {
                //Выбросить ошибку дальше
                throw error;
            }

        }
  
    }
}
