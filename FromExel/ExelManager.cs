using System;
using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;

namespace FromExel
{
    /// <summary>
    /// Класс, управляющий работой Exel приложения
    /// </summary>
    static class ExelManager
    {
        /// <summary>
        /// Указатель на приложение Exel - Если указатель = null - приложение закрыто
        /// </summary>
        static Application ExelApp = null;

        /// <summary>
        /// Открытые в приложении книги
        /// </summary>
        static Dictionary<int, ExelBook> OpenBooks;

        /// <summary>
        /// Идентификатор книги
        /// </summary>
        static int idBook = 0;

        /// <summary>
        /// Конструктор для класса
        /// </summary>
        static ExelManager()
        {
            //Создание нового массива с книгами
            OpenBooks = new Dictionary<int, ExelBook>();

        }

        /// <summary>
        /// Открытие EXEL приложения
        /// </summary>
        static public void Start()
        {
            ExelApp = new Application();
            //ExelApp.Visible = true;
            ExelApp.DisplayAlerts = false;
        }

        /// <summary>
        /// Закрыть EXEL прилоежние
        /// </summary>
        static public void Stop()
        {
            //Проверка, все ли книги закрыты
            if (OpenBooks.Count > 0) throw new ExceptionError(42, "Попытка закрыть приложение при открытой книге", "Location:ExManger");
            //Иначе - выход из приложения
            ExelApp.Quit();
            ExelApp = null;
        }
        
        /// <summary>
        /// При закрытии программы проверка, выключен ли EXEL
        /// </summary>
        static public void Quit()
        {
            //Если есть открыте книги
            if(OpenBooks.Count > 0)
            {
                //Получение ключей всех открытых книг
                Dictionary<int, ExelBook>.KeyCollection keyColl =  OpenBooks.Keys;
                //Цикл по всем ключам в коллекции
                foreach (int i in keyColl)
                {
                    //Закрытие книги
                    OpenBooks[i].CloseNotsave();
                }
            }

            //Если не закрыто приложение, закрытие прилоежния
            if (ExelApp != null) ExelApp.Quit();
            //Удаление ссылки на приложение
            ExelApp = null;
        }

        /// <summary>
        /// Получение книги по пути к ней
        /// </summary>
        /// <param name="path">путь к книге</param>
        /// <returns></returns>
        public static ExelBook getBook(string path)
        {
            //Создание новой книжки
            ExelBook book = new ExelBook(ExelApp.Workbooks.Open(path), idBook);
            //Добавление новой книиги в словарь
            OpenBooks.Add(idBook++, book);
            //Возвращение указателя на книгу
            return book;
        }

        /// <summary>
        /// Метод, который вызывается, когда закрывается книга
        /// </summary>
        /// <param name="idBook">id закрываемой книги</param>
        public static void CloseBooks(int idBook)
        {
            //Исключение книги из списка открытых
            //Закрытие книжки, если она ещё открыта
            ExelBook book = OpenBooks[idBook];
            book.CloseNotsave();
            //Удаление книжки из списка
            OpenBooks.Remove(idBook);
            book = null;
        }
    }
}
