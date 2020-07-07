using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

using Microsoft.Office.Interop.Excel;

namespace FromExel
{
    public partial class OrderScoreFormcs : Form
    {
        /// <summary>
        /// Информация о еденичной работе в заказе
        /// </summary>
        protected struct workInf
        {
            /// <summary>
            /// Номер работы в оригинальном листе
            /// </summary>
            public string Num;

            /// <summary>
            /// Наименование из оригинального листа
            /// </summary>
            public string Name;

            /// <summary>
            /// Формат работы: А4, А3 плотная, ламининирование и т.д.
            /// </summary>
            public string Form;

            /// <summary>
            /// Количество элементов в одном экземпляре
            /// </summary>
            public string Exem;

            /// <summary>
            /// Количество экземпляров
            /// </summary>
            public string Count;

            /// <summary>
            /// Суммарно элементов за работу
            /// </summary>
            public string Summ;
        }


        /// <summary>
        /// структура с информацией о заказе
        /// </summary>
        protected struct zakazInf
        {
            /// <summary>
            /// Номер заказа
            /// </summary>
            public string Number;

            /// <summary>
            /// Дата заказа
            /// </summary>
            public string Data;

            /// <summary>
            /// Имя заказчика
            /// </summary>
            public string Name;

            /// <summary>
            /// Номер листа
            /// </summary>
            public int ListNum;

            /// <summary>
            /// Номер строки
            /// </summary>
            public int NumberStr;

            /// <summary>
            /// Список работ в одном заказе
            /// </summary>
            public List<workInf> works;
        }

        /// <summary>
        /// Список всех не подписанных заказов
        /// </summary>
        List<zakazInf> list;

        /// <summary>
        /// Exel документов с заказами
        /// </summary>
        ExelBook book = null;

        /// <summary>
        /// Путь к файлу
        /// </summary>
        string _path;

        /// <summary>
        /// Коллекция файлов-заказов
        /// Используется, что бы только один раз получить всю коллекцию и искать в ней
        /// </summary>
        IEnumerable<string> filesWork = null;

        public OrderScoreFormcs()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Чтение файла с заказом и создание информации о нём
        /// </summary>
        /// <param name="book"></param>
        /// <param name="list"></param>
        private void ReadOrderFile(ExelBook book, List<workInf> list)
        {
            //Строки и столбец первой ячейки с заказом
            int row = 9;
            string columnNum = "B";
            string columnWorkName = "C";
            string columnFormat = "E";
            string columnCount = "G";
            string columnExemp = "H";
            string columnSummCount = "J";

            book.SetActiveSheet(1);

            //Циклический проход по всем работам в листе
            while (true)
            {                
                //Полученеи имени работы
                string Name = book.Read(row, columnWorkName);
                //Если нет имени, список работы закончен
                if (Name == "") break;

                workInf currentInf = new workInf();

                //Получение информации о заказе из всех остальных столбцов
                currentInf.Num = book.Read(row, columnNum);
                currentInf.Name = Name;
                currentInf.Form = book.Read(row, columnFormat);
                currentInf.Count = book.Read(row, columnCount);
                currentInf.Exem = book.Read(row, columnExemp);
                currentInf.Summ = book.Read(row, columnSummCount);

                //Добавление конркетной работы в список
                list.Add(currentInf);

                //Переход к другой строке
                row++;

            }

            
        }
        
        /// <summary>
        /// Получить текущую выбранную строку таблицки
        /// </summary>
        /// <returns></returns>
        private DataGridViewRow getSelectRow(DataGridView Table)
        {
            //Цикл по всем строкам
            foreach (DataGridViewRow CurrRow in Table.Rows)
            {
                //Если строка - выбранна
                if (CurrRow.Selected)
                {
                    //Возвращение строки
                    return CurrRow;
                }
            }

            //Если не нашлось строки выбранной - вернуть первую
            return Table.Rows[0];
        }

        /// <summary>
        /// Метод обновляет информацию о заказе в правой половине столбца
        /// </summary>
        private void UpdateInfWork()
        {
            //Очищение таблицы с информацией о работе 
            dataGridView2.Rows.Clear();
            //Получение выбранной строки
            DataGridViewRow row = getSelectRow(dataGridView1);

            //Получение информации о заказе
            zakazInf zakaz = (zakazInf)(row.Cells[5].Value);

            //Получение содержимого заказа
            List<workInf> works = zakaz.works;

            //Если заказ - пустой, то конец работы

            if (works == null) return;

            //Просмотр всех позиций по работе
            foreach (workInf work in works)
            {
                //Создание новой строки
                DataGridViewRow newRow = new DataGridViewRow();
                //Добавление столбцов по шаблону
                newRow.CreateCells(dataGridView2);
                DataGridViewCellCollection newCell = newRow.Cells;

                //Заполнение таблиц
                newCell[0].Value = work.Num;
                newCell[1].Value = work.Name;
                newCell[2].Value = work.Form;
                newCell[3].Value = work.Count;
                newCell[4].Value = work.Exem;
                newCell[5].Value = work.Summ;

                //Добавление строки в таблицу
                dataGridView2.Rows.Add(newRow);
            }
        }

        /// <summary>
        /// Инициализация формы
        /// </summary>
        /// <param name="Path_histor">Путь к файлу с записями</param>
        /// <param name="Path_orders">Путь к папке с заказами</param>
        public void initial(string Path_histor, string Path_orders)
        {
            _path = Path_histor;
            //Создание нового списка заказов
            list = new List<zakazInf>();
            //Октрытие EXEL программы
            //Получение книги
            book = ExelManager.getBook(_path);
            
            

            //Цикл по всем листам файлам, в поисках не закрытых заказов
            for(int NumberSheet = 1; NumberSheet <= book.SheetCount; NumberSheet++)
            {
                //Установка активного листа
                book.SetActiveSheet(NumberSheet);

                //Проход по всем не заполненым я листе ячейкам
                for(int i = 1; i < 32000; i++)
                {
                    //Проверка ячейки на пустоту
                    if(book.Read(i, "D") != null)
                    {
                        //Если ячейка пустая - конец цикла для данного листа
                        if (book.Read(i, "D") == "") break;
                        
                        //Не выполненые заказа необходимо запомнить и вывести
                        if (book.Read(i, "D") == "0")
                        {
                            //Формирование нового элемента
                            zakazInf z = new zakazInf();
                            z.Number = book.Read(i, "A");
                            z.Data = book.Read(i, "B");
                            z.Name = book.Read(i, "C");
                            z.ListNum = NumberSheet;
                            z.NumberStr = i;
                            
                            //После того, как в список добавили заказ, происходит поиск файла с тем же номером
                            ExelBook b = getFile(Path_orders, z.Number);

                            //Проверка, удалось ли найти файл с таким навзанием
                            if (b != null)
                            {
                                //Создание массива с работами в заказе
                                z.works = new List<workInf>();

                                //Чтение файла
                                ReadOrderFile(b, z.works);

                                //Окончаение работы с открытием файла
                                b.CloseNotsave();
                            }

                            //Добавление элемента в список
                            list.Add(z);
                        }

                    }
                        
                }
            }

            //Добавление всех элементов из списка в таблицу для пользователя
            //Указатель на элемент в списке
            int index = 0;

            foreach(zakazInf z in list)
            {
                int rownum = dataGridView1.Rows.Add();
                //Получение новой строки
                DataGridViewRow row = dataGridView1.Rows[rownum];
                row.Cells[0].Value = z.Number;
                row.Cells[1].Value = z.Data;
                row.Cells[2].Value = z.Name;
                row.Cells[3].Value = "Подписать";
                row.Cells[4].Value = index;
                //В данной ячейки хранится объект с информацией о заказе
                row.Cells[5].Value = z;
                index++;
            }

            //Обновление правой таблицы
            UpdateInfWork();
        }

        /// <summary>
        /// Функция возвращает уже созданный файл заказа в формате ExelBook 
        /// или Null, если файл не найден
        /// </summary>
        /// <param name="path_to_orders">Путь к папке с файлами заказа</param>
        /// <param name="OrderNum">Номер заказа, который надо найти</param>
        /// <returns>Открытая книга с файлом заказа с тем же номером, или null </returns>
        private ExelBook getFile(string path_to_orders, string OrderNum)
        {
            //Создание укзателя для книги
            ExelBook book = null;

            //Если не было раньше получено содержимое папки ранее 
            if (filesWork == null)
                //Получение всего содержимого папки с файлами заказов
                filesWork = Directory.EnumerateFiles(path_to_orders);
            //Path.GetFileName()

            //Поиск среди всех файлов
            foreach(string file in filesWork)
            {
                //Проверка, входит ли номер в имя файла
                if (file.Contains(OrderNum))
                {
                    //Если входит, то из названия файла извлекается номер в формате строки
                    //Удаление всех символов перед последним символов \ и самого символа
                    string t = file.Substring(file.LastIndexOf('\\') + 1);
                    //Получение номера из стоки
                    string num_file = t.Substring(0, t.IndexOf(' '));
                    //Если номера совпадают, то открывается книга
                    if(OrderNum == num_file)
                    {
                        //Создание новой книги
                        book = ExelManager.getBook(file);
                        //Конец работы цикла
                        return book;
                    }
                }
                
            }

            return book;
        }

        /// <summary>
        /// Обработка действия при нажатии на кнопку выполнить в ячейке таблицы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //Если книжка была закрыта, открытие заново
            if (book == null) book = ExelManager.getBook(_path);
            //Если была нажата клавиши действия
            if(e.ColumnIndex == 3 && e.RowIndex >= 0)
            {
                //Получение номер элемента списка из строки
                zakazInf z = list[(int)dataGridView1.Rows[e.RowIndex].Cells[4].Value];
                //Изменение записи в EXEL
                book.SetActiveSheet(z.ListNum);
                book.Write(z.NumberStr, "D", "1");
                //Удаление строчки из таблицы
                dataGridView1.Rows.RemoveAt(e.RowIndex);
            }
        }

        /// <summary>
        /// При закрытии фаормы закрыть EXEL
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OrderScoreFormcs_FormClosing(object sender, FormClosingEventArgs e)
        {
            //Если книжка сохранена, на неё нет ссылки
            if (book == null) return;
            //Закрыть книжку с сохранением
            book.CloseSave();
        }

        private void просмотретьВсеToolStripMenuItem_Click(object sender, EventArgs e)
        {

            //Сохранить изменнеия в книжки, перед тем как покзать её пользователю
            if(book != null) book.CloseSave();
            //Удалить ссылку на книгу
            book = null;
            //Через проводник открыть файл
            Process.Start(_path);
        }

        /// <summary>
        /// Метод, который вызывается при нажатии на клавишу
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            UpdateInfWork();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            UpdateInfWork();
        }
    }
}
