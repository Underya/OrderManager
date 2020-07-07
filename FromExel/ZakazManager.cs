using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FromExel
{
    /// <summary>
    /// Класс для работы с заказами через Exel
    /// </summary>
    class ZakazManager : WorkExel, saveData
    {
        /// <summary>
        /// Количество мест для работы в шаблоне
        /// </summary>
        int CountWork = -1;

        /// <summary>
        /// Смещение по столбцу работы
        /// </summary>
        int currWork = 0;

        /// <summary>
        /// Флаг, существует ли в данный момент активная книжка
        /// </summary>
        bool initial = false;

        /// <summary>
        /// Создание нового экземпляра по файлу с параметрами
        /// </summary>
        /// <param name="conf"></param>
        public ZakazManager(SetupConfig conf) : base(conf["tempname"], conf["attname"], conf["savepath"])
        {
            //Получение числа работ в шаблоне
            CountWork = int.Parse(conf["countwork"]);
        }

        /// <summary>
        /// Инициалзиция класса для работы с заказом
        /// </summary>
        void saveData.StartWork()
        {
            OpenNewFile();
            initial = true;
        }

        /// <summary>
        /// Конец работы и сохранения заказа
        /// </summary>
        void saveData.EndWork()
        {
            saveOrder();
            initial = false;
            //Обнуление смещения
            currWork = 0;
        }

        /// <summary>
        /// Добавление одной работы в заказ
        /// </summary>
        /// <param name="NameWork"></param>
        /// <param name="format"></param>
        /// <param name="countSheet"></param>
        /// <param name="countExem"></param>
        /// <param name="Sum"></param>
        void saveData.AddWork(string NameWork, string format, int countSheet, int countExem, int Sum)
        {
            if (initial == false) throw new ExceptionError(21, "Не инициализирован класс для работы с заказами", "Location:ExelManager");

            //Проверка, хватило ли места для записи работ
            if (currWork == CountWork) throw new ExceptionError(22, "Нет места в таблице заказа", Location:"Location:ExelManegr");

            //Запись данных в указаную ячейку
            int currStr = 9 + currWork;

            //Запись данных в таблицу
            write(currStr, "C", NameWork);
            write(currStr, "E", format);
            if(countSheet != -1) write(currStr, "G", countSheet.ToString());
            if (countExem != -1) write(currStr, "H", countExem.ToString());
            if (Sum != -1) write(currStr, "J", Sum.ToString());

            //Переход к следующей работе
            currWork++;
        }

        /// <summary>
        /// Установка клиента
        /// </summary>
        /// <param name="sunvision">Подразделение</param>
        /// <param name="Name">Имя клиента</param>
        void saveData.SetClient(string sunvision, string Name)
        {
            if (initial == false) throw new ExceptionError(21, "Не инициализирован класс для работы с заказами", "Location:ExelManager");

            //Указание подразделения
            write(3, "B", sunvision);
            //Указание заказчика
            write(4, "B", Name);
            //Установка заказчика для правильного формирования имени
            CurrName = Name;
            
        }

        /// <summary>
        /// Установка ответсвенного
        /// </summary>
        /// <param name="num"></param>
        void saveData.SetMain(short num)
        {
            if (initial == false) throw new ExceptionError(21, "Не инициализирован класс для работы с заказами", "Location:ExelManager");

            //Установка ответсвенного 
            switch (num)
            {
                case 1:
                    write(2, "G", "Директор");
                    write(3, "J", "И.А. Иванов");
                    break;
                case 2:
                    write(2, "G", "Зам директора по УР");
                    write(3, "J", "О.В. Федорова");
                    break;
            }

            //Установка сегодняшней даты
            System.DateTime dateT = DateTime.Now;
            string data = dateT.Day + " " + GetMonth(dateT.Month) + " " + dateT.Year + " г.";
            write(4, "G", data);
        }

        /// <summary>
        /// Получение сокращённого названия месяца по числу
        /// </summary>
        /// <param name="Month"></param>
        /// <returns></returns>
        string GetMonth(int Month)
        {
            switch (Month)
            {
                case 1:
                    return "января";
                case 2:
                    return "февраля";
                case 3:
                    return "марта";
                case 4:
                    return "апреля";
                case 5:
                    return "мая";
                case 6:
                    return "июня";
                case 7:
                    return "июля";
                case 8:
                    return "августа";
                case 9:
                    return "сентября";
                case 10:
                    return "октября";
                case 11:
                    return "ноября";
                case 12:
                    return "декабря";
                default:
                    throw new ExceptionError(51, "Не правильная дата", "Location:ExelManeger");

            }
                        
        }

        /// <summary>
        /// Отмена измений файла
        /// </summary>
        void saveData.Cancel()
        {
            quitNoSave();
            initial = false;
            //Сброс счётчика
            currWork = 0;
        }

        /// <summary>
        /// Возвращение имени последнего заказа
        /// </summary>
        /// <returns></returns>
        public string GetLastFileName()
        {
            return LastFileName;
        }

        /// <summary>
        /// Возвращение номера последнего заказа
        /// </summary>
        /// <returns></returns>
        public int GetLastFileNumber()
        {
            return LastFileNumber;
        }

        /// <summary>
        /// Установка номер заказ
        /// </summary>
        /// <param name="order"> Номер заказа</param>
        public void SetNumberOrder(int number)
        {
            //Вызов метода для записи в EXEL файл
            this.write(6, "E", LastFileNumber.ToString());
        }
    }
}
