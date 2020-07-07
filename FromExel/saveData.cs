using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FromExel
{
    //На случай перехода на БД этот класс

    /// <summary>
    /// Интерфейс для сохранения заказов
    /// </summary>
    interface saveData
    {
        /// <summary>
        /// Указать заказчика
        /// </summary>
        /// <param name="sunvision">Подразделение</param>
        /// <param name="Name">Имя</param>
        void SetClient(string sunvision, string Name);

        /// <summary>
        /// Указать ответственного; 
        /// 1 - Директор
        /// 2 - Зам. Директора
        /// </summary>
        /// <param name="num">Номер ответсвенного</param>
        void SetMain(short num);

        /// <summary>
        /// Добавить работу в отчёт
        /// </summary>
        /// <param name="NameWork">Наименование работы</param>
        /// <param name="format">Формат</param>
        /// <param name="countSheet">Количетсво едениц в экземпляре</param>
        /// <param name="countExem">Количество экземпляров</param>
        /// <param name="Sum">Суммарное количество</param>
        void AddWork(string NameWork, string format, int countSheet, int countExem, int Sum);

        /// <summary>
        /// Начало работы с заказом
        /// </summary>
        void StartWork();

        /// <summary>
        /// Конец работы с заказом
        /// </summary>
        void EndWork();

        /// <summary>
        /// Отмена измнеий 
        /// </summary>
        void Cancel();

        /// <summary>
        /// Получение имени последнего сохранённого заказа
        /// </summary>
        /// <returns></returns>
        string GetLastFileName();

        /// <summary>
        /// Установка номера работа текстов в заказе
        /// </summary>
        /// <param name="order">Номер заказа, который необходимо установить в бумажке</param>
        void SetNumberOrder(int order);

        /// <summary>
        /// Получение номера последнего сохранённого заказа
        /// </summary>
        /// <returns></returns>
        int GetLastFileNumber();
        
    }

}
