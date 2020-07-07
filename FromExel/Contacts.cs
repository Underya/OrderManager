using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace FromExel
{
    /// <summary>
    /// Все контакты, что есть в в EXEL таблице
    /// </summary>
    public class Contacts
    {
        /// <summary>
        /// Еденичный конткакт
        /// </summary>
        public class Contact
        {
            /// <summary>
            /// ФИО контакта
            /// </summary>
            public string Name = "";

            /// <summary>
            /// Отдел
            /// </summary>
            public string Otdel = "";

            /// <summary>
            /// Электронная почта контакта
            /// </summary>
            public string email = "";
        }

        /// <summary>
        /// Список всех контактов из файла
        /// </summary>
        List<Contact> contact;

        /// <summary>
        /// Создание всех контактов из EXEL файла
        /// </summary>
        /// <param name="path">путь к файлу с контактами</param>
        public Contacts(string path)
        {
            //Получение книжки
            ExelBook b = ExelManager.getBook(path);
            //Выбор активного листа
            b.SetActiveSheet(1);

            //Выделение памяти под контакты
            contact = new List<Contact>();

            int row = 1;
            //Цикл, в ходе которого читаются строки до первой пустой
            while (true)
            {
                //Чтение имени
                string name = b.Read(row, "A");
                //Если строка пустая, конец цикла
                if (name == "") break;
                //Иначе, получение подразделения и емейла
                string otdel = b.Read(row, "B");
                string email = b.Read(row, "C");
                //Формирование и добавление нового контакта
                Contact c = new Contact();
                c.Name = name;
                c.Otdel = otdel;
                c.email = email;
                //Добавление в список
                contact.Add(c);
                //Переход к следующей строке
                row++;

            }

            //Закрытие книги и приложения
            b.CloseNotsave();
        }

        /// <summary>
        /// Число контактов 
        /// </summary>
        public int count
        {
            get { return contact.Count; }
        }

        /// <summary>
        /// Получение i контакта
        /// </summary>
        /// <param name="i"></param>
        /// <returns></returns>
        public Contact this[int i]
        {
            get { return contact[i]; }
        }


    }
}
