using System;
using System.Collections.Generic;
using System.IO;

namespace FromExel
{
    /// <summary>
    /// Чтение файла конфигураций
    /// </summary>
    class SetupConfig
    {
        /// <summary>
        /// Аттрибуты, хранимые в файле конфигурации
        /// </summary>
        Dictionary<string, string> att;

        /// <summary>
        /// Чтение файла 
        /// </summary>
        public SetupConfig()
        {
            att = new Dictionary<string, string>();
            //Открытие файла с конфигурациями
            FileStream cf = File.OpenRead("config.ini");
            StreamReader cfr = new StreamReader(cf);
            //Чтение файла по строчке
            string str = cfr.ReadLine();
            while(str != null)
            {
                //Поиск в строке символа =
                int pos = -1;
                pos = str.IndexOf('=');
                //Выделение аттрибута и его значения
                string satt, val;
                satt = str.Substring(0, pos);
                val = str.Substring(pos + 1, str.Length - 1 - pos);
                //Добавление новых аттрибутов
                att.Add(satt, val);
                //Чтение новой строки
                str = cfr.ReadLine();
            }

        }

        /// <summary>
        /// Получение атрибута по имени
        /// </summary>
        /// <param name="attName"></param>
        /// <returns></returns>
        public string this[string attName]
        {
            get { return att[attName]; }
        }
    }
}
