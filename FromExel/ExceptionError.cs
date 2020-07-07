using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FromExel
{
    /// <summary>
    /// Исключение, выбрасываемые внутри приложения
    /// все ошибки описаны в файле ErrorCode
    /// </summary>
    class ExceptionError : Exception
    {
        /// <summary>
        /// Номер исключения
        /// </summary>
        public short innerNumber;

        /// <summary>
        /// Внетреннее описание исключения
        /// </summary>
        public string innerText;

        /// <summary>
        /// Местоположение исключения
        /// </summary>
        public string location;

        /// <summary>
        /// Изначальное исключение
        /// </summary>
        /// <returns></returns>
        public Exception orignExcept = null;

        /// <summary>
        /// Создание внутреннего исключения
        /// </summary>
        /// <param name="Num">Номер исключения</param>
        /// <param name="Text">Текст исключения</param>
        /// <param name="Location">Место появления исключения</param>
        public ExceptionError(short Num, string Text, string Location)
        {
            innerNumber = Num;
            innerText = Text;
            location = Location;
        }
    }
}
