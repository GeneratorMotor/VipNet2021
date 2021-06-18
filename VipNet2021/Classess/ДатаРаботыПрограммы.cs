using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Класс настраевает программу на работу с данными в выбранном году
    /// </summary>
    class ДатаРаботыПрограммы
    {
        /// <summary>
        /// Возвращает дату на с которой начинает работать программа
        /// </summary>
        /// <param name="выбранныйГод"></param>
        /// <returns></returns>
        public static string ДатаНастройкиПрограммы(int выбранныйГод)
        {
            //return выбранныйГод - 1 + "0101";
            return выбранныйГод + "0101";
        }

        /// <summary>
        /// Возвращает 1 января следующего года
        /// </summary>
        /// <param name="выбранныйГод"></param>
        /// <returns></returns>
        public static string ДатаСледующийГод(int выбранныйГод)
        {
            int последующийГод = выбранныйГод + 1;
            return последующийГод.ToString() + "0101"; 
        }

        /// <summary>
        /// Возвращает предыдущий год
        /// </summary>
        /// <param name="выбранныйГод"></param>
        /// <returns></returns>
        public static string ВыбранныйГод(int выбранныйГод)
        {
            int годНазад = выбранныйГод - 1;
            return годНазад.ToString();
        }
    }
}
