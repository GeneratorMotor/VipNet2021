using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace RegKor.Classess
{
    class ГодДокументооборота
    {
        /// <summary>
        /// возвращает истину если год документа 2012 или больше
        /// </summary>
        /// <returns>логическое значение true если год больше 2012 года</returns>
        public bool ГодВКонфигурационномФайле()
        {
            bool равенИлиБольше;
            //string год = ConfigurationSettings.AppSettings["ГодДокумента"].ToString();

            ВерсияРаботы версияРаботы = new ВерсияРаботы();
            равенИлиБольше = версияРаботы.УзнатьВерсиюРаботы();



            //if (год == true)
            //{
            //    равенИлиБольше = true;
            //}
            //else
            //{
            //    равенИлиБольше = false;
            //}
            return равенИлиБольше;
        }
    }
}
