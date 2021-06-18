using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace RegKor.Classess
{
    class БазаДанныхДокументооборот
    {
        public string СтрокаПодключения(string КлючСтрокиПодключения)
        {
            string sCon = ConfigurationSettings.AppSettings[КлючСтрокиПодключения].ToString();
            return sCon;
        }
    }
}
