using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace RegKor.Classess
{
    public class ПодключитьБД
    {
        public string СтрокаПодключения()
        {
            string sConnection = ConfigurationSettings.AppSettings["строкаДокументооборот"].ToString();
            return sConnection;
        }
    }
}
