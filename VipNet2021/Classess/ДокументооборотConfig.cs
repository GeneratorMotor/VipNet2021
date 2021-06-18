using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;

namespace RegKor.Classess
{
    public static class ДокументооборотConfig
    {
        public static bool ВключитьДокументооборот()
        {
            bool flag = false;

            string flagConfig = ConfigurationSettings.AppSettings["Документооборот"].ToString();

            if (flagConfig.Trim() == "0")
            {
                flag = false;
            }
            else
            {
                flag = true;
            }
            

            return flag;
        }
    }
}
