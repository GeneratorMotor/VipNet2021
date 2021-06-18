using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace RegKor.Classess
{
    class СтатистикаОтправленныхДокументов
    {
        public DataSet ВременнойИнтервал(string НачалоОтчета, string КонецОтчёта)
        {
            //string scommand = "select ОписаниеКорреспондента,COUNT(*) AS 'Исходящие документы' from Выборка " +
            //                  "where РезультатВыполнения <> '' and ДатаИсхода >= '"+ НачалоОтчета +"' and ДатаИсхода <= '" + КонецОтчёта +"' " +
            //                  "group by ОписаниеКорреспондента";


            string scommand = "select ОписаниеАдресата,COUNT(*) AS 'Исходящие документы' from ВыборкаИсходящихДокументов " +
                              "where Дата >= '" + НачалоОтчета + "' and Дата <= '" + КонецОтчёта + "' " +
                              "group by ОписаниеАдресата";

            SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["строкаДокументооборот"].ToString());
            con.Open();
            SqlDataAdapter da = new SqlDataAdapter(scommand, con);
            DataSet ds = new DataSet();
            da.Fill(ds);
            con.Close();
            return ds;
        }
    }
}
