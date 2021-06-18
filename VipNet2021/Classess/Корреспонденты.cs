using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    class Корреспонденты
    {
        //public DataSet ЗаполнитьКорреспонденты_DataSet()
        public DataSet ЗаполнитьКорреспонденты_DataSet()
        {
            string scommand = "select * from Корреспонденты where Удален is null";
            БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

            SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
            SqlDataAdapter daКорреспонденты = new SqlDataAdapter(scommand,con);

            DataSet dsКорреспонденты = new DataSet();
            con.Open();

            daКорреспонденты.Fill(dsКорреспонденты);
            con.Close();

            return dsКорреспонденты;
        }

        /// <summary>
        /// Скрывает запись в базе данных
        /// </summary>
        /// <param name="Корреспондент"></param>
        public void Скрыть(string Корреспондент)
        {
            string sCommand = "update Корреспонденты set Удален = 1 where ОписаниеКорреспондента = '"+ Корреспондент +"'";
            БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

            SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
            SqlCommand com = new SqlCommand(sCommand, con);

            con.Open();
            com.ExecuteNonQuery();
            con.Close();

        }

        public DataSet ПоказатьСкрытые()
        {
            string sCommand = "Select * from Корреспонденты where Удален is not null";

            БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

            SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
            SqlDataAdapter daКорреспонденты = new SqlDataAdapter(sCommand, con);

            DataSet dsКорреспонденты = new DataSet();
            con.Open();

            daКорреспонденты.Fill(dsКорреспонденты);
            con.Close();

            return dsКорреспонденты;
        }

        public void Открыть(string Корреспондент)
        {
            string sCommand = "update Корреспонденты set Удален = NULL where ОписаниеКорреспондента = '" + Корреспондент + "'";

            БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

            SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
            SqlCommand com = new SqlCommand(sCommand, con);

            con.Open();
            com.ExecuteNonQuery();
            con.Close();
        }
                
    }
}
