using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    class Выборка
    {
        //public DataSet ВыборкаОткрытыхЗаписей()
        //{
        //    string scommand = "select * from Выборка where ОписаниеКорреспондента in  (select ОписаниеКорреспондента from Корреспонденты where Удален is null)";
        //    БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

        //    SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
        //    SqlDataAdapter daВыборка = new SqlDataAdapter(scommand, con);

        //    DataSet dsВыборка = new DataSet();
        //    con.Open();

        //    daВыборка.Fill(dsВыборка);
        //    con.Close();

        //    return dsВыборка;
        //}

        public DataSet ВыборкаНеПросроченныеДокументы(string person)
        {
            string scommand = "select * from ViewВыборкаОтчет where ОписаниеПолучателя LIKE '%" + person + "%' AND НаКонтроле='True' AND ВДело='False' and СрокВыполнения > '" + Время.Дата(DateTime.Now.Date.ToShortDateString()) + "' ";
            БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

            SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
            SqlDataAdapter daВыборка = new SqlDataAdapter(scommand, con);

            DataSet dsВыборка = new DataSet();
            con.Open();

            daВыборка.Fill(dsВыборка);
            con.Close();

            return dsВыборка;
        }

        /// <summary>
        /// Просроченные документы срок выполнения которых меньше текущей даты.
        /// </summary>
        /// <param name="person"></param>
        /// <returns></returns>
        public DataSet ВыборкаПросроченныеДокументы(string person)
        {
            string scommand = "select * from ViewВыборкаОтчет where ОписаниеПолучателя LIKE '%" + person + "%' AND НаКонтроле='True' AND ВДело='False' and СрокВыполнения<='" + Время.Дата(DateTime.Now.Date.ToShortDateString()) + "' ";
            БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

            SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
            SqlDataAdapter daВыборка = new SqlDataAdapter(scommand, con);

            DataSet dsВыборка = new DataSet();
            con.Open();

            daВыборка.Fill(dsВыборка);
            con.Close();

            return dsВыборка;
        }

        /// <summary>
        /// Возвращает количество документов на контроле.
        /// </summary>
        /// <returns></returns>
        public DataSet ВыборкаДокументовНаКонтроле(string person)
        {
            string scommand = "select * from ViewВыборкаОтчет where ОписаниеПолучателя LIKE '%" + person + "%' AND НаКонтроле='True' AND ВДело='False' ";
            БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

            SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
            SqlDataAdapter daВыборка = new SqlDataAdapter(scommand, con);

            DataSet dsВыборка = new DataSet();
            con.Open();

            daВыборка.Fill(dsВыборка);
            con.Close();

            return dsВыборка;
        }

        // Документы у которых завтра заканчивается срок.
        public DataTable ВыборкаДокументовИстекающимСроком()//string person)
        {
            string scommand = "select * from ViewВыборкаОтчет";// +
            //"where СрокВыполнения<'" + DateTime.Now.Date + "' AND ОписаниеПолучателя LIKE '%" + person + "%' AND НаКонтроле=True AND ВДело=False ";
            
            БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

            SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
            SqlDataAdapter daВыборка = new SqlDataAdapter(scommand, con);

            DataTable dsВыборка = new DataTable();
            con.Open();

            daВыборка.Fill(dsВыборка);
            con.Close();

            return dsВыборка;//.Rows;
        }

        public DataRow[] ВыборкаВсегоПолучателей()//(string person)
        {
            string scommand = "select ОписаниеПолучателя from ViewВыборкаОтчет " +
                              "where ДатаПоступ >= '20170112' AND НаКонтроле= 'True' AND ВДело= 'False' " +
                              "group by ОписаниеПолучателя ";

            БазаДанныхДокументооборот документооборот = new БазаДанныхДокументооборот();

            SqlConnection con = new SqlConnection(документооборот.СтрокаПодключения("строкаДокументооборот"));
            SqlDataAdapter daВыборка = new SqlDataAdapter(scommand, con);

            DataTable dsВыборка = new DataTable();
            con.Open();

            daВыборка.Fill(dsВыборка);
            con.Close();

            return dsВыборка.Select();
        }


            
                
    }
}
