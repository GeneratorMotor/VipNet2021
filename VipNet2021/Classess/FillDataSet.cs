using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    class FillDataSet:IFillDataSet 
    {
        /// <summary>
        /// Заполняет таблицу в определённом DataSet данными
        /// </summary>
        /// <param name="query"></param>
        /// <param name="dataSet"></param>
        /// <param name="названиеТаблицы"></param>
        /// <param name="connection"></param>
        /// <param name="transaction"></param>
        public void FillTable(string query, DataSet dataSet, string названиеТаблицы, SqlConnection connection, SqlTransaction transaction)
        {
            //SqlConnection con = new SqlConnection(строкаПодключения);
            //con.Open();
            SqlDataAdapter da = new SqlDataAdapter(query, connection);
            da.SelectCommand.Transaction = transaction;

            da.Fill(dataSet, названиеТаблицы);
            connection.Close();
        }


        //public void UpdateTable(string query, DataSet dataSet, string названиеТаблицы, SqlConnection connection, SqlTransaction transaction)
        //{
        //    throw new Exception("The method or operation is not implemented.");
        //}


    }
}
