using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    public static class DataTableSql
    {
        /// <summary>
        /// Возвращаем таблицу в результате запроса.
        /// </summary>
        /// <param name="query">sql запрос</param>
        /// <returns>Таблица</returns>
        public static DataTable GetDataTable(string query)
        {
            DataTable table = new DataTable();

            ПодключитьБД conn = new ПодключитьБД();
            string sConnect = conn.СтрокаПодключения();

            using (SqlConnection con = new SqlConnection(sConnect))
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(query.Trim(), con);
                da.Fill(table);
           }

           return table;
        }

        /// <summary>
        /// Возвращаем таблицу в результате запроса.
        /// </summary>
        /// <param name="query">sql запрос</param>
        /// <param name="sConnect">строка подключения</param>
        /// <returns>Таблица</returns>
        public static DataTable GetDataTable(string query, SqlConnection sConnect)
        {
            DataTable table = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(query.Trim(), sConnect);
            da.Fill(table);

            return table;
        }

        /// <summary>
        /// Возвращает массив строк.
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public static DataRow[] GetDataTableRows(string query)
        {
            DataTable table = new DataTable();

            ПодключитьБД conn = new ПодключитьБД();
            string sConnect = conn.СтрокаПодключения();

            using (SqlConnection con = new SqlConnection(sConnect))
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(query.Trim(), con);
                da.Fill(table);
            }

            return table.Select();;
        }

    }
}
