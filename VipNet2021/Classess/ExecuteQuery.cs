using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    /// <summary>
    /// Выполняет Sql инструкцию.
    /// </summary>
    public class ExecuteQuery
    {
        private string sqlQuery = string.Empty;

        public ExecuteQuery(string query)
        {
            sqlQuery = query;
        }

        /// <summary>
        /// Выполнение sql команду.
        /// </summary>
        public void Excecute()
        {
            ПодключитьБД conn = new ПодключитьБД();
            string sConnect = conn.СтрокаПодключения();

            using (SqlConnection con = new SqlConnection(sConnect))
            {
                con.Open();
                SqlCommand com = new SqlCommand(sqlQuery.Trim(), con);
                com.ExecuteNonQuery();
            }

        }
    }
}
