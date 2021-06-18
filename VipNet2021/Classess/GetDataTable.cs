using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;


namespace RegKor.Classess
{
    public class GetDataTable
    {
        // Переменная для хранения строки запроса.
        string strQuery = string.Empty;

        public GetDataTable(string query)
        {
            strQuery = query;
        }


        /// <summary>
        /// Возвращает заполенную таблицу.
        /// </summary>
        /// <returns></returns>
        public DataTable DataTable()
        {
            ПодключитьБД stringConnection = new ПодключитьБД();
            string sCon = stringConnection.СтрокаПодключения();

            DataSet ds = new DataSet();

            using (SqlConnection con = new SqlConnection(sCon))
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(strQuery, con);
                da.Fill(ds, "ПерсональныеДанные");
            }

            return ds.Tables["ПерсональныеДанные"];
        }

        /// <summary>
        /// Возвращает заполненную таблицу.
        /// </summary>
        /// <param name="nametable"></param>
        /// <returns></returns>
        public DataTable DataTable(string nametable)
        {
            ПодключитьБД stringConnection = new ПодключитьБД();
            string sCon = stringConnection.СтрокаПодключения();

            DataSet ds = new DataSet();

            using (SqlConnection con = new SqlConnection(sCon))
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(strQuery, con);
                da.Fill(ds, nametable);
            }

            return ds.Tables[nametable];
        }

        public DataTable DataTableToConnect(string nametable, SqlConnection con)
        {
            //ПодключитьБД stringConnection = new ПодключитьБД();
            //string sCon = stringConnection.СтрокаПодключения();

            DataSet ds = new DataSet();

            //using (SqlConnection con = new SqlConnection(sCon))
            //{
                //con.Open();
                SqlDataAdapter da = new SqlDataAdapter(strQuery, con);
                da.Fill(ds, nametable);
            //}

            return ds.Tables[nametable];
        }

        public void DataTableSqlTransaction(SqlConnection con, SqlTransaction transact)
        {
            DataSet ds = new DataSet();

            SqlCommand com = new SqlCommand();
            com.Transaction = transact;

            com.ExecuteNonQuery();
        }


    }
}
