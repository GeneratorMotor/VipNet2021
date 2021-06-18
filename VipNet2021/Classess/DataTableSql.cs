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
        /// ���������� ������� � ���������� �������.
        /// </summary>
        /// <param name="query">sql ������</param>
        /// <returns>�������</returns>
        public static DataTable GetDataTable(string query)
        {
            DataTable table = new DataTable();

            ������������ conn = new ������������();
            string sConnect = conn.�����������������();

            using (SqlConnection con = new SqlConnection(sConnect))
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(query.Trim(), con);
                da.Fill(table);
           }

           return table;
        }

        /// <summary>
        /// ���������� ������� � ���������� �������.
        /// </summary>
        /// <param name="query">sql ������</param>
        /// <param name="sConnect">������ �����������</param>
        /// <returns>�������</returns>
        public static DataTable GetDataTable(string query, SqlConnection sConnect)
        {
            DataTable table = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(query.Trim(), sConnect);
            da.Fill(table);

            return table;
        }

        /// <summary>
        /// ���������� ������ �����.
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        public static DataRow[] GetDataTableRows(string query)
        {
            DataTable table = new DataTable();

            ������������ conn = new ������������();
            string sConnect = conn.�����������������();

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
