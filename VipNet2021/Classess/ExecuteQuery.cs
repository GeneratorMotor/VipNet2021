using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    /// <summary>
    /// ��������� Sql ����������.
    /// </summary>
    public class ExecuteQuery
    {
        private string sqlQuery = string.Empty;

        public ExecuteQuery(string query)
        {
            sqlQuery = query;
        }

        /// <summary>
        /// ���������� sql �������.
        /// </summary>
        public void Excecute()
        {
            ������������ conn = new ������������();
            string sConnect = conn.�����������������();

            using (SqlConnection con = new SqlConnection(sConnect))
            {
                con.Open();
                SqlCommand com = new SqlCommand(sqlQuery.Trim(), con);
                com.ExecuteNonQuery();
            }

        }
    }
}
