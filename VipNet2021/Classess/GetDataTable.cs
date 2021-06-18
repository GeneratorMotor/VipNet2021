using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;


namespace RegKor.Classess
{
    public class GetDataTable
    {
        // ���������� ��� �������� ������ �������.
        string strQuery = string.Empty;

        public GetDataTable(string query)
        {
            strQuery = query;
        }


        /// <summary>
        /// ���������� ���������� �������.
        /// </summary>
        /// <returns></returns>
        public DataTable DataTable()
        {
            ������������ stringConnection = new ������������();
            string sCon = stringConnection.�����������������();

            DataSet ds = new DataSet();

            using (SqlConnection con = new SqlConnection(sCon))
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(strQuery, con);
                da.Fill(ds, "������������������");
            }

            return ds.Tables["������������������"];
        }

        /// <summary>
        /// ���������� ����������� �������.
        /// </summary>
        /// <param name="nametable"></param>
        /// <returns></returns>
        public DataTable DataTable(string nametable)
        {
            ������������ stringConnection = new ������������();
            string sCon = stringConnection.�����������������();

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
            //������������ stringConnection = new ������������();
            //string sCon = stringConnection.�����������������();

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
