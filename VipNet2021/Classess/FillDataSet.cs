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
        /// ��������� ������� � ����������� DataSet �������
        /// </summary>
        /// <param name="query"></param>
        /// <param name="dataSet"></param>
        /// <param name="���������������"></param>
        /// <param name="connection"></param>
        /// <param name="transaction"></param>
        public void FillTable(string query, DataSet dataSet, string ���������������, SqlConnection connection, SqlTransaction transaction)
        {
            //SqlConnection con = new SqlConnection(�����������������);
            //con.Open();
            SqlDataAdapter da = new SqlDataAdapter(query, connection);
            da.SelectCommand.Transaction = transaction;

            da.Fill(dataSet, ���������������);
            connection.Close();
        }


        //public void UpdateTable(string query, DataSet dataSet, string ���������������, SqlConnection connection, SqlTransaction transaction)
        //{
        //    throw new Exception("The method or operation is not implemented.");
        //}


    }
}
