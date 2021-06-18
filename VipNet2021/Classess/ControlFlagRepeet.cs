using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    /// <summary>
    /// ������������ ���� ����������� ������.
    /// </summary>

    public class ControlFlagRepeet
    {
        private bool controlFlag;
        private int id;


        public ControlFlagRepeet(int id_��������, bool flag)
        {
            controlFlag = flag;
            id = id_��������;
        }

        /// <summary>
        /// ���������� ���� ����������� � ����� ��������� ��������.
        /// </summary>
        /// <returns></returns>
        public bool CompareRepet()
        {

            bool flag;
            ������������ connect = new ������������();
            string sCon = connect.�����������������();

            string query = "select FlagCardRepeet from �������� " +
                           "where id_�������� = "+ id +" ";

            using(SqlConnection con = new SqlConnection(sCon))
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(query, con);
                
                DataSet ds = new DataSet();
                da.Fill(ds, "���������������");

                DataRow row = ds.Tables["���������������"].Rows[0];
                flag = Convert.ToBoolean(row[0]);
            }

            return flag;
        }
    }
}
