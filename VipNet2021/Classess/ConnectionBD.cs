using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;

namespace RegKor.Classess
{
    class ��������������������������������
    {
        public DataSet �����������������(string ������������, string �����������)
        {
            //string scommand = "select ����������������������,COUNT(*) AS '��������� ���������' from ������� " +
            //                  "where ������������������� <> '' and ���������� >= '"+ ������������ +"' and ���������� <= '" + ����������� +"' " +
            //                  "group by ����������������������";


            string scommand = "select ����������������,COUNT(*) AS '��������� ���������' from �������������������������� " +
                              "where ���� >= '" + ������������ + "' and ���� <= '" + ����������� + "' " +
                              "group by ����������������";

            SqlConnection con = new SqlConnection(ConfigurationSettings.AppSettings["���������������������"].ToString());
            con.Open();
            SqlDataAdapter da = new SqlDataAdapter(scommand, con);
            DataSet ds = new DataSet();
            da.Fill(ds);
            con.Close();
            return ds;
        }
    }
}
