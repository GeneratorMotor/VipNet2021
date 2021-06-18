using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlClient;
using RegKor.Classess;

namespace RegKor.Classess
{
    public class Statistic
    {
        string dataYear = string.Empty;

        public Statistic(int year)
        {
            dataYear = year.ToString().Trim() + "0101";
        }

        /// <summary>
        /// ����� �������� � ��������� ����������.
        /// </summary>
        /// <param name="sCon"></param>
        /// <param name="transaction"></param>
        /// <returns></returns>
        public DataTable ���������������(SqlConnection sCon)
        {
            string query = " declare @id1 int " +
                           "  declare @id2 int " +
                           " SELECT @id1 = COUNT(id_��������) FROM [��������������������������] " +
                             " where ���� >= '" + dataYear + "' " +
                             " select @id2 = COUNT(id_��������) from ������� " +
                             " where ���������� >= '" + dataYear + "' " +
                             " select @id1 + @id2 ";

            return  DataTableSql.GetDataTable(query, sCon);
        }

        /// <summary>
        /// ����� �������� ����������.
        /// </summary>
        /// <param name="sCon"></param>
        /// <returns></returns>
        public DataTable �����������������������(SqlConnection sCon)
        {
            string query = " select COUNT(id_��������) from ������� " +
                           " where ���������� >= '" + dataYear + "' ";

            return DataTableSql.GetDataTable(query, sCon);
        }

        /// <summary>
        /// ����� ���������� ������������ �� ��������.
        /// </summary>
        /// <param name="sCon"></param>
        /// <returns></returns>
        public DataTable �������������������������������������(SqlConnection sCon)
        {
            string query = " select COUNT(id_��������) from ������� " +
                           " where ���������� >= '" + dataYear + "'  and (���������� = 'True' and ����� = 'False') or ( ���������� = 'True' and ����� = 'True') ";

            return DataTableSql.GetDataTable(query, sCon);
        }

        ///// <summary>
        /// ����� ���������� ���������� ������������ �� ��������.
        /// </summary>
        /// <param name="sCon"></param>
        /// <returns></returns>
        public DataTable ������������������������������������������������(SqlConnection sCon)
        {
            string query = " select COUNT(id_��������) from ������� " +
                           " where ���������� >= '" + dataYear + "'  and ���������� = 'True' and ����� = 'True' and LEN(�������������������) > 2 ";

            return DataTableSql.GetDataTable(query, sCon);
        }

        /// <summary>
        /// ����� ��������� ����������.
        /// </summary>
        /// <param name="sCon"></param>
        /// <returns></returns>
        public DataTable ������������������������(SqlConnection sCon)
        {
            string query = " select COUNT(id_��������) from ����������������� " +
                          " where ���� >= '" + dataYear + "' ";

            return DataTableSql.GetDataTable(query, sCon);
        }



        
                
    }
}
