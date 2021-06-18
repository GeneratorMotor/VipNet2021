using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    interface IFillDataSet
    {
        void FillTable(string query, DataSet dataSet, string название“аблицы, SqlConnection connection, SqlTransaction transaction);
        //void UpdateTable(string query, DataSet dataSet, string название“аблицы, SqlConnection connection, SqlTransaction transaction);
    }
}
