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
        /// Всего входящих и исходящих документов.
        /// </summary>
        /// <param name="sCon"></param>
        /// <param name="transaction"></param>
        /// <returns></returns>
        public DataTable ВсегоДокументов(SqlConnection sCon)
        {
            string query = " declare @id1 int " +
                           "  declare @id2 int " +
                           " SELECT @id1 = COUNT(id_карточки) FROM [ВыборкаИсходящихДокументов] " +
                             " where Дата >= '" + dataYear + "' " +
                             " select @id2 = COUNT(id_карточки) from Выборка " +
                             " where ДатаПоступ >= '" + dataYear + "' " +
                             " select @id1 + @id2 ";

            return  DataTableSql.GetDataTable(query, sCon);
        }

        /// <summary>
        /// Всего входящих документов.
        /// </summary>
        /// <param name="sCon"></param>
        /// <returns></returns>
        public DataTable ВсегоВходящихДокументов(SqlConnection sCon)
        {
            string query = " select COUNT(id_карточки) from Выборка " +
                           " where ДатаПоступ >= '" + dataYear + "' ";

            return DataTableSql.GetDataTable(query, sCon);
        }

        /// <summary>
        /// Всего документов поставленных на контроль.
        /// </summary>
        /// <param name="sCon"></param>
        /// <returns></returns>
        public DataTable ВсегоДокументовПоставленныхНаКонтроль(SqlConnection sCon)
        {
            string query = " select COUNT(id_карточки) from Выборка " +
                           " where ДатаПоступ >= '" + dataYear + "'  and (НаКонтроле = 'True' and ВДело = 'False') or ( НаКонтроле = 'True' and ВДело = 'True') ";

            return DataTableSql.GetDataTable(query, sCon);
        }

        ///// <summary>
        /// Всего исполненно документов поставленных на контроль.
        /// </summary>
        /// <param name="sCon"></param>
        /// <returns></returns>
        public DataTable ВсегоИсполненныхДокументовПоставленныхНаКонтроль(SqlConnection sCon)
        {
            string query = " select COUNT(id_карточки) from Выборка " +
                           " where ДатаПоступ >= '" + dataYear + "'  and НаКонтроле = 'True' and ВДело = 'True' and LEN(РезультатВыполнения) > 2 ";

            return DataTableSql.GetDataTable(query, sCon);
        }

        /// <summary>
        /// Всего исходящих документов.
        /// </summary>
        /// <param name="sCon"></param>
        /// <returns></returns>
        public DataTable ВсегоИсходящихДокументов(SqlConnection sCon)
        {
            string query = " select COUNT(id_карточки) from КарточкаИсходящая " +
                          " where Дата >= '" + dataYear + "' ";

            return DataTableSql.GetDataTable(query, sCon);
        }



        
                
    }
}
