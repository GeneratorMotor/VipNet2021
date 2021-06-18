using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    public class СписокНачальников : IСписокНачальников
    {
        #region IСписокНачальников реализация интерфейса

        public List<string> GetDirectors()
        {
            ПодключитьБД connectDB = new ПодключитьБД();

            List<string> listPerson = new List<string>();

            string query = "select ОписаниеПолучателя from dbo.Получатели " +
                           "where Удален is null";

            // Получим таблицу с дейстыующими начальнриками отделов и управлений комитета соц защиты.
            DataTable tabDirectors = DataTableSql.GetDataTable(query);

            // Проверим есть ли звписис в таблице.
            if (tabDirectors.Rows.Count > 0)
            {
                // Получим список начальников отделов.
                foreach (DataRow row in tabDirectors.Rows)
                {
                    listPerson.Add(row["ОписаниеПолучателя"].ToString().Trim());
                }
            }

            return listPerson;
        }

        #endregion
    }
}
