using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;



namespace RegKor.Classess
{
    public class СтатусКарточка
    {
        private int idКарточка = 0;
        public СтатусКарточка(int id_карточка)
        {
            idКарточка = id_карточка;
        }

        /// <summary>
        /// Если true то возвращает статус повторяющегося ответа на данную карточку.
        /// </summary>
        /// <returns></returns>
        public bool СтатусПовторяющийсяОтвет()
        {
            string query = "select FlagCardRepeet from Карточка " +
                           "where id_карточки = " + idКарточка + " ";

            GetDataTable tab = new GetDataTable(query);
            bool flag = Convert.ToBoolean(tab.DataTable("КарточкаСтатус").Rows[0]["FlagCardRepeet"]);

            return flag;
        }

        /// <summary>
        /// Если true, то ответ повторный.
        /// </summary>
        /// <returns></returns>
        public bool GetОтветПовторный(SqlConnection con)
        {
            string query = "select ВДело from Карточка " +
                          "where id_карточки = " + idКарточка + " ";

            GetDataTable tab = new GetDataTable(query);
            bool flag = Convert.ToBoolean(tab.DataTableToConnect("СтатусВДело", con).Rows[0]["ВДело"]);

            return flag;
        }
    }
}
