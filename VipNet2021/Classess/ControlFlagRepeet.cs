using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    /// <summary>
    /// Контролирует флаг повторнения записи.
    /// </summary>

    public class ControlFlagRepeet
    {
        private bool controlFlag;
        private int id;


        public ControlFlagRepeet(int id_карточки, bool flag)
        {
            controlFlag = flag;
            id = id_карточки;
        }

        /// <summary>
        /// Возвращает флаг указывающий в каком положении карточка.
        /// </summary>
        /// <returns></returns>
        public bool CompareRepet()
        {

            bool flag;
            ПодключитьБД connect = new ПодключитьБД();
            string sCon = connect.СтрокаПодключения();

            string query = "select FlagCardRepeet from Карточка " +
                           "where id_карточки = "+ id +" ";

            using(SqlConnection con = new SqlConnection(sCon))
            {
                con.Open();
                SqlDataAdapter da = new SqlDataAdapter(query, con);
                
                DataSet ds = new DataSet();
                da.Fill(ds, "КарточкаКонтрол");

                DataRow row = ds.Tables["КарточкаКонтрол"].Rows[0];
                flag = Convert.ToBoolean(row[0]);
            }

            return flag;
        }
    }
}
