using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    /// <summary>
    /// Класс генерирующий отчет.
    /// </summary>
    public  class GenerateStaticDocInput
    {
        private List<StatisticDocInput> list;
        private StatisticDocInput itemReportCount;

        public GenerateStaticDocInput(List<StatisticDocInput> listReport)
        {
            list = listReport;
            itemReportCount = new StatisticDocInput();
        }

        public StatisticDocInput ItemCount
        {
            get
            {
                return itemReportCount;
            }

        }

        public void Generate(DateTime dataStart, DateTime dataEnd, string fio)
        {
            DataTable dTab = new DataTable();
            ПодключитьБД strConnect = new ПодключитьБД();

            StatisticDocInput itemCount = new StatisticDocInput();
            
            // Получим статистику входящих писем.
            using (SqlConnection con = new SqlConnection(strConnect.СтрокаПодключения()))
            {
                SqlCommand com = new SqlCommand("ReportStatInputLetter", con);
                com.CommandType = CommandType.StoredProcedure;

                com.Parameters.Add(new SqlParameter("@dateStart", SqlDbType.DateTime));
                com.Parameters["@dateStart"].Value = dataStart;// ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString());

                com.Parameters.Add(new SqlParameter("@dateEnd", SqlDbType.DateTime));
                com.Parameters["@dateEnd"].Value = dataEnd;// ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString());

                com.Parameters.Add(new SqlParameter("@fio", SqlDbType.VarChar,100));
                com.Parameters["@fio"].Value = fio;

                SqlDataAdapter da = new SqlDataAdapter(com);
                da.Fill(dTab);
            }

            int iNum = 1;
            string person = string.Empty;

            // Заполним коллекцию классов описывающий отчет данными.
            foreach (DataRow row in dTab.Rows)
            {
                StatisticDocInput item = new StatisticDocInput();


                item.Num = iNum.ToString();
                item.НаименованиеКорреспондента = row["ОписаниеКорреспондента"].ToString().Trim();
                item.КолвоВходКорреспонденции = Convert.ToInt32(row["КоличествоВходящихДокументов"]);
                itemCount.КолвоВходКорреспонденции += Convert.ToInt32(row["КоличествоВходящихДокументов"]);

                if (row["БумажныйНоситель"] != DBNull.Value)
                {
                    item.БумажныйНоситель = Convert.ToInt32(row["БумажныйНоситель"]);
                    itemCount.БумажныйНоситель += Convert.ToInt32(row["БумажныйНоситель"]);
                }
                else
                {
                    item.БумажныйНоситель = null;
                }

                if (row["e-mail"] != DBNull.Value)
                {
                    item.Email = Convert.ToInt32(row["e-mail"]);
                    itemCount.Email += Convert.ToInt32(row["e-mail"]);
                }
                else
                {
                    item.Email = null;
                }

                if (row["VipNet"] != DBNull.Value)
                {
                    item.VipNet = Convert.ToInt32(row["VipNet"]);
                    itemCount.VipNet += Convert.ToInt32(row["VipNet"]);
                }
                else
                {
                    item.VipNet = null;
                }

                if (row["Fax"] != DBNull.Value)
                {
                    item.Fax = Convert.ToInt32(row["Fax"]);
                    itemCount.Fax += Convert.ToInt32(row["Fax"]);
                }
                else
                {
                    item.Fax = null;
                }
                item.Исполнитель = fio.Trim();
                itemCount.Исполнитель = fio.Trim();

                list.Add(item);

                iNum++;

                //itemReportCount.БумажныйНоситель += itemCount.БумажныйНоситель;
                //itemReportCount.КолвоВходКорреспонденции += itemCount.КолвоВходКорреспонденции;
                //itemReportCount.VipNet += itemCount.VipNet;
                //itemReportCount.Fax += itemCount.Fax;
                //itemReportCount.Email += itemCount.Email;

            }

            itemCount.Num = "Итого по исполнителю " + itemCount.Исполнитель.Trim();

            itemReportCount = itemCount;


            list.Add(itemCount);
        }

    }
}
