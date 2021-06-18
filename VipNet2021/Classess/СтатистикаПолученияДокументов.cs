using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    public static class СтатистикаПолученияДокументов
    {
        public static List<ItemStatisticDoc> GetStatisticDoc(int year, int numMonth)
        {
            string query = "SELECT     dbo.Получатели.ОписаниеПолучателя, dbo.ВидПоступленияДокумента.ВидПоступленияДокумента, " +
                           "COUNT(dbo.СвязующаяВидПоступленияДокПолучатели.id_ВидПоступленияДок) AS Количество " +
                           "FROM         dbo.СвязующаяВидПоступленияДокПолучатели INNER JOIN " +
                           "dbo.ВидПоступленияДокумента ON " +
                           "dbo.ВидПоступленияДокумента.id = dbo.СвязующаяВидПоступленияДокПолучатели.id_ВидПоступленияДок INNER JOIN " +
                           "dbo.Получатели ON dbo.Получатели.id_получателя = dbo.СвязующаяВидПоступленияДокПолучатели.id_person INNER JOIN " +
                           "dbo.Карточка ON dbo.Карточка.id_карточки = dbo.СвязующаяВидПоступленияДокПолучатели.id_карточки " +
                           "WHERE     (MONTH(dbo.Карточка.ДатаПоступ) = "+ numMonth +") and (YEAR(dbo.Карточка.ДатаПоступ) = "+ year +") " +
                           "GROUP BY dbo.Получатели.ОписаниеПолучателя, dbo.ВидПоступленияДокумента.ВидПоступленияДокумента,  " +
                           "dbo.СвязующаяВидПоступленияДокПолучатели.id_ВидПоступленияДок";

            DataTable tab = DataTableSql.GetDataTable(query);

            // Список для хранения статистики.
            List<ItemStatisticDoc> listStatic = new List<ItemStatisticDoc>();

            if (tab.Rows.Count > 0)
            {
                foreach (DataRow row in tab.Rows)
                {
                    ItemStatisticDoc item = new ItemStatisticDoc();
                    item.ФИО = row["ОписаниеПолучателя"].ToString().Trim();
                    item.ВидПоступления = row["ВидПоступленияДокумента"].ToString().Trim();
                    item.Count = Convert.ToInt32(row["Количество"]);

                    listStatic.Add(item);
                }
            }

            return listStatic;
        }

        /// <summary>
        /// Возвращает итоговое значение отчета.
        /// </summary>
        /// <param name="year"></param>
        /// <param name="numMonth"></param>
        /// <returns></returns>
        public static List<ItemStatisticDoc> GetStatisticDocCount(int year, int numMonthStart, int numMonthEnd)
        {
            string query = "SELECT     dbo.Получатели.ОписаниеПолучателя, dbo.ВидПоступленияДокумента.ВидПоступленияДокумента, COUNT(dbo.СвязующаяВидПоступленияДокПолучатели.id_ВидПоступленияДок) AS Количество FROM   dbo.СвязующаяВидПоступленияДокПолучатели " +
                           "INNER JOIN dbo.ВидПоступленияДокумента  " +
                            "ON dbo.ВидПоступленияДокумента.id = dbo.СвязующаяВидПоступленияДокПолучатели.id_ВидПоступленияДок  " +
                            "INNER JOIN dbo.Получатели  " +
                            "ON dbo.Получатели.id_получателя = dbo.СвязующаяВидПоступленияДокПолучатели.id_person  " +
                            "INNER JOIN dbo.Карточка  " + 
                            "ON dbo.Карточка.id_карточки = dbo.СвязующаяВидПоступленияДокПолучатели.id_карточки  " +
                            "WHERE     (MONTH(dbo.Карточка.ДатаПоступ) >= "+ numMonthStart +") and (MONTH(dbo.Карточка.ДатаПоступ) <= "+ numMonthEnd +") and (YEAR(dbo.Карточка.ДатаПоступ) = "+ year +")  " +
                            "GROUP BY dbo.Получатели.ОписаниеПолучателя, dbo.ВидПоступленияДокумента.ВидПоступленияДокумента,  dbo.СвязующаяВидПоступленияДокПолучатели.id_ВидПоступленияДок ";


            DataTable tab = DataTableSql.GetDataTable(query);

            // Список для хранения статистики.
            List<ItemStatisticDoc> listStatic = new List<ItemStatisticDoc>();

            if (tab.Rows.Count > 0)
            {
                foreach (DataRow row in tab.Rows)
                {
                    ItemStatisticDoc item = new ItemStatisticDoc();
                    item.ФИО = row["ОписаниеПолучателя"].ToString().Trim();
                    item.ВидПоступления = row["ВидПоступленияДокумента"].ToString().Trim();
                    item.Count = Convert.ToInt32(row["Количество"]);

                    listStatic.Add(item);
                }
            }

            return listStatic;
        }

        /// <summary>
        /// Возвращает общую статистику поступления документов по месяцам.
        /// </summary>
        /// <param name="year"></param>
        /// <param name="numMonth"></param>
        /// <returns></returns>
        public static List<ItemStatisticDoc> GetStatisticMontch(int year, int numMonth, string vidDoc)
        {
            string query = "SELECT  dbo.ВидПоступленияДокумента.ВидПоступленияДокумента, COUNT(dbo.СвязующаяВидПоступленияДокПолучатели.id_ВидПоступленияДок) AS Количество FROM dbo.СвязующаяВидПоступленияДокПолучатели " +
                           "INNER JOIN dbo.ВидПоступленияДокумента  " +
                            "ON dbo.ВидПоступленияДокумента.id = dbo.СвязующаяВидПоступленияДокПолучатели.id_ВидПоступленияДок  " +
                            "INNER JOIN dbo.Получатели  " +
                            "ON dbo.Получатели.id_получателя = dbo.СвязующаяВидПоступленияДокПолучатели.id_person  " +
                            "INNER JOIN dbo.Карточка  " +
                            "ON dbo.Карточка.id_карточки = dbo.СвязующаяВидПоступленияДокПолучатели.id_карточки  " +
                            "WHERE     (MONTH(dbo.Карточка.ДатаПоступ) = " + numMonth + ") and (YEAR(dbo.Карточка.ДатаПоступ) = " + year + " ) and  dbo.ВидПоступленияДокумента.ВидПоступленияДокумента = '" + vidDoc + "' " +
                            "GROUP BY  dbo.ВидПоступленияДокумента.ВидПоступленияДокумента,  dbo.СвязующаяВидПоступленияДокПолучатели.id_ВидПоступленияДок ";

            DataTable tab = DataTableSql.GetDataTable(query);

            // Список для хранения статистики.
            List<ItemStatisticDoc> listStatic = new List<ItemStatisticDoc>();

            if (tab.Rows.Count > 0)
            {
                foreach (DataRow row in tab.Rows)
                {
                    ItemStatisticDoc item = new ItemStatisticDoc();
                    item.ФИО = null; // row["ОписаниеПолучателя"].ToString().Trim();
                    item.ВидПоступления = row["ВидПоступленияДокумента"].ToString().Trim();
                    item.Count = Convert.ToInt32(row["Количество"]);

                    listStatic.Add(item);
                }
            }

            return listStatic;
        }
    }
}
