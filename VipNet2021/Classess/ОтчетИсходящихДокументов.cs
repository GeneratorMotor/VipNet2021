using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    public class ОтчетИсходящихДокументов
    {
        private RangeDate rd;

        public ОтчетИсходящихДокументов(RangeDate диапазонДат)
        {
            rd = диапазонДат;
        }


        /// <summary>
        /// 
        /// /// <summary>
        /// Возвращает список адресатов исходящих документоа за период.
        /// </summary>
        /// <returns></returns>
        public DataTable GetАдресаты()
        {
            string query = "SELECT distinct [id_Адресата] " +
                           " ,[Адресат] FROM [ViewОтчетИсходящихДокументов] " +
                           " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "'";

            return DataTableSql.GetDataTable(query);
        }

        /// <summary>
        /// Получить всех адресатов по исполнителям.
        /// </summary>
        /// <param name="idPerson"></param>
        /// <returns></returns>
        public DataTable GetАдресатыПоИсполнителям(int idPerson)
        {
            string query = "SELECT distinct [id_Адресата] " +
                           " ,[Адресат] FROM [ViewОтчетИсходящихДокументов] " +
                           " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "' and [id_получателя] = " + idPerson + " ";

            return DataTableSql.GetDataTable(query);
        }
        
        /// <summary>
        /// Возвращает список исполнителей за период.
        /// </summary>
        /// <returns></returns>
        public DataTable GetИсполнители()
        {
            string query = "SELECT distinct id_получателя " +
                           ", [ОписаниеПолучателя] FROM [ViewОтчетИсходящихДокументов] " +
                           " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "'";

            return DataTableSql.GetDataTable(query);
        }



        public DataTable GetИсполнителиПоКорреспонденту(int idAddress)
        {
            string query = "SELECT distinct id_получателя " +
                           ", [ОписаниеПолучателя] FROM [ViewОтчетИсходящихДокументов] " +
                           " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "'  and [id_Адресата] = " + idAddress + " ";

            return DataTableSql.GetDataTable(query);
        }

        /// <summary>
        /// Возвращает все документы за указанный период.
        /// </summary>
        /// <returns></returns>
        public DataTable ВсеДокументыЗаПериод()
        {
          //string query = "SELECT [id_Адресата] " +
          //                " ,id_получателя " +
          //               " ,[Адресат] " +
          //               " ,[ДатаИсходящая] " +
          //               " ,[НомерКомитета] +'-'+" +
          //               " [НомерНоменклатурный] +'-'+ " +
          //               " [НомерПодразделения] +'/'+" +
          //               " CONVERT(VARCHAR,[НомерПорядковый]) AS 'Номер исходящий' " +
          //               " ,[Содержание] " +
          //               " ,[ОписаниеПолучателя] " +
          //               ",CONVERT(varchar,[номерПП]) + N'/' +[НомерВход] AS 'Документы, на которые дан ответ' " +
          //           " FROM [ViewОтчетИсходящихДокументов] " +
          //           " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "'";

          string query = "SELECT [id_Адресата]  ,id_получателя  ,[Адресат] " +
                         ",[ДатаИсходящая]  ,[НомерКомитета] +'-'+ [НомерНоменклатурный] +'-'+  [НомерПодразделения] +'/'+ CONVERT(VARCHAR,[НомерПорядковый]) AS 'Номер исходящий' " +
                         ",[Содержание]  ,[ОписаниеПолучателя] ,CONVERT(varchar,[номерПП]) + N'/' +[НомерВход] AS 'Документы, на которые дан ответ'  FROM [ViewОтчетИсходящихДокументов] " +
                          " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "'  and ДспDesc is null " +
                         "union " +
                         "SELECT [id_Адресата]  ,id_получателя  ,[Адресат]  ,[ДатаИсходящая] " +
                         ",[НомерКомитета] +'-'+ [НомерНоменклатурный] +'-'+  [НомерПодразделения] +'/'+ CONVERT(VARCHAR,[НомерПорядковый]) + '- ДСП' AS 'Номер исходящий' " +
                         ",[Содержание]  ,[ОписаниеПолучателя] ,CONVERT(varchar,[номерПП]) + N'/' +[НомерВход] AS 'Документы, на которые дан ответ'  FROM [ViewОтчетИсходящихДокументов] " +
                         " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "' and LOWER(RTRIM(LTRIM(ДспDesc)))  = LOWER(RTRIM(LTRIM('ДСП'))) " +
                         "order by [ДатаИсходящая] ";

            return DataTableSql.GetDataTable(query);
        }


        /// <summary>
        /// Выборка за период по корреспондентам (Адресатам).
        /// </summary>
        /// <param name="idAddress"></param>
        /// <returns></returns>
        public DataTable ДокументыВыборкаПоКорреспондентам(int idAddress)
        {
           
            string query = "SELECT [id_Адресата] " +
                         " ,id_получателя " +
                        " ,[Адресат] " +
                        " ,[ДатаИсходящая] " +
                        " ,[НомерКомитета] +'-'+" +
                        " [НомерНоменклатурный] +'-'+ " +
                        " [НомерПодразделения] +'/'+" +
                        " CONVERT(VARCHAR,[НомерПорядковый]) AS 'Номер исходящий' " +
                        " ,[Содержание] " +
                        " ,[ОписаниеПолучателя] " +
                        ",CONVERT(varchar,[номерПП]) + N'/' +[НомерВход] AS 'Документы, на которые дан ответ' " +
                    " FROM [ViewОтчетИсходящихДокументов] " +
                     " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "' and [id_Адресата] = " + idAddress + " ";

            //string query = "SELECT [id_Адресата]  ,id_получателя  ,[Адресат] " +
            //           ",[ДатаИсходящая]  ,[НомерКомитета] +'-'+ [НомерНоменклатурный] +'-'+  [НомерПодразделения] +'/'+ CONVERT(VARCHAR,[НомерПорядковый]) AS 'Номер исходящий' " +
            //           ",[Содержание]  ,[ОписаниеПолучателя] ,CONVERT(varchar,[номерПП]) + N'/' +[НомерВход] AS 'Документы, на которые дан ответ'  FROM [ViewОтчетИсходящихДокументов] " +
            //            " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "'  and ДспDesc is null " +
            //           "union " +
            //           "SELECT [id_Адресата]  ,id_получателя  ,[Адресат]  ,[ДатаИсходящая] " +
            //           ",[НомерКомитета] +'-'+ [НомерНоменклатурный] +'-'+  [НомерПодразделения] +'/'+ CONVERT(VARCHAR,[НомерПорядковый]) + '- ДСП' AS 'Номер исходящий' " +
            //           ",[Содержание]  ,[ОписаниеПолучателя] ,CONVERT(varchar,[номерПП]) + N'/' +[НомерВход] AS 'Документы, на которые дан ответ'  FROM [ViewОтчетИсходящихДокументов] " +
            //           " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "' and LOWER(RTRIM(LTRIM(ДспDesc)))  = LOWER(RTRIM(LTRIM('ДСП'))) " +
            //           "order by [ДатаИсходящая] ";

            return DataTableSql.GetDataTable(query);

        }

        /// <summary>
        /// Выборка за период по корреспондентам (Адресатам).
        /// </summary>
        /// <param name="idAddress"></param>
        /// <returns></returns>
        public DataTable ДокументыВыборкаПоИсполнителям(int idPerson)
        {
            //string query = "SELECT [id_Адресата] " +
            //               " ,id_получателя " +
            //              " ,[Адресат] " +
            //              " ,[ДатаИсходящая] " +
            //              " ,[НомерКомитета] " +
            //              " ,[НомерНоменклатурный] " +
            //              " ,[НомерПодразделения] " +
            //              " ,[НомерПорядковый] " +
            //              " ,[Содержание] " +
            //              " ,[ОписаниеПолучателя] " +
            //          " FROM [ViewОтчетИсходящихДокументов] " +
            //          " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "' and [id_получателя] = " + idPerson + " ";


            string query = "SELECT [id_Адресата] " +
                         " ,id_получателя " +
                        " ,[Адресат] " +
                        " ,[ДатаИсходящая] " +
                        " ,[НомерКомитета] +'-'+" +
                        " [НомерНоменклатурный] +'-'+ " +
                        " [НомерПодразделения] +'/'+" +
                        " CONVERT(VARCHAR,[НомерПорядковый]) AS 'Номер исходящий' " +
                        " ,[Содержание] " +
                        " ,[ОписаниеПолучателя] " +
                        ",CONVERT(varchar,[номерПП]) + N'/' +[НомерВход] AS 'Документы, на которые дан ответ' " +
                    " FROM [ViewОтчетИсходящихДокументов] " +
                     " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "' and [id_получателя] = " + idPerson + " ";

            return DataTableSql.GetDataTable(query);

        }


        /// <summary>
        /// Выборка за период по корреспондентам (Адресатам).
        /// </summary>
        /// <param name="idAddress"></param>
        /// <returns></returns>
        public DataTable ДокументыВыборкаПоИсполнителямАдресатам(int idPerson, int idAddress)
        {
            //string query = "SELECT [id_Адресата] " +
            //               " ,id_получателя " +
            //              " ,[Адресат] " +
            //              " ,[ДатаИсходящая] " +
            //              " ,[НомерКомитета] " +
            //              " ,[НомерНоменклатурный] " +
            //              " ,[НомерПодразделения] " +
            //              " ,[НомерПорядковый] " +
            //              " ,[Содержание] " +
            //              " ,[ОписаниеПолучателя] " +
            //          " FROM [ViewОтчетИсходящихДокументов] " +
            //          " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "' and [id_получателя] = " + idPerson + " and  id_Адресата = "+ idAddress +" ";


            string query = "SELECT [id_Адресата] " +
                          " ,id_получателя " +
                         " ,[Адресат] " +
                         " ,[ДатаИсходящая] " +
                         " ,[НомерКомитета] +'-'+" +
                         " [НомерНоменклатурный] +'-'+ " +
                         " [НомерПодразделения] +'/'+" +
                         " CONVERT(VARCHAR,[НомерПорядковый]) AS 'Номер исходящий' " +
                         " ,[Содержание] " +
                         " ,[ОписаниеПолучателя] " +
                         ",CONVERT(varchar,[номерПП]) + N'/' +[НомерВход] AS 'Документы, на которые дан ответ' " +
                     " FROM [ViewОтчетИсходящихДокументов] " +
                     " where [ДатаИсходящая] >= '" + ДатаSQL.Дата(rd.DataStart.ToShortDateString()) + "' and [ДатаИсходящая] <= '" + ДатаSQL.Дата(rd.DataEnd.ToShortDateString()) + "' and [id_получателя] = " + idPerson + " and  id_Адресата = " + idAddress + " ";


            return DataTableSql.GetDataTable(query);

        }



    }
}
