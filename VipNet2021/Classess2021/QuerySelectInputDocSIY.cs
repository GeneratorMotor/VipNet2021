using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess2021
{
    /// <summary>
    /// Sql Запрос на входящие документы СЭУ.
    /// </summary>
    public class QuerySelectInputDocSIY : IQueryStringSQL
    {
        // Переменные для хранения даты в представлении SQL 104.
        private string dataStart = string.Empty;
        private string dataEnd = string.Empty;

        /// <summary>
        /// Создаем запрос к БД.
        /// </summary>
        /// <param name="dateStart"> Дата начала отчета в виде '20210325' </param>
        /// <param name="dateEnd"> Дата начала отчета в виде '20210325' </param>
        public QuerySelectInputDocSIY(string dateStart, string dateEnd)
        {
            if (dateStart == null)
            {
                throw new ArgumentNullException("Нет даты начала отчета");
            }

            if (dateEnd == null)
            {
                throw new ArgumentNullException("Нет даты окончания отчета");
            }

            dataStart = dateStart;
            dataEnd = dateEnd;
        }

        public string Query()
        {
//            string query = @"SELECT [id_карточки]
//                          ,[ОписаниеКорреспондента]
//                          ,[КраткоеСодержание]
//                          ,[НомерВходящий]
//                          ,[ДатаПоступления]
//                          ,[НомерИсход]
//                      FROM [VipNet2019].[dbo].[ViewЖурналУчетаСЭУ]
//                      where ДатаПоступления >= '" + dataStart + "' and ДатаПоступления <= '" + dataEnd + "' ";

            string query = @"select id_карточки, ОписаниеКорреспондента, КраткоеСодержание, НомерВходящий, ДатаПоступления, НомерИсход,
                                ОснованиеПередачи from (
                                SELECT        Tab1.id_карточки, ОписаниеКорреспондента, КраткоеСодержание, НомерВходящий, ДатаПоступления, НомерИсход,
                                Основаниепередачи.ОснованиеПередачи
                                FROM            (SELECT        dbo.Карточка.id_карточки, dbo.Корреспонденты.ОписаниеКорреспондента, dbo.Карточка.КраткоеСодержание, dbo.Карточка.НомерВход + '/' + CAST(dbo.Карточка.номерПП AS nvarchar(10)) 
                                                                                    AS НомерВходящий, CAST(dbo.Карточка.ДатаПоступ AS date) AS ДатаПоступления, dbo.Карточка.НомерИсход
                                                          FROM            dbo.Корреспонденты INNER JOIN
                                                                                    dbo.Карточка ON dbo.Корреспонденты.id_корреспондента = dbo.Карточка.id_корреспондента
                                                        WHERE        (dbo.Карточка.НомерИсход = 'СЭУ')  
						                                  ) AS Tab1
                                left outer join СвязующаяУчетаПерсональныхДанных
                                on СвязующаяУчетаПерсональныхДанных.id_карточки = Tab1.id_карточки
                                left outer join Основаниепередачи
                                on Основаниепередачи.id_основаниеПередачи = СвязующаяУчетаПерсональныхДанных.id_СоставПерсДанных
                                UNION
                                                         select Tab2.id_карточки, ОписаниеКорреспондента, КраткоеСодержание, НомерВходящий, ДатаПоступления, НомерИсход,
                                Основаниепередачи.ОснованиеПередачи from (
						                                  SELECT        Карточка_1.id_карточки, Корреспонденты_1.ОписаниеКорреспондента, Карточка_1.КраткоеСодержание, Карточка_1.НомерВход + '/' + CAST(Карточка_1.номерПП AS nvarchar(10)) AS НомерВходящий, 
                                                                                   CAST(Карточка_1.ДатаПоступ AS date) AS ДатаПоступления, Карточка_1.НомерИсход
                                                          FROM            dbo.Корреспонденты AS Корреспонденты_1 INNER JOIN
                                                                                   dbo.Карточка AS Карточка_1 ON Корреспонденты_1.id_корреспондента = Карточка_1.id_корреспондента INNER JOIN
                                                                                   dbo.ВидПоступленияДокумента ON dbo.ВидПоступленияДокумента.id = Карточка_1.idВидПоступленияДокумента
                                                          WHERE        (Карточка_1.idВидПоступленияДокумента = 6)) as Tab2
						                                  left outer join СвязующаяУчетаПерсональныхДанных
                                on СвязующаяУчетаПерсональныхДанных.id_карточки = Tab2.id_карточки
                                left outer join Основаниепередачи
                                on Основаниепередачи.id_основаниеПередачи = СвязующаяУчетаПерсональныхДанных.id_СоставПерсДанных) as TabCrystall
                                where ДатаПоступления >= '" + dataStart + "' and ДатаПоступления <= '" + dataEnd + "' ";
                                

            return query;
        }
    }
}
