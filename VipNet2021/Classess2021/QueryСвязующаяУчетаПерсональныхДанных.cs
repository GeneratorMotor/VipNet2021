using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess2021
{
    public class QueryСвязующаяУчетаПерсональныхДанных : IQueryStringSQL
    {
        private int idCard = 0;

        public QueryСвязующаяУчетаПерсональныхДанных(int idCard)
        {
            idCard = idCard;
        }

        public string Query()
        {
            string quer = @"select * from Основаниепередачи 
inner join СвязующаяУчетаПерсональныхДанных 
on Основаниепередачи.id_основаниеПередачи = СвязующаяУчетаПерсональныхДанных.id_СоставПерсДанных
where id_карточки = " + idCard + " ";

            return quer;
        }
    }
}
