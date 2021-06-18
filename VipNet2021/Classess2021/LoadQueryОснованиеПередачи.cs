using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess2021
{
    /// <summary>
    /// Возвращает SQL запрос на основание передачи персональных данных для карточки входящих.
    /// </summary>
    public class LoadQueryОснованиеПередачи : IQueryStringSQL
    {
        private int id_карточки = 0;

        public LoadQueryОснованиеПередачи(int idКарточки)
        {
            id_карточки = idКарточки;
        }

        /// <summary>
        /// Возвращает sql запрос.
        /// </summary>
        /// <returns></returns>
        public string Query()
        {
            string query = @" select id_основаниеПередачи,ОснованиеПередачи from Основаниепередачи
                          inner join СвязующаяУчетаПерсональныхДанных
                          on СвязующаяУчетаПерсональныхДанных.id_СоставПерсДанных = Основаниепередачи.id_основаниеПередачи
                          where id_карточки = " + this.id_карточки + " ";

            return query;
        }
    }
}
