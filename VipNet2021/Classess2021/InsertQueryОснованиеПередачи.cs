using System;
using System.Collections.Generic;
using System.Text;
using RegKor.Classess;
using System.Collections;

namespace RegKor.Classess2021
{
    /// <summary>
    /// Формирует строку Основание на добавление новых записей в БД.
    /// </summary>
    public class InsertQueryОснованиеПередачи : IQueryStringSQL
    {
        // Переменная для хранения списка оснований к передаче персональных данных.
        private IEnumerable<ОснованиеПередачи> list;

        // id для хранения id карты.
        private int id_card = 0;

        public InsertQueryОснованиеПередачи(IEnumerable<ОснованиеПередачи> listDate, int idInputCard)
        {
            if (listDate == null)
            {
                throw new NullReferenceException("Список онснования передачи персональных данных пуст");
            }

            list = listDate;

            id_card = idInputCard;
        }


        public string Query()
        {
            StringBuilder builder = new StringBuilder();

            if (this.list == null)
            {
                return "";
            }


            foreach (ОснованиеПередачи itm in list)
            {
                string query = @" INSERT INTO СвязующаяУчетаПерсональныхДанных (id_карточки,id_СоставПерсДанных)
                                 values(@id_карточки," + itm.Id_основаниеПередачи + ") ";

                builder.Append(query);
            }

            return builder.ToString();
        }


    }
}
