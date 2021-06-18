using System;
using System.Collections.Generic;
using System.Text;
using RegKor.Classess;

namespace RegKor.Classess2021
{
    public class UpdateQueryОснованиеПередачи : IQueryStringSQL
    {
         // Переменная для хранения списка оснований к передаче персональных данных.
        private IEnumerable<ОснованиеПередачи> list;

        // id для хранения id карты.
        private int id_card = 0;

        public UpdateQueryОснованиеПередачи(IEnumerable<ОснованиеПередачи> listDate, int idInputCard)
        {
            if (listDate == null)
            {
                throw new NullReferenceException("Список основания передачи персональных данных пуст");
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

            string delet = @" delete dbo.СвязующаяУчетаПерсональныхДанных
                           where id_карточки = " + id_card + " ";

            builder.Append(delet);

            foreach (ОснованиеПередачи itm in list)
            {
                string query = @" INSERT INTO СвязующаяУчетаПерсональныхДанных (id_карточки,id_СоставПерсДанных)
                                 values( " + id_card + " ," + itm.Id_основаниеПередачи + ") ";

                builder.Append(query);
            }

            return builder.ToString();
        }
    }
}
