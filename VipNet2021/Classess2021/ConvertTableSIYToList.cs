using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace RegKor.Classess2021
{
    /// <summary>
    /// Конвертор таблицы DataTable в список List<IntemЖурналУчетаСЭУ>.
    /// </summary>
    class ConvertTableSIYToList : IConvertTableToList<IntemЖурналУчетаСЭУ>
    {
        private DataTable tab;

        private List<IntemЖурналУчетаСЭУ> list;

        public ConvertTableSIYToList(DataTable tableInput)
        {
            if (tableInput == null || tableInput.Rows.Count == 0)
                throw new ArgumentNullException("Таблица с исходными данными не может быть пустой","Ошибка преобразования данныэ по СЭУ в список для отчета");
            
            tab = tableInput;

            list = new List<IntemЖурналУчетаСЭУ>();
        }

        /// <summary>
        /// Преобразует таблицу в список.
        /// </summary>
        /// <returns></returns>
        public List<IntemЖурналУчетаСЭУ> GetList()
        {
            foreach (DataRow row in tab.Rows)
            {
                IntemЖурналУчетаСЭУ item = new IntemЖурналУчетаСЭУ();
                //item.Id_карточки = row["id_карточки"].ToString().Trim();

                //calEventDTO.endTime = (DateTime?)(Convert.IsDBNull(row["endTime"]) ? null : row["endTime"]);

                if (Convert.IsDBNull(row["ДатаПоступления"]) != null)
                {
                    item.ДатаПоступления = row["ДатаПоступления"].ToString().Trim();
                }
                else
                {
                    item.ДатаПоступления = "";
                }

                if (Convert.IsDBNull(row["КраткоеСодержание"]) != null)
                {
                    item.КраткоеСодержание = row["КраткоеСодержание"].ToString().Trim();
                }
                else
                {
                    item.КраткоеСодержание = "";
                }

                if (Convert.IsDBNull(row["НомерВходящий"]) != null)
                {
                    item.НомерВходящий = row["НомерВходящий"].ToString().Trim();
                }
                else
                {
                    item.НомерВходящий = "";
                }

                if (Convert.IsDBNull(row["НомерИсход"]) != null)
                {
                    item.НомерИсход = row["НомерИсход"].ToString().Trim();
                }
                else
                {
                    item.НомерИсход = "";
                }

                if (Convert.IsDBNull(row["НомерИсход"]) != null)
                {
                    item.ОписаниеКорреспондента = row["ОписаниеКорреспондента"].ToString().Trim();
                }
                else
                {
                    item.ОписаниеКорреспондента = "";
                }

                if (Convert.IsDBNull(row["ОснованиеПередачи"]) != null)
                {
                    item.ОснованиеПередачи = row["ОснованиеПередачи"].ToString().Trim();
                }
                else
                {
                    item.ОснованиеПередачи = "";
                }

                list.Add(item);
            }

            return list;
        }
    }
}
