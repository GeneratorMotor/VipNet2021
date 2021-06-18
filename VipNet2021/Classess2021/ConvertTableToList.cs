using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using RegKor.Classess;

namespace RegKor.Classess2021
{
    /// <summary>
    /// Преобразуем таблицу с основанием передачи в список.
    /// </summary>
    public class ConvertTableToList 
    {
        private DataTable tab;

        private List<ОснованиеПередачи> list;

        public ConvertTableToList(DataTable tab, List<ОснованиеПередачи> list)
        {
            if(tab == null || tab.Rows.Count == 0)
            {
                throw new ArgumentNullException("Таблица основания передачи пустая");
            }

            if (list == null)
            {
                throw new ArgumentNullException("Отсутствует список знчений основания передачи");
            }

            this.tab = tab;
            this.list = list;
        }

        /// <summary>
        /// Возвращает список с основанием передачи.
        /// </summary>
        /// <returns></returns>
        public List<ОснованиеПередачи> Get()
        {
            // Поставим первым списком 475 постановление правительства.
            DataRow[] rows_475 = tab.Select("id_основаниеПередачи = 15");

            if (rows_475 != null && rows_475.Length > 0)
            {
                DataRow row475 = rows_475[0];

                ОснованиеПередачи item = new ОснованиеПередачи();
                item.Id_основаниеПередачи = Convert.ToInt32(row475["id_основаниеПередачи"]);
                item.Основание = row475["ОснованиеПередачи"].ToString().Trim();

                list.Add(item);

            }

            
            foreach (DataRow row in tab.Rows)
            {
                if (Convert.ToInt32(row["id_основаниеПередачи"]) != 15)
                {
                    ОснованиеПередачи item = new ОснованиеПередачи();
                    item.Id_основаниеПередачи = Convert.ToInt32(row["id_основаниеПередачи"]);
                    item.Основание = row["ОснованиеПередачи"].ToString().Trim();

                    list.Add(item);
                }
            }

            return list;
        }

    }
}
