using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace RegKor.Classess
{
    class PersonDataDefault
    {
        private int id_карточка»сход€ща€ = 0;
        public PersonDataDefault(int id арточки)
        {
            id_карточка»сход€ща€ = id арточки;
        }

        /// <summary>
        /// ¬озвращает список оснований передачи данных установленных формой по умолчанию.
        /// </summary>
        /// <returns></returns>
        public List<ќснованиеѕередачи> GetList()
        {
            string quer = "select * from ќснованиепередачи " +
                          "where id_основаниеѕередачи in ( " +
                          "SELECT [id_ќснованиеѕередачи] " +
                          "FROM [—в€зующа€÷ельѕолучениперсональныхƒанных] " +
                          "where id_карточки = " + id_карточка»сход€ща€ + " )";

            GetDataTable getTabC = new GetDataTable(quer);
            DataTable tabContr = getTabC.DataTable("ќснованиеѕередачи онтрольное");

            List<ќснованиеѕередачи> list = new List<ќснованиеѕередачи>();
            foreach (DataRow row in tabContr.Rows)
            {
                ќснованиеѕередачи item = new ќснованиеѕередачи();
                item.Id_основаниеѕередачи = Convert.ToInt32(row["id_основаниеѕередачи"]);
                item.ќснование = row["ќснованиеѕередачи"].ToString().Trim();
                item.FlagSelect = true;
               
                list.Add(item);
            }

            return list;
        }
                
    }
}
