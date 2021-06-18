using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace RegKor.Classess
{
    public class СтрокаОтчёта
    {
        private DataTable table;
        private int leng = 0;
        public СтрокаОтчёта(DataTable tab)
        {
            table = tab;
        }

        /// <summary>
        /// Конвертирует поле таблицы в строку.
        /// </summary>
        /// <returns></returns>
        public string ConvertStringBuilder()
        {
            StringBuilder build = new StringBuilder();

            foreach (DataRow row in table.Rows)
            {
                string sItem = row[0].ToString().Trim();
                //build.Append(sItem + ",\n");
                build.Append(sItem + "\n");

                // Удалим последний символ.
                leng = build.Length;
                 
            }

            return build.ToString().Trim();//.Remove(leng-3,3);
        }
    }
}
