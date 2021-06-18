using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
   public class ДатаSQL
    {
            /// <summary>
            /// Преобразует дату в формат SQL
            /// </summary>
            /// <param name="дата">дата</param>
            /// <returns></returns>
            public static string Дата(string дата)
            {
                //string BeginDateSQL = System.Text.RegularExpressions.Regex.Replace(дата, "\\b(?<day>\\d{1,2}).(?<month>\\d{1,2}).(?<year>\\d{2,4})\\b", "${month}-${day}-${year}");
                string BeginDateSQL = System.Text.RegularExpressions.Regex.Replace(дата, "\\b(?<day>\\d{1,2}).(?<month>\\d{1,2}).(?<year>\\d{2,4})\\b", "${year}${month}${day}");

                return BeginDateSQL;
            }

            /// <summary>
            /// Преобразует дату в формае SQL в дату в формате dd.mm.gggg.
            /// </summary>
            /// <param name="дата">дата</param>
            /// <returns></returns>
            public static string SqlToДата(string дата)
            {

                // Проверим длинну троки.
                string[] arraySatring = new string[3];

                arraySatring[0] = дата.Substring(0, 4);

                arraySatring[1] = дата.Substring(4, 2);

                arraySatring[2] = дата.Substring(6, 2);

                StringBuilder dataBuild = new StringBuilder();
                dataBuild.Append(arraySatring[2] + ".");
                dataBuild.Append(arraySatring[1] + ".");
                dataBuild.Append(arraySatring[0]);

                return dataBuild.ToString();





                //string BeginDateSQL = System.Text.RegularExpressions.Regex.Replace(дата, "\\b(?<day>\\d{1,2})(?<month>\\d{1,2})(?<year>\\d{2,4})\\b", "${day}.${month}.${year}");
                //string BeginDateSQL = System.Text.RegularExpressions.Regex.Replace(дата, "\\b(?<day>\\d{1,2}).(?<month>\\d{1,2}).(?<year>\\d{2,4})\\b", "${year}${month}${day}");

            }

    }
}
