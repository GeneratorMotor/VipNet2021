using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess2021
{
    /// <summary>
    /// SQL скрипт основание передачи персональных данных.
    /// </summary>
    public class QueryОснованиеПередачи : IQueryStringSQL
    {
        private string basTrafic = string.Empty;

        public QueryОснованиеПередачи(string basisTrafic)
        {
            if(string.IsNullOrEmpty(basisTrafic))
            {
                this.basTrafic = basisTrafic.Trim();
            }
            else
            {
                throw new ArgumentNullException("Не указаны основания передачи");
            }
        }

        /// <summary>
        /// Возвращает SQL скрипт для поиска оснований передачи персональных данных.
        /// </summary>
        /// <returns></returns>
        public string Query()
        {
            string query = "SELECT [id_основаниеПередачи] " +
                         ",[ОснованиеПередачи] " +
                         "FROM [Основаниепередачи] where [ОснованиеПередачи] like '%" + this.basTrafic.Trim() + "%'";

            return query;

        }
    }
}
