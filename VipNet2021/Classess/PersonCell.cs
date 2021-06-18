using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Содержит данные по количество отписанных документов для текущего начальника.
    /// </summary>
    public class PersonCell
    {
        private string name = string.Empty;

        /// <summary>
        /// ФИО начальника которому отписан документ.
        /// </summary>
        public string Name 
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        private List<string> list = new List<string>();

        /// <summary>
        /// Список способов получения документов.
        /// </summary>
        public List<string> СпособПолучДокумент
        {
            get
            {
                return list;
            }
            set
            {
                list = value;
            }
        }


    }
}
