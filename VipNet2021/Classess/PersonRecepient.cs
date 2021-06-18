using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Класс описывает получателей.
    /// </summary>
    public class PersonRecepient
    {
        int id;
        string famili;

        /// <summary>
        ///  ID получателя документа.
        /// </summary>
        public int ID
        {
            get
            {
                return id;
            }
            set
            {
                id = value;
            }
        }

        /// <summary>
        /// ФИО получателя документа.
        /// </summary>
        public string Famili
        {
            get
            {
                return famili;
            }
            set
            {
                famili = value;
            }
        }
    }
}
