using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Вспомагательный класс для хранения информации по идентификации карточки исходящей.
    /// </summary>
    public class NumOutputCardVipNet
    {
        private int id;

        private string номерПорядковый = string.Empty;

        /// <summary>
        /// ID карточки
        /// </summary>
        public int Id
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

        public string НомерПорядковый
        {
            get
            {
                return номерПорядковый;
            }
            set
            {
                номерПорядковый = value;
            }
        }
    }
}
