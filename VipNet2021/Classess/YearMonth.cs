using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Вспомогательный класс для формирования отчета статистики по способу поступления документа.
    /// </summary>
    public class YearMonth
    {
        private int year = 0;
        private int month = 0;

        /// <summary>
        /// Год отчета.
        /// </summary>
        public int Year
        {
            get
            {
                return year;
            }
            set
            {
                year = value;
            }
        }

        /// <summary>
        /// Порядковый номер месяца отчета.
        /// </summary>
        public int NumMonth
        {
            get
            {
                return month;
            }
            set
            {
                month = value;
            }
        }
    }
}
