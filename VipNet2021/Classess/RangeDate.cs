using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class RangeDate
    {
        private DateTime _dataStart;
        private DateTime _dateEnd;

        /// <summary>
        /// Начальная дата.
        /// </summary>
        public DateTime DataStart
        {
            get
            {
                return _dataStart;
            }
            set
            {
                _dataStart = value;
            }
        }

        /// <summary>
        /// Конечная дата.
        /// </summary>
        public DateTime DataEnd
        {
            get
            {
                return _dateEnd;
            }
            set
            {
                _dateEnd = value;
            }
        }
    }
}
