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
        /// ��������� ����.
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
        /// �������� ����.
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
