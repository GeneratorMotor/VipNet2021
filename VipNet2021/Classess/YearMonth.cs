using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// ��������������� ����� ��� ������������ ������ ���������� �� ������� ����������� ���������.
    /// </summary>
    public class YearMonth
    {
        private int year = 0;
        private int month = 0;

        /// <summary>
        /// ��� ������.
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
        /// ���������� ����� ������ ������.
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
