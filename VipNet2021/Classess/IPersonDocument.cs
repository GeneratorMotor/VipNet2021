using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public interface IPersonDocument
    {
        private int countDoc;
        private int countDocLastDate;

        int ������������������������
        {
            get
            {
                return countDoc;
            }
            set
            {
                countDoc = value;
            }
        }

        public int ��������������������������������
        {
            get
            {
                return countDocLastDate;
            }
            set
            {
                countDocLastDate = value;
            }
        }

    }
}
