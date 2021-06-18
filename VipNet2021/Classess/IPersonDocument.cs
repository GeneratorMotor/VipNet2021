using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public interface IPersonDocument
    {
        private int countDoc;
        private int countDocLastDate;

        int ВсегоДокументыНаКонтроле
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

        public int КоличествоПросроченныхДокументов
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
