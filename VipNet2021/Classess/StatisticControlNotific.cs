using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Класс содержащий данные для отчета Контрольное уведомление.
    /// </summary>
    public class StatisticControlNotific 
    {
        private int countDoc;
        private int countDocLastDate;
        private List<PersonDocument> listPerson;

        public StatisticControlNotific()
        {
            listPerson = new List<PersonDocument>();
        }

        public int ВсегоДокументыНаКонтроле
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

        public List<PersonDocument> СписокИсполнителей
        {
            get
            {
                return listPerson;
            }
            set
            {
                listPerson = value;
            }
        }

    }
}
