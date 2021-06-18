using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// ����� ���������� ������ ��� ������ ����������� �����������.
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

        public int ������������������������
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

        public List<PersonDocument> ������������������
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
