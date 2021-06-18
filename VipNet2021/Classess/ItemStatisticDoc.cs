using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class ItemStatisticDoc
    {
        private string фио = string.Empty;
        private string видПоступления = string.Empty;
        private int count = 0;


        public string ФИО
        {
            get
            {
                return фио;
            }
            set
            {
                фио = value;
            }
        }

        public string ВидПоступления
        {
            get
            {
                return видПоступления;
            }
            set
            {
                видПоступления = value;
            }
        }

        public int Count
        {
            get
            {
                return count;
            }
            set
            {
                count = value;
            }
        }



    }
}
