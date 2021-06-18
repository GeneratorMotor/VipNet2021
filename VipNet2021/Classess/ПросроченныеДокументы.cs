using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class ПросроченныеДокументы
    {
        private string номерПП = string.Empty;
        private string исполнитель = string.Empty;
        private string датаПоступления = string.Empty;
        private string номерВходящий = string.Empty;
        private string срокВыполнения = string.Empty;

        public string НомерПП
        {
            get
            {
                return номерПП;
            }
            set
            {
                номерПП = value;
            }
        }

        public string ОтветственныйИсполнитель
        {
            get
            {
                return исполнитель;
            }
            set
            {
                исполнитель = value;
            }
        }

        public string ДатаПоступления
        {
            get
            {
                return датаПоступления;
            }
            set
            {
                датаПоступления = value;
            }
        }

        public string НомерВходящий
        {
            get
            {
                return номерВходящий;
            }
            set
            {
                номерВходящий = value;
            }
        }

        public string СрокВыполнения
        {
            get
            {
                return срокВыполнения;
            }
            set
            {
                срокВыполнения = value;
            }
        }
    }
}
