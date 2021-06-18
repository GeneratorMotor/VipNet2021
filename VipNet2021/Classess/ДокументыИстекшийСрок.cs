using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class ДокументыИстекшийСрок : ПросроченныеДокументы
    {
        private string корреспонденты;
        public string Корреспонденты
        {
            get
            {
                return корреспонденты;
            }
            set
            {
                корреспонденты = value;
            }
        }

        private string краткоеСодержание;
        public string КраткоеСодержание
        {
            get
            {
                return краткоеСодержание;
            }
            set
            {
                краткоеСодержание = value;
            }
        }

        private string датаИсходящий;
        public string ДатаИсходящая
        {
            get
            {
                return датаИсходящий;
            }
            set
            {
                датаИсходящий = value;
            }
        }

        private string номерИсходящий;
        public string НомерИсходящий
        {
            get
            {
                return номерИсходящий;
            }
            set
            {
                номерИсходящий = value;
            }
        }


    }
}
