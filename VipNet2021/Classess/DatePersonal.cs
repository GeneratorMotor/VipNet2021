using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class DatePersonal
    {
        private string list;
        private int _цельѕерсонƒанных = 0;
        ќтметкаѕередачаќтказ отметка;

        /// <summary>
        /// ’ранит состав персональных данных.
        /// </summary>
        public string —отавѕерсональныхƒанных
        {
            get
            {
                return list;
            }
            set
            {
                list = value;
            }
        }

        /// <summary>
        /// ÷ель получени€ персональных данных.
        /// </summary>
        public int Id÷ельѕолучени€ѕерсональныхƒанных
        {
            get
            {
                return _цельѕерсонƒанных;
            }
            set
            {
                _цельѕерсонƒанных = value;
            }
        }

        /// <summary>
        /// —одержит отметку о передаче перс данных и описание почему его отказали.
        /// </summary>
        public ќтметкаѕередачаќтказ ќтметкаќтказѕередача
        {
            get
            {
                return отметка;
            }
            set
            {
                отметка = value;
            }
        }


    }
}
