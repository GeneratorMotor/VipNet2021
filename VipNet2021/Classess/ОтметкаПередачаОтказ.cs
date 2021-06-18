using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class ОтметкаПередачаОтказ
    {
        private bool flag = false;
        private string причины;

        /// <summary>
        /// Содержит отметку об отказе или положительном решении.
        /// </summary>
        public bool Отметка
        {
            get
            {
                return flag;
            }
            set
            {
                flag = value;
            }
        }

        /// <summary>
        /// Хранит причину отказа.
        /// </summary>
        public string ПричиныОтказа
        {
            get
            {
                return причины;
            }
            set
            {
                причины = value;
            }
        }
    }
}
