using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class ОснованиеПередачи
    {
        private int id_основаниеПередачи;
        private string основаниеПередачи;
        private bool flag;

        /// <summary>
        /// id основание передачи.
        /// </summary>
        public int Id_основаниеПередачи
        {
            get
            {
                return id_основаниеПередачи;
            }
            set
            {
                id_основаниеПередачи = value;
            }
        }

        /// <summary>
        ///  Хранит основание передачи.
        /// </summary>
        public string Основание
        {
            get
            {
                return основаниеПередачи;
            }
            set
            {
                основаниеПередачи = value;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public bool FlagSelect
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
                
    }
}
