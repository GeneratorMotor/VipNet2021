using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class ItemСпособПоступленияДокумента : ItemAbstractПоступлениеДокументов
    {
        private int id;
        private string processName = string.Empty;

        // Хранит id 
        public int Id
        {
            get
            {
                return id;
            }
            set
            {
                id= value;
            }
        }

        /// <summary>
        /// Название спопба передачи документа.
        /// </summary>
        public string ProcessName
        {
            get
            {
                return processName;
            }
            set
            {
                processName = value;
            }
        }


    }
}
