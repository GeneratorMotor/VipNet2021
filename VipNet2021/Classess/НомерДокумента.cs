using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class НомерДокумента
    {
        private int номер;

        /// <summary>
        /// Хранит номер документа
        /// </summary>
        public int Номер
        {
            get
            {
                return номер;
            }
            set
            {
                номер = value;
            }
        }

        private string префикс;

        /// <summary>
        /// Хранит префикс номера документа.
        /// </summary>
        public string Префикс
        {
            get
            {
                return префикс;
            }
            set
            {
                префикс = value;
            }
        }

        public bool flagUpdate;

        /// <summary>
        /// Флаг обновления записи.
        /// </summary>
        public bool FlagUpdate
        {
            get
            {
                return flagUpdate;
            }
            set
            {
                flagUpdate = value;
            }
        }

        private string полныйНомерДокумента = string.Empty;

        public string ПолныйНомерДокумента
        {
            get
            {
                if (FlagUpdate == false)
                {
                    полныйНомерДокумента = номер.ToString().Trim() + "/" + префикс.Trim();
                    return полныйНомерДокумента;
                }
                else
                {
                    полныйНомерДокумента = префикс.Trim();
                    return полныйНомерДокумента;
                }
            }
            
        }
               
    }
}
