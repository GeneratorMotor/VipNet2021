using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    ///  ласс дл€ хранени€ номеров документов наход€щихс€ на контроле.
    /// </summary>
    public class ItemDoc
    {
        private int id;
        private string номерƒокумента = string.Empty;

        /// <summary>
        /// ’ранит id карточки.
        /// </summary>
        public int id_карточки
        {
            get
            {
                return id;
            }
            set
            {
                id = value;
            }
        }

        /// <summary>
        /// ¬ход€щий номер документа.
        /// </summary>
        public string Ќомер¬ход
        {
            get
            {
                return номерƒокумента;
            }
            set
            {
                номерƒокумента = value;
            }
        }
    
    }
}
