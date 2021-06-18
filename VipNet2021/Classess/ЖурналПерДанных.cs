using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class ЖурналПерДанных
    {
        private string описаниеКорреспондента = string.Empty;
        private string краткоеСодержание = string.Empty;
        private string номерИсходящий = string.Empty;
        private string дата = string.Empty;
        private string датаПоступления = string.Empty;
        private string номерВходящий = string.Empty;
        private int id_cardOutput = 0;


        private int id = 0;

        /// <summary>
        /// Хранит Id карточки.
        /// </summary>
        public int Id
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
        /// Хранит id исходящей карточки.
        /// </summary>
        public int IdКарточкаИсходящая 
        {
            get
            {
                return id_cardOutput;
            }
            set
            {
                id_cardOutput = value;
            }
        }
    
            
                

        /// <summary>
        /// Описание корреспондента.
        /// </summary>
        public string ОписаниеКорреспондента
        {
            get
            {
                return описаниеКорреспондента;
            }
            set
            {
                описаниеКорреспондента = value;
            }
        }

        /// <summary>
        /// Краткое содержание.
        /// </summary>
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

        /// <summary>
        /// № исходящий.
        /// </summary>
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

        /// <summary>
        /// Дата отправки документа.
        /// </summary>
        public string ДатаОтправки
        {
            get
            {
                return дата;
            }
            set
            {
                дата = value;
            }
        }

        /// <summary>
        /// Дата поступления документов.
        /// </summary>
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

        /// <summary>
        /// Номер входящий.
        /// </summary>
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


    }
}
