using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess2021
{
    public class IntemЖурналУчетаСЭУ
    {
        private int id_карточки = 0;
        private string описаниеКорреспондента = string.Empty;
        private string краткоеСодержание = string.Empty;
        private string номерВходящий = string.Empty;
        private string датаПоступления = string.Empty;
        private string номерИсход = string.Empty;
        private string основаниеПередачи = string.Empty;

        /// <summary>
        /// id карточки.
        /// </summary>
        public int Id_карточки
        {
            get
            {
                return id_карточки;
            }
            set
            {
                id_карточки = value;
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

        /// <summary>
        /// Дата поступления.
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
        /// Дата поступления.
        /// </summary>
        public string НомерИсход
        {
            get
            {
                return номерИсход;
            }
            set
            {
                номерИсход = value;
            }
        }

        /// <summary>
        /// Основание передачи.
        /// </summary>
        public string ОснованиеПередачи
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


    }
}
