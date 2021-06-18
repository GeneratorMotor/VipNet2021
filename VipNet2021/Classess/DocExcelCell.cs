using System;
using System.Collections.Generic;
using System.Text;


namespace RegKor.Classess
{
    /// <summary>
    /// ¬спомогательный класс содержащий €чейку с данными дл€ отчета.
    /// </summary>
    public class DocExcelCell
    {
        private string valueDate = string.Empty;
        private int countColumn = 0;
        private bool flagEdit = false;

        /// <summary>
        /// ’ранит значение €чейки.
        /// </summary>
        public string ValueCell
        {
            get
            {
                return valueDate;
            }
            set
            {
                valueDate = value;
            }
        }

        /// <summary>
        ///  оличество столбцов которые занимает €чейка с данными.
        /// </summary>
        public int CountColumn
        {
            get
            {
                return countColumn;
            }
            set
            {
                countColumn = value;
            }
        }

        /// <summary>
        /// ‘лаг указывает, что в €чейку установлены данные (значение - true).
        /// </summary>
        public bool FlagEdit
        {
            get
            {
                return flagEdit;
            }
            set
            {
                flagEdit = value;
            }
        }

        private string month = string.Empty;

        public string Month
        {
            get
            {
                return month;
            }
            set
            {
                month = value;
            }
        }
    }
}
