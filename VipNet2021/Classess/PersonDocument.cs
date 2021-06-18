using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;


namespace RegKor.Classess
{
    public class PersonDocument
    {
        private int countDoc;
        private int countDocLastDate;
        private int countDocNoOverDate;
        private string fio = string.Empty;

        // Таблица с не просроченными документами.
        private DataRow[] dtNotOverDoc;

        // Таблица с просроченными документами.
        private DataRow[] dtOverDoc;

        // Таблица с документами на контроле.
        private DataRow[] dtDocControl;

        public int ВсегоДокументыНаКонтроле
        {
            get
            {
                return countDoc;
            }
            set
            {
                countDoc = value;
            }
        }

        public int КоличествоПросроченныхДокументов
        {
            get
            {
                return countDocLastDate;
            }
            set
            {
                countDocLastDate = value;
            }
        }

        public int КоличествоНеПросроченныхДокументов
        {
            get
            {
                return countDocNoOverDate;
            }
            set
            {
                countDocNoOverDate = value;
            }
        }

        public string FioPerson
        {
            get
            {
                return fio;
            }
            set
            {
                fio = value;
            }
        }

        /// <summary>
        /// Таблица с не просроченными документами.
        /// </summary>
        public DataRow[] НеПрсороченныеДокументы
        {
            get
            {
                return dtNotOverDoc;
            }
            set
            {
                dtNotOverDoc = value;
            }
        }

        /// <summary>
        /// Таблица с просроченными документами.
        /// </summary>
        public DataRow[] ПросроченныеДокументы
        {
            get
            {
                return dtOverDoc;
            }
            set
            {
                dtOverDoc = value;
            }
        }

        public DataRow[] ДокументыНаКонтроле
        {
            get
            {
                return dtDocControl;
            }
            set
            {
                dtDocControl = value;
            }
        }

        



    }
}
