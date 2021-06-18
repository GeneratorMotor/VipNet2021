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

        // ������� � �� ������������� �����������.
        private DataRow[] dtNotOverDoc;

        // ������� � ������������� �����������.
        private DataRow[] dtOverDoc;

        // ������� � ����������� �� ��������.
        private DataRow[] dtDocControl;

        public int ������������������������
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

        public int ��������������������������������
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

        public int ����������������������������������
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
        /// ������� � �� ������������� �����������.
        /// </summary>
        public DataRow[] �����������������������
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
        /// ������� � ������������� �����������.
        /// </summary>
        public DataRow[] ���������������������
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

        public DataRow[] �������������������
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
