using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// ����� ��������� �����������.
    /// </summary>
    public class PersonRecepient
    {
        int id;
        string famili;

        /// <summary>
        ///  ID ���������� ���������.
        /// </summary>
        public int ID
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
        /// ��� ���������� ���������.
        /// </summary>
        public string Famili
        {
            get
            {
                return famili;
            }
            set
            {
                famili = value;
            }
        }
    }
}
