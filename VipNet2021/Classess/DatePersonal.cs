using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class DatePersonal
    {
        private string list;
        private int _���������������� = 0;
        �������������������� �������;

        /// <summary>
        /// ������ ������ ������������ ������.
        /// </summary>
        public string �����������������������
        {
            get
            {
                return list;
            }
            set
            {
                list = value;
            }
        }

        /// <summary>
        /// ���� ��������� ������������ ������.
        /// </summary>
        public int Id�������������������������������
        {
            get
            {
                return _����������������;
            }
            set
            {
                _���������������� = value;
            }
        }

        /// <summary>
        /// �������� ������� � �������� ���� ������ � �������� ������ ��� ��������.
        /// </summary>
        public �������������������� ��������������������
        {
            get
            {
                return �������;
            }
            set
            {
                ������� = value;
            }
        }


    }
}
