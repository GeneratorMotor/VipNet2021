using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// ����� ��� �������� ������� ���������� ����������� �� ��������.
    /// </summary>
    public class ItemDoc
    {
        private int id;
        private string �������������� = string.Empty;

        /// <summary>
        /// ������ id ��������.
        /// </summary>
        public int id_��������
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
        /// �������� ����� ���������.
        /// </summary>
        public string ���������
        {
            get
            {
                return ��������������;
            }
            set
            {
                �������������� = value;
            }
        }
    
    }
}
