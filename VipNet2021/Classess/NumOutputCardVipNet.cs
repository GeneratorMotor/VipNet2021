using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// ��������������� ����� ��� �������� ���������� �� ������������� �������� ���������.
    /// </summary>
    public class NumOutputCardVipNet
    {
        private int id;

        private string ��������������� = string.Empty;

        /// <summary>
        /// ID ��������
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

        public string ���������������
        {
            get
            {
                return ���������������;
            }
            set
            {
                ��������������� = value;
            }
        }
    }
}
