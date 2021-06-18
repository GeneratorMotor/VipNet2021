using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// �������� ������ �� ���������� ���������� ���������� ��� �������� ����������.
    /// </summary>
    public class PersonCell
    {
        private string name = string.Empty;

        /// <summary>
        /// ��� ���������� �������� ������� ��������.
        /// </summary>
        public string Name 
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }

        private List<string> list = new List<string>();

        /// <summary>
        /// ������ �������� ��������� ����������.
        /// </summary>
        public List<string> �������������������
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


    }
}
