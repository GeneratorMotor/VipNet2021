using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// ��������������� ����� ����������� ��������� ������.
    /// </summary>
    public class StatisticDocInput
    {
        private string id = string.Empty;
        private string _�������������������������� = string.Empty;
        private int ������������������������ = 0;
        private int? ���������������� = 0;
        private int? email = 0;
        private int? vipNet = 0;
        private int? fax = 0;
        private string ����������� = string.Empty;


        public string Num
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

        public string ��������������������������
        {
            get
            {
                return _��������������������������;
            }
            set
            {
                _�������������������������� = value;
            }
        }

        public int ������������������������
        {
            get
            {
                return ������������������������;
            }
            set
            {
                ������������������������ = value;
            }
        }

        public int? ����������������
        {
            get
            {
                return ����������������;
            }
            set
            {
                ���������������� = value;
            }
        }

        public int? Email
        {
            get
            {
                return email;
            }
            set
            {
                email = value;
            }
        }

        public int? VipNet
        {
            get
            {
                return vipNet;
            }
            set
            {
                vipNet = value;
            }
        }

        public int? Fax
        {
            get
            {
                return fax;
            }
            set
            {
                fax = value;
            }
        }

        public string �����������
        {
            get
            {
                return �����������;
            }
            set
            {
                ����������� = value;
            }
        }
    }
}
