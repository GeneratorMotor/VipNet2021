using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Вспомогательный класс описывающий структуру отчета.
    /// </summary>
    public class StatisticDocInput
    {
        private string id = string.Empty;
        private string _НаименованиеКорреспондента = string.Empty;
        private int колвоВходКорреспонденции = 0;
        private int? бумажныйНоситель = 0;
        private int? email = 0;
        private int? vipNet = 0;
        private int? fax = 0;
        private string исполнитель = string.Empty;


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

        public string НаименованиеКорреспондента
        {
            get
            {
                return _НаименованиеКорреспондента;
            }
            set
            {
                _НаименованиеКорреспондента = value;
            }
        }

        public int КолвоВходКорреспонденции
        {
            get
            {
                return колвоВходКорреспонденции;
            }
            set
            {
                колвоВходКорреспонденции = value;
            }
        }

        public int? БумажныйНоситель
        {
            get
            {
                return бумажныйНоситель;
            }
            set
            {
                бумажныйНоситель = value;
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

        public string Исполнитель
        {
            get
            {
                return исполнитель;
            }
            set
            {
                исполнитель = value;
            }
        }
    }
}
