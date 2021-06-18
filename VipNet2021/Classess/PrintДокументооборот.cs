using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// Класс описиывающий строку отчёта.
    /// </summary>
    public class PrintДокументооборот
    {
          private string _РегистрационныйHомерДокумента;
        private string _ДатаПоступ;
          private string _КраткоеСодержание;
          private string _NameFileDocument;
          private string _GuidName;
          private string _ОписаниеКорреспондента;
        private string _СрокВыполнения;
        private string _ОтметкаПрочтение;
        private string _РезультатВыполнения;
          private string _ОписаниеПолучателя;

        public string РегистрационныйHомерДокумента
        {
            get
            {
                return _РегистрационныйHомерДокумента;
            }
            set
            {
                _РегистрационныйHомерДокумента = value;
            }
        }


        public string ДатаПоступ
        {
            get
            {
                return _ДатаПоступ;
            }
            set
            {
                _ДатаПоступ = value;
            }
        }

        public string КраткоеСодержание
        {
            get
            {
                return _КраткоеСодержание;
            }
            set
            {
                _КраткоеСодержание = value;
            }
        }

        public string NameFileDocument
        {
            get
            {
                return _NameFileDocument;
            }
            set
            {
                _NameFileDocument = value;
            }
        }

        public string GuidName
        {
            get
            {
                return _GuidName;
            }
            set
            {
                _GuidName = value;
            }
        }

        public string ОписаниеКорреспондента
        {
            get
            {
                return _ОписаниеКорреспондента;
            }
            set
            {
                _ОписаниеКорреспондента = value;
            }
        }

        public string СрокВыполнения
        {
            get
            {
                return _СрокВыполнения;
            }
            set
            {
                _СрокВыполнения = value;
            }
        }

        public string ОтметкаПрочтение
        {
            get
            {
                return _ОтметкаПрочтение;
            }
            set
            {
                _ОтметкаПрочтение = value;
            }
        }

        public string РезультатВыполнения
        {
            get
            {
                return _РезультатВыполнения;
            }
            set
            {
                _РезультатВыполнения = value;
            }
        }

        public string ОписаниеПолучателя
        {
            get
            {
                return _ОписаниеПолучателя;
            }
            set
            {
                _ОписаниеПолучателя = value;
            }
        }

    }
}
