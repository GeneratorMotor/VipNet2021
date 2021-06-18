using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class DocumentPerson
    {
        private string fioPerson = string.Empty;

        private int всегоДокументов;
        private int истекшиеДокументы;
        private int неИстекшиеДокументы;

        public string ФиоПолучатель
        {
            get
            {
                return fioPerson;
            }
            set
            {
                fioPerson = value;
            }
        }

        public int ВсегоДокументоНаКонтроле
        {
            get
            {
                return всегоДокументов;
            }
            set
            {
                всегоДокументов = value;
            }
        }

        public int ПросроченныеДокументы
        {
            get
            {
                return истекшиеДокументы;
            }
            set
            {
                истекшиеДокументы = value;
            }
        }

        public int НеПросроченныеДокументы
        {
            get
            {
                return неИстекшиеДокументы;
            }
            set
            {
                неИстекшиеДокументы = value;
            }
        }
    }
}
