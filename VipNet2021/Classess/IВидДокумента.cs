using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class ВидДокумента
    {
        private string видДок = string.Empty;

        public string НазваниеВидаДокумента
        {
            get
            {
                return видДок;
            }
            set
            {
                видДок = value;
            }
        }
    }
}
