using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public static class Montch
    {
        public static string GetMonth(int numMonth)
        {
            switch (numMonth)
            {
                case 1: return "������";
                case 2: return "�������";
                case 3: return "����";
                case 4: return "������";
                case 5: return "���";
                case 6: return "����";
                case 7: return "����";
                case 8: return "������";
                case 9: return "��������";
                case 10: return "�������";
                case 11: return "������";
                case 12: return "�������";
                case 13: return "����� :";
            }

            return string.Empty;
        }

        public static int GetNumMonth(string month)
        {
            switch (month)
            {
                case "������": return 1;
                case "�������": return 2;
                case "����": return 3;
                case "������": return 4;
                case "���": return 5;
                case "����": return 6;
                case "����": return 7;
                case "������": return 8;
                case "��������": return 9;
                case "�������": return 10;
                case "������": return 11;
                case "�������": return 12;
            }

            return 0;
        }
    }
}
