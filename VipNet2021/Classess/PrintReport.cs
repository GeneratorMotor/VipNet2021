using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public class PrintReport
    {
        private IPrintReport _print;

        public void SetCommand(IPrintReport print)
        {
            if (print != null)
            {
                _print = print;
            }
            else
            {
                throw new Exception("Нет данных для отчета");
            }
        }

        public void Execute()
        {
            _print.Execute();
        }
    }
}
