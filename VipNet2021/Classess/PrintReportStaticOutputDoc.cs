using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// ¬ыводит на печать отчет.
    /// </summary>
    class PrintReportStaticOutputDoc : IPrintReport
    {
        private Report—татистика»сход€щихƒокументов _report;

        private List<—татистика¬ход»сполнителей> _list;

        /// <summary>
        /// —писок с данными дл€ отчета.
        /// </summary>
        public List<—татистика¬ход»сполнителей> ListDate
        {
            get
            {
                return _list;
            }
            set
            {
                _list = value;
            }
        }

        public PrintReportStaticOutputDoc(Report—татистика»сход€щихƒокументов report)
        {
            if(report != null)
            {
                _report = report;
               
            }
            else
            {
                throw new Exception("ќтсутствуют данные дл€ отчета");
            }
        }

        public void Execute()
        {
            _report.PrintReport(_list);
        }

    }
}
