using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace RegKor.Classess
{
    class ReportНомераИсходящихДокументов :IPrintReport
    {
        private ReportОтчетИсходящиеДокументы _report;

        private  DataGridViewRowCollection _list;

        /// <summary>
        /// Список с данными для отчета.
        /// </summary>
        public DataGridViewRowCollection ListDate
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

        public ReportНомераИсходящихДокументов(ReportОтчетИсходящиеДокументы report)
        {
            if(report != null)
            {
                _report = report;
               
            }
            else
            {
                throw new Exception("Отсутствуют данные для отчета");
            }
        }

        public void Execute()
        {
            _report.PrintReportStaticOutputDoc(_list);
        }
    }
}
