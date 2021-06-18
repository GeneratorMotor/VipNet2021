using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    /// <summary>
    /// ������� �� ������ �����.
    /// </summary>
    class PrintReportStaticOutputDoc : IPrintReport
    {
        private Report����������������������������� _report;

        private List<��������������������������> _list;

        /// <summary>
        /// ������ � ������� ��� ������.
        /// </summary>
        public List<��������������������������> ListDate
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

        public PrintReportStaticOutputDoc(Report����������������������������� report)
        {
            if(report != null)
            {
                _report = report;
               
            }
            else
            {
                throw new Exception("����������� ������ ��� ������");
            }
        }

        public void Execute()
        {
            _report.PrintReport(_list);
        }

    }
}
