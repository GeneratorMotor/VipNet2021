using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;
using RegKor.Classess2021;


namespace RegKor
{
    public partial class FormSelectInputDatePerson : Form
    {
        public FormSelectInputDatePerson()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Установим дату начала отчета.
            string dateStart = ДатаSQL.Дата(this.dateTimePicker1.Value.ToShortDateString());

            // Установим дату окончания отчета.
            string dateEnd = ДатаSQL.Дата(this.dateTimePicker2.Value.ToShortDateString());

            // Создадим SQL скрипт к отчету.
            IQueryStringSQL selectQuery = new QuerySelectInputDocSIY(dateStart.Trim(), dateEnd.Trim());

            string query = selectQuery.Query();

            ПодключитьБД connectBD = new ПодключитьБД();

            GetDataTable table = new GetDataTable(query);
            DataTable tabInputCardСэу = table.DataTable();

            // Преобразуем полученную таблицу в список.
            ConvertTableSIYToList convaertTable = new ConvertTableSIYToList(tabInputCardСэу);
            List<IntemЖурналУчетаСЭУ> list = convaertTable.GetList();

            // Сформируем документ.
            IPrintReport reportInputCardSiy = new ЖурналУчетаСЭУ(list, dateStart, dateEnd);

            // Выведим на печать.
            PrintReport print = new PrintReport();

            print.SetCommand(reportInputCardSiy);
            print.Execute();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}