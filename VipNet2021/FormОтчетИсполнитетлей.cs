using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;

using Microsoft.Office.Interop.Excel;

namespace RegKor
{
    public partial class FormОтчетИсполнитетлей : Form
    {

        // Переменная для хранения выбранного года.
        private string year = string.Empty;

        private List<string> listMonthStart;
        private List<string> listMonthEnd;



        /// <summary>
        /// Выбранный год.
        /// </summary>
        public string YearSelect
        {
            get
            {
                return year;
            }
            set
            {
                year = value;
            }
        }

        public FormОтчетИсполнитетлей()
        {
            InitializeComponent();

            listMonthStart = new List<string>();

            listMonthEnd = new List<string>();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // Получим номер первого месяца.
            int startMonth = Montch.GetNumMonth(this.cmbStart.Text);

            int endMonth = Montch.GetNumMonth(this.cmbEnd.Text);

            if (endMonth < startMonth)
            {
                MessageBox.Show("Месяц окончания периода указан раньше начала периода");
                return;
            }

            // Список номеров месяцев.
            List<int> listNum = new List<int>();

            // Получим список номеров выбранных месяцев.
            for (int ms = startMonth; ms <= endMonth; ms++)
            {
                listNum.Add(ms);
            }

            // Список начальников отделов и управлений.
            СписокНачальников listDirect = new СписокНачальников();

            // Список спсобов поступления документов.
            ProcessGetDoc processDoc = new ProcessGetDoc();

            ExcelGenerate excel = new ExcelGenerate(listDirect, processDoc);
            excel.StartMonth = startMonth;
            excel.EndMonth = endMonth;
            excel.Year = Convert.ToInt16(this.lblYear.Text.Trim());

            // Список с названиями месяцев.
            List<YearMonth> listYM = new List<YearMonth>();

            foreach (int numMontch in listNum)
            {
                YearMonth ym1 = new YearMonth();
                ym1.Year = Convert.ToInt16(this.lblYear.Text.Trim());
                ym1.NumMonth = numMontch;
                listYM.Add(ym1);
            }

           

            excel.CreateExcel(listYM, 4);


        }

        private void FormОтчетИсполнитетлей_Load(object sender, EventArgs e)
        {
            int year = Convert.ToInt16(this.YearSelect) + 1;
            this.lblYear.Text = year.ToString().Trim();

            // Список для хранения названия месяцев.
            //List<string> listMonth = new List<string>();

            // Заполним список названий месяцев.
            listMonthStart.Add("Январь");
            listMonthStart.Add("Февраль");
            listMonthStart.Add("Март");
            listMonthStart.Add("Апрель");
            listMonthStart.Add("Май");
            listMonthStart.Add("Июнь");
            listMonthStart.Add("Июль");
            listMonthStart.Add("Август");
            listMonthStart.Add("Сентябрь");
            listMonthStart.Add("Октябрь");
            listMonthStart.Add("Ноябрь");
            listMonthStart.Add("Декабрь");



            listMonthEnd.Add("Январь");
            listMonthEnd.Add("Февраль");
            listMonthEnd.Add("Март");
            listMonthEnd.Add("Апрель");
            listMonthEnd.Add("Май");
            listMonthEnd.Add("Июнь");
            listMonthEnd.Add("Июль");
            listMonthEnd.Add("Август");
            listMonthEnd.Add("Сентябрь");
            listMonthEnd.Add("Октябрь");
            listMonthEnd.Add("Ноябрь");
            listMonthEnd.Add("Декабрь");

            // Заполним раскрывающиеся списки названиями месяцев.
            this.cmbStart.DataSource = listMonthStart;
            this.cmbEnd.DataSource = listMonthEnd;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}