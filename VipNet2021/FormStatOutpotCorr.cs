using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;

namespace RegKor
{
    public partial class FormStatOutpotCorr : Form
    {
        private RangeDate rd;

        // Таблица с данными.
        private DataTable rez;

        // Таблица с исполнителями.
        private DataTable rezIsp;

        /// <summary>
        /// Хранит диапазон дат.
        /// </summary>
        public RangeDate ДиапазонДат
        {
            get
            {
                return rd;
            }
            set
            {
                rd = value;
            }
        }

        public FormStatOutpotCorr()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
           
            СтатиистикаИсходящейКорреспонденции getData = new СтатиистикаИсходящейКорреспонденции(rd);

            DataTable tPerson = getData.Исполнители();

            List<СтатистикаВходИсполнителей> list = new List<СтатистикаВходИсполнителей>();

            // Шапка отчета.
            СтатистикаВходИсполнителей itemHead = new СтатистикаВходИсполнителей();
            itemHead.НомерПП = "№ п.п.";
            itemHead.НаименованиеКорреспондента = "Наименование корреспондента";
            itemHead.КоличесвтоВходДокументов = "Количество исходящих документов";
            itemHead.БумажныйНоститель = "Бумажный носитель";
            itemHead.VipNet = "VipNet";
            itemHead.Fax = "факс";
            itemHead.EMail = "e-mail";
            itemHead.Исполнитель = "Исполнитель";

            list.Add(itemHead);

            // Пройдемся по исполнителям.
            foreach (DataRow row in tPerson.Rows)
            {
                // Получим перечень строк для текущего пользователя.
                DataRow[] rows = rez.Select("ОписаниеПолучателя = '" + row["ОписаниеПолучателя"].ToString().Trim() + "' ");

                int iCount = 1;
                foreach (DataRow r in rows)
                {
                    СтатистикаВходИсполнителей itm = new СтатистикаВходИсполнителей();
                    itm.НомерПП = iCount.ToString();
                    itm.НаименованиеКорреспондента = r["ОписаниеКорреспондента"].ToString().Trim();

                    itm.КоличесвтоВходДокументов = r["КоличествоИсходДокументов"].ToString().Trim();
                    itm.БумажныйНоститель = r["бумага"].ToString().Trim();
                    itm.EMail = r["мыло"].ToString().Trim();
                    itm.VipNet = r["вип"].ToString().Trim();
                    itm.Fax = r["факс"].ToString().Trim();
                    itm.Исполнитель = r["ОписаниеПолучателя"].ToString().Trim();

                    list.Add(itm);

                    iCount++;
                }

               

                DataRow rCount = getData.ОтобразитьИтогоИсполнитель(row["ОписаниеПолучателя"].ToString().Trim()).Rows[0];

                // Зпапишем строку итого для исполнителя.
                СтатистикаВходИсполнителей itCount = new СтатистикаВходИсполнителей();
                itCount.НомерПП = "Итого по исполнителю " + rCount["ОписаниеПолучателя"].ToString().Trim();
                itCount.Исполнитель = "----------";
                itCount.КоличесвтоВходДокументов = rCount["КолИсходДокументов"].ToString().Trim();
                itCount.БумажныйНоститель = rCount["бумага"].ToString().Trim();
                itCount.EMail = rCount["мыло"].ToString().Trim();
                itCount.VipNet = rCount["вип"].ToString().Trim();
                itCount.Fax = rCount["факс"].ToString().Trim();
                itCount.Исполнитель = rCount["ОписаниеПолучателя"].ToString().Trim();

                list.Add(itCount);
            }

            int iCount2 = 1;

            int countВсеДокументы = 0;
            int countВсеБумага = 0;
            int countВсеМыло = 0;
            int countВсеВип = 0;
            int countВсеФакс = 0;

            DataTable dtRowsCorrespondent = getData.ОтобразитьИтогоАдресат();
            // Запишем итого по по комитету по корреспондентам.
            foreach (DataRow r in dtRowsCorrespondent.Rows)
            {
                СтатистикаВходИсполнителей itm = new СтатистикаВходИсполнителей();
                itm.НомерПП = iCount2.ToString();
                itm.НаименованиеКорреспондента = r["ОписаниеКорреспондента"].ToString().Trim();

                itm.КоличесвтоВходДокументов = r["КоличествоИсходДокументов"].ToString().Trim();

                if (DBNull.Value != r["КоличествоИсходДокументов"])
                {
                    countВсеДокументы += Convert.ToInt32(r["КоличествоИсходДокументов"]);
                }

                itm.БумажныйНоститель = r["бумага"].ToString().Trim();

                if (DBNull.Value != r["бумага"])
                {
                    countВсеБумага += Convert.ToInt32(r["бумага"]);
                }

                itm.EMail = r["мыло"].ToString().Trim();

                if (DBNull.Value != r["мыло"])
                {
                    countВсеМыло += Convert.ToInt32(r["мыло"]);
                }

                itm.VipNet = r["вип"].ToString().Trim();

                if (DBNull.Value != r["вип"])
                {
                    countВсеВип += Convert.ToInt32(r["вип"]);
                }

                itm.Fax = r["факс"].ToString().Trim();

                if (DBNull.Value != r["факс"])
                {
                    countВсеФакс += Convert.ToInt32(r["факс"]);
                }

                itm.Исполнитель = "-----";

                list.Add(itm);

                iCount2++;
            }

            СтатистикаВходИсполнителей count = new СтатистикаВходИсполнителей();
            count.НомерПП = "Итого в целом по комитету";
            count.НаименованиеКорреспондента = "-----";
            count.КоличесвтоВходДокументов = countВсеДокументы.ToString();
            count.БумажныйНоститель = countВсеБумага.ToString();
            count.EMail = countВсеМыло.ToString();
            count.VipNet = countВсеВип.ToString();
            count.Fax = countВсеФакс.ToString();

            list.Add(count);

            ReportСтатистикаИсходящихДокументов reportStatistic = new ReportСтатистикаИсходящихДокументов("Статистика по исходящей корреспонденции c " + this.ДиапазонДат.DataStart.ToShortDateString() + " по " + this.ДиапазонДат.DataEnd.ToShortDateString() + " по КСЗН г. Саратова");

            PrintReportStaticOutputDoc report = new PrintReportStaticOutputDoc(reportStatistic);
            report.ListDate = list;
            report.Execute();


            //ExcelPrint excel = new ExcelPrint(" Статистика по исходящей корреспонденции c " + this.ДиапазонДат.DataStart.ToShortDateString() + " по " + this.ДиапазонДат.DataEnd.ToShortDateString() + " по КСЗН г. Саратова");

            //// Сохраним коллекцию в формате текстового файла.
            ////excel.PrintСтатистикаВходящейКорреспонденции(list);

            //// Работает быстро пока закоментируем вдруг не надо будет.
            //excel.SaveFileCSV(list);

            //// Передадим всё жто в Excel.
            this.Close();
        }

        private void FormStatOutpotCorr_Load(object sender, EventArgs e)
        {
            СтатиистикаИсходящейКорреспонденции getData = new СтатиистикаИсходящейКорреспонденции(rd);
            
            rez = getData.ОтобразитьDataGridView();

            this.dataGridView1.DataSource = rez;

            this.dataGridView1.Columns["мыло"].HeaderText = "e-mail";

        }
    }
}