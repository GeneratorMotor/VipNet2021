using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;

namespace RegKor
{
    public partial class FormДиапазонДатОтправка : Form
    {
        //Поля для вставки значения дат
        public string BeginDate = null;
        public string EndDate = null;

        public FormДиапазонДатОтправка()
        //public FormДиапазонДатОтправка(RegKor.DS1 ds)
        {
            InitializeComponent();
            //this.dS11 = ds;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
            BeginDate = null;
            EndDate = null;
        }

        private void FormДиапазонДатОтправка_Load(object sender, EventArgs e)
        {
            dt1.Value = DateTime.Now.AddMonths(-1);//начало отчетного периода на 1 месяц меньше сегодняшней даты
            dt2.Value = DateTime.Now;// конец отчетного период равен сегодняшняй дате
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            FormView frmPrint = new FormView();

            
            BeginDate = dt1.Value.ToShortDateString();
            EndDate = dt2.Value.ToShortDateString();

            //string BeginDateSQL = System.Text.RegularExpressions.Regex.Replace(BeginDate, "\\b(?<day>\\d{1,2}).(?<month>\\d{1,2}).(?<year>\\d{2,4})\\b", "${month}-${day}-${year}");
            //string EndDateSQL = System.Text.RegularExpressions.Regex.Replace(EndDate, "\\b(?<day>\\d{1,2}).(?<month>\\d{1,2}).(?<year>\\d{2,4})\\b", "${month}-${day}-${year}");

            string BeginDateSQL = System.Text.RegularExpressions.Regex.Replace(BeginDate, "\\b(?<day>\\d{1,2}).(?<month>\\d{1,2}).(?<year>\\d{2,4})\\b", "${year}${month}${day}");
            string EndDateSQL = System.Text.RegularExpressions.Regex.Replace(EndDate, "\\b(?<day>\\d{1,2}).(?<month>\\d{1,2}).(?<year>\\d{2,4})\\b", "${year}${month}${day}");

            //Подключаемся к БД и заполняем DataSet
            Classess.СтатистикаОтправленныхДокументов статистикаДокументов = new RegKor.Classess.СтатистикаОтправленныхДокументов();
            //this.dS11 = (RegKor.DS1)статистикаДокументов.ВременнойИнтервал(BeginDateSQL, EndDateSQL);
            DataSet ds = статистикаДокументов.ВременнойИнтервал(BeginDateSQL, EndDateSQL);

            //Заполним dS11 данными

            foreach(DataRow rowДокумент in ds.Tables[0].Rows)
            {
                DataRow row1 = dS11.ВыборкаКоличествоИсходящихДокументов.NewRow();
                row1[0] = rowДокумент[0];
                row1[1] = rowДокумент[1];
                dS11.ВыборкаКоличествоИсходящихДокументов.Rows.Add(row1);
            }

            try
            {
                ReportDocument rptDoc = new ReportDocument();
                //string fileName = @"..\report\Statistic.rpt";
                string fileName = @"..\report\StatisticOutPost.rpt";
                // загружает файл отчета:
                rptDoc.Load(fileName);
                //// источник данных:
                rptDoc.SetDataSource(this.dS11);
                // просмотрщику передали источник отчета:
                frmPrint.reportViewer.ReportSource = rptDoc;
                FormГлавная.ПараметрыДляОтчета("BeginDate", BeginDate, frmPrint.reportViewer.ParameterFieldInfo);
                FormГлавная.ПараметрыДляОтчета("EndDate", EndDate, frmPrint.reportViewer.ParameterFieldInfo);
                frmPrint.Text = "Статистика поступления";
                //// показываем форму:
                frmPrint.reportViewer.ShowGroupTreeButton = false;
                this.Hide();
                frmPrint.ShowDialog(this);
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, "Произошла ошибка при открытии файла отчета \"Статистика поступления\".\n" + exc.Message, "Ошибка открытия файла отчета");
                return;
            }
            finally
            {
                this.Enabled = true;
                BeginDate = null;
                EndDate = null;
                //Dispose(true);
                this.Close();
            }
        }

        
    }
}