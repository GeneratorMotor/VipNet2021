using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;
using CrystalDecisions.CrystalReports.Engine;

namespace RegKor
{
    public partial class FormПечатьКарточки : Form
    {
        private string _текущийГод;
        private string _будущийГод;

        /// <summary>
        /// Хранит первое января будущего года
        /// </summary>
        public string БудущийГод
        {
            get
            {
                return _будущийГод;
            }
            set
            {
                _будущийГод = value;
            }
        }

        /// <summary>
        /// Хранит первое января текущего года
        /// </summary>
        public string ТекущийГод
        {
            get
            {
                return _текущийГод;
            }
            set
            {
                _текущийГод = value;
            }
        }

        public FormПечатьКарточки()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            КарточкаРаспечатанная карточка = new КарточкаРаспечатанная();
            карточка.ПервыйНомерКарточки = this.textBox1.Text;
            карточка.КрайнийНомерКарточки = this.textBox2.Text;

            if (карточка.ПервыйНомерКарточки.Length == 0 && карточка.КрайнийНомерКарточки.Length == 0)
            {
                System.Windows.Forms.MessageBox.Show("Введите номер карточек");
                return;
            }

            карточка.ТекущийГод = this.ТекущийГод;
            карточка.БудущийГод = this.БудущийГод;

            //DataSet dsКарточка = карточка.ПолучитьДанные();
            List<Карточка> dsКарточка = карточка.ПолучитьДанные();

            //// Попробуем скопировать в файл несколько шаблонов.
            //string fName = startDate + "_Карточка_" + endDate;

            ////try
            ////{
            ////Скопируем шаблон в папку Документы
            //FileInfo fn = new FileInfo(System.Windows.Forms.Application.StartupPath + @"\Шаблоны\Карточка.doc");
            //fn.CopyTo(System.Windows.Forms.Application.StartupPath + @"\Отчёты\" + fName + ".doc", true);

            //string filName = System.Windows.Forms.Application.StartupPath + @"\Отчёты\" + fName + ".doc";

            ////Создаём новый Word.Application
            //Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            ////Загружаем документ
            //Microsoft.Office.Interop.Word.Document doc = null;

            //object fileName = filName;
            //object falseValue = false;
            //object trueValue = true;
            //object missing = Type.Missing;
            //object writePasswordDocument = "12A86Asd";

            ////1

            ////старая рабочая реализация 
            ////doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ////ref missing, ref missing, ref missing, ref missing, ref missing,
            ////ref missing, ref missing, ref missing, ref missing, ref missing,
            ////ref missing, ref missing, ref missing);

            //doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            //ref missing, ref missing, ref missing, ref missing, ref writePasswordDocument,
            //ref missing, ref missing, ref missing, ref missing, ref trueValue,
            //ref missing, ref missing, ref missing);

            //////Дата начало отчёта.
            //object wdrepl = WdReplace.wdReplaceAll;
            ////object searchtxt = "GreetingLine";
            //object searchtxt = "DATESTART";
            //object newtxt = (object)DateTime.Today.ToShortDateString();
            ////object frwd = true;
            //object frwd = false;
            //doc.Content.Find.Execute(ref searchtxt, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd, ref missing, ref missing, ref newtxt, ref wdrepl, ref missing, ref missing,
            //ref missing, ref missing);

            //doc.AddDocumentWorkspaceHeader(

            FormView frmPrint = new FormView();

            try
            {

                ReportDocument rptDoc = new ReportDocument();
                string fileName = @"..\report\MCard.rpt";

                // загружает файл отчета:
                rptDoc.Load(fileName);

                // источник данных: попробовать создать динамически DataTable заполнить их и добавить в динамически созданный DataSet
                rptDoc.SetDataSource(dsКарточка);
                //////rptDoc.SetDataSource(списокКоличествоЖалобОтчёт);

                // просмотрщику передали источник отчета:
                frmPrint.reportViewer.ReportSource = rptDoc;
                //FormГлавная.ПараметрыДляОтчета("BeginDate", BeginDate, frmPrint.reportViewer.ParameterFieldInfo);
                //FormГлавная.ПараметрыДляОтчета("EndDate", EndDate, frmPrint.reportViewer.ParameterFieldInfo);
                frmPrint.Text = "Картчки";
                // показываем форму:
                frmPrint.reportViewer.ShowGroupTreeButton = false;
                //this.Hide();
                frmPrint.ShowDialog(this);
                //frmPrint.Show();
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, "Произошла ошибка при открытии файла отчета \"Статистика поступления\".\n" + exc.Message, "Ошибка открытия файла отчета");
                return;
            }
            finally
            {
                this.Enabled = true;
                //BeginDate = null;
                //EndDate = null;
                Dispose(true);
            }
        }
    }
}