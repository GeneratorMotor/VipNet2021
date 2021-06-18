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
    public partial class FormViewOutputDoc : Form
    {

        private RangeDate rd;

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

        public FormViewOutputDoc()
        {
            InitializeComponent();
        }

       
        private void LoadData()
        {
            this.lblPeriod.Text = "Отчет за период с " + rd.DataStart.ToShortDateString() + " по " + rd.DataEnd.ToShortDateString();

            ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);
            this.dataGridView1.DataSource = view.ВсеДокументыЗаПериод();

            DisplayDataGrid();

            // Отобразим раскрывающейся список.
            this.comboBox1.DataSource = view.GetАдресаты();
            this.comboBox1.DisplayMember = "Адресат";
            this.comboBox1.ValueMember = "id_Адресата";

            //// Выведим исполнителей.
            this.comboBox2.DataSource = view.GetИсполнители();
            this.comboBox2.DisplayMember = "ОписаниеПолучателя";
            this.comboBox2.ValueMember = "id_получателя";
        }

        private void FormViewOutputDoc_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            // Выборка все.
            if (this.radioButton1.Checked == true)
            {
                LoadData();

                this.comboBox1.Enabled = false;
                this.comboBox2.Enabled = false;
            }
            else
            {
                this.comboBox1.Enabled = true;
                this.comboBox2.Enabled = true;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            // Выборка по корреспондентам.
            if (this.radioButton2.Checked == true)
            {
                ViewDocCorrespondent();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            // Выборка по корреспондентам.
            if (this.radioButton2.Checked == true)
            {
                ViewDocCorrespondent();
            }

            // Выборка по исполнителям.
            if (this.radioButton3.Checked == true)
            {
                LoadDataВыборкаПоИсполнителям();
            }

            // Выборка и по исполнителям и по корреспондентам.
            if (this.radioButton4.Checked == true)
            {
                if (this.chkFiltrPerson.Checked == true)
                {
                    // id льготника.
                    int idPerson = (int)this.comboBox1.SelectedValue;

                    // Выборка всех исполнителей за текущий период.
                    LoadDataВыборкаИсполнителейПоКорреспонденту(idPerson);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void ViewDocCorrespondent()
        {
            int id = (int)this.comboBox1.SelectedValue;

            ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);
            this.comboBox1.DataSource = view.GetАдресаты();
            this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
            this.comboBox1.ValueMember = "id_Адресата";

            this.dataGridView1.DataSource = view.ДокументыВыборкаПоКорреспондентам(id);

            DisplayDataGrid();
        }

        /// <summary>
        /// Выдорка по исполнителям.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            int id = (int)this.comboBox2.SelectedValue;

            ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);
            this.dataGridView1.DataSource = view.ДокументыВыборкаПоИсполнителям(id);

            DisplayDataGrid();

            // Отобразим раскрывающейся список.
            this.comboBox1.DataSource = view.GetАдресаты();
            this.comboBox1.DisplayMember = "Адресат";
            this.comboBox1.ValueMember = "id_Адресата";

            //// Выведим исполнителей.
            this.comboBox2.DataSource = view.GetИсполнители();
            this.comboBox2.DisplayMember = "ОписаниеПолучателя";
            this.comboBox2.ValueMember = "id_получателя";
        }

        /// <summary>
        /// Список по исполнителям.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (this.radioButton3.Checked == true)
            {
                int id = (int)this.comboBox2.SelectedValue;

                ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);

                // Отобразим список корреспондентов.
                this.comboBox1.DataSource = view.GetАдресатыПоИсполнителям(id);
                this.comboBox1.DisplayMember = "Адресат";
                this.comboBox1.ValueMember = "id_Адресата";

                // Отобразим список документов.
                this.dataGridView1.DataSource = view.ДокументыВыборкаПоИсполнителям(id);

                DisplayDataGrid();

            }

            if (this.radioButton4.Checked == true)
            {
                if (this.chkCorr.Checked == true)
                {
                    // Выбераем документы и корреспонденты в зависимости от выбранного исполнителя.
                    int id = (int)this.comboBox2.SelectedValue;

                    ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);

                    this.comboBox1.DataSource = view.GetАдресатыПоИсполнителям(id);
                    this.comboBox1.DisplayMember = "Адресат";
                    this.comboBox1.ValueMember = "id_Адресата";

                    this.dataGridView1.DataSource = null;
                }
            }

            //DisplayDataGrid();
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            // Выборка по корреспондентам и исполнителям.
            if (this.radioButton4.Checked == true)
            {
                // Выведим исполнителей.
                ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);

                // Выведим всех исполнителей.
                this.comboBox2.DataSource = view.GetИсполнители();
                this.comboBox2.DisplayMember = "ОписаниеПолучателя";
                this.comboBox2.ValueMember = "id_получателя";

                // Выведем всех корреспондентов.
                this.comboBox1.DataSource = view.GetАдресаты();
                this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
                this.comboBox1.ValueMember = "id_Адресата";

                // Выведим список документов за текущий период которые 
                int idCorr = (int)this.comboBox1.SelectedValue;

                int idPers = (int)this.comboBox2.SelectedValue;

                //this.dataGridView1.DataSource = view.ДокументыВыборкаПоИсполнителямАдресатам(idPers, idCorr);
                //DisplayDataGrid();

                this.btnDisplay.Enabled = true;

                this.chkFiltrPerson.Checked = true;
                this.chkFiltrPerson.Enabled = true;
                this.chkCorr.Enabled = true;
                this.chkCorr.Checked = false;

                this.dataGridView1.DataSource = view.ДокументыВыборкаПоИсполнителямАдресатам(idPers, idCorr);
                DisplayDataGrid();

            }
            else
            {
                this.btnDisplay.Enabled = false;
                this.chkFiltrPerson.Checked = false;
                this.chkFiltrPerson.Enabled = false;
                this.chkCorr.Enabled = false;
                this.chkCorr.Checked = false;

                DisplayDataGrid();

            }
        }

        /// <summary>
        /// Скрывает от пользователя лишную информацию.
        /// </summary>
        private void DisplayDataGrid()
        {
            this.dataGridView1.Columns["id_Адресата"].Visible = false;
            this.dataGridView1.Columns["id_получателя"].Visible = false;
        }

        //"SELECT [id_Адресата] " +
        //                  " ,id_получателя " +
        //                 " ,[Адресат] " +
        //                 " ,[ДатаИсходящая] " +
        //                 " ,[НомерКомитета] " +
        //                 " ,[НомерНоменклатурный] " +
        //                 " ,[НомерПодразделения] " +
        //                 " ,[НомерПорядковый] " +
        //                 " ,[Содержание] " +
        //                 " ,[ОписаниеПолучателя] " +

        private void chkFiltrPerson_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkFiltrPerson.Checked == true)
            {
                // Скрыть выбор по корреспонденту.
                this.chkCorr.Checked = false;

                // Отобразим всех за текущий период всех исполнителей.
                // Выведим исполнителей.
                ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);

                // Выведим всех исполнителей.
                this.comboBox2.DataSource = view.GetИсполнители();
                this.comboBox2.DisplayMember = "ОписаниеПолучателя";
                this.comboBox2.ValueMember = "id_получателя";

                // Выведем всех корреспондентов.
                this.comboBox1.DataSource = view.GetАдресаты();
                this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
                this.comboBox1.ValueMember = "id_Адресата";
            }
        }

        private void chkCorr_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkCorr.Checked == true)
            {
                // Скрыть выбо по исполнителям.
                this.chkFiltrPerson.Checked = false;

                // Выведим исполнителей.
                ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);

                // Выведим всех исполнителей.
                this.comboBox2.DataSource = view.GetИсполнители();
                this.comboBox2.DisplayMember = "ОписаниеПолучателя";
                this.comboBox2.ValueMember = "id_получателя";

                // Выведем всех корреспондентов.
                this.comboBox1.DataSource = view.GetАдресаты();
                this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
                this.comboBox1.ValueMember = "id_Адресата";
            }
        }

        private void LoadDataВыборкаПоИсполнителям()
        {
            int id = (int)this.comboBox2.SelectedValue;

            ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);

            // Отобразим список корреспондентов.
            this.comboBox1.DataSource = view.GetАдресатыПоИсполнителям(id);
            this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
            this.comboBox1.ValueMember = "id_Адресата";

            // Отобразим список документов.
            this.dataGridView1.DataSource = view.ДокументыВыборкаПоИсполнителям(id);
            DisplayDataGrid();
        }


        /// <summary>
        /// Выборка всех исполнителей по выбранным корреспондентам.
        /// </summary>
        private void LoadDataВыборкаИсполнителейПоКорреспонденту(int idPerson)
        {
            if (this.radioButton4.Checked == true)
            {

                //int id = (int)this.comboBox1.SelectedValue;

                ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);
                //this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

                //DisplayDataGrid();

                // Выведим исполнителей по корреспонденту.
                this.comboBox2.DataSource = null;
                this.comboBox2.DataSource = view.GetИсполнителиПоКорреспонденту(idPerson);
                this.comboBox2.DisplayMember = "ОписаниеПолучателя";
                this.comboBox2.ValueMember = "id_получателя";

                this.dataGridView1.DataSource = null;
                //this.dataGridView1.DataSource = view.ДокументыВыборкаПоИсполнителямАдресатам(idPers, idCorr);
            }

        }

        /// <summary>
        /// Выборка всех исполнителей по выбранным корреспондентам.
        /// </summary>
        private void LoadDataВыборкаКорреспондентовПоИсполнителю(int idPerson)
        {
            if (this.radioButton4.Checked == true)
            {

                //int id = (int)this.comboBox1.SelectedValue;

                ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);
                //this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

                //DisplayDataGrid();

                // Выведим исполнителей по корреспонденту.
                this.comboBox1.DataSource = null;
                this.comboBox1.DataSource = view.GetИсполнителиПоКорреспонденту(idPerson);
                this.comboBox1.DisplayMember = "ОписаниеПолучателя";
                this.comboBox1.ValueMember = "id_Адресата";

                this.dataGridView1.DataSource = null;
                //this.dataGridView1.DataSource = view.ДокументыВыборкаПоИсполнителямАдресатам(idPers, idCorr);
            }

        }

        private void ВыборкаДокументовИКорреспондентов_IdИсполнители()
        {
            // Получим корреспондентов.
            int idPers = (int)this.comboBox2.SelectedValue;
            int idCorr = (int)this.comboBox1.SelectedValue;

            ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);
            //DisplayDataGrid();

            // Выведим корреспондентов по исполнителю.
            this.comboBox1.DataSource = null;
            this.comboBox1.DataSource = view.GetАдресатыПоИсполнителям(idPers);
            this.comboBox1.DisplayMember = "ОписаниеПолучателя";
            this.comboBox1.ValueMember = "id_получателя";

            this.dataGridView1.DataSource = null;

            // Выберим документы в зависимости от исполнителя и корреспондента.
            this.dataGridView1.DataSource = view.ДокументыВыборкаПоИсполнителямАдресатам(idPers, idCorr);

            DisplayDataGrid();

        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            if (this.radioButton4.Checked == true)
            {
                //Отобразить документы в зависимости от выбранного корреспондента и исполнителя.
                // Получим корреспондентов.
                int idPers = (int)this.comboBox2.SelectedValue;
                int idCorr = (int)this.comboBox1.SelectedValue;

                ОтчетИсходящихДокументов view = new ОтчетИсходящихДокументов(rd);
                //DisplayDataGrid();

                //// Выведим корреспондентов по исполнителю.
                //this.comboBox1.DataSource = null;
                //this.comboBox1.DataSource = view.GetАдресатыПоИсполнителям(idPers);
                //this.comboBox1.DisplayMember = "ОписаниеПолучателя";
                //this.comboBox1.ValueMember = "id_получателя";

                this.dataGridView1.DataSource = null;

                // Выберим документы в зависимости от исполнителя и корреспондента.
                this.dataGridView1.DataSource = view.ДокументыВыборкаПоИсполнителямАдресатам(idPers, idCorr);

                DisplayDataGrid();
            }

        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // Распечатаем содержимое DatagridView.
            //ОтчетОВходДокументах report = new ОтчетОВходДокументах();
            //report.DataGridView1 = this.dataGridView1;

            string caption = "Отчет об исходящих документах за период с " + rd.DataStart.ToShortDateString() + " по " + rd.DataEnd.ToShortDateString();

            ReportОтчетИсходящиеДокументы reportDoc = new ReportОтчетИсходящиеДокументы(caption);

            ReportНомераИсходящихДокументов report = new ReportНомераИсходящихДокументов(reportDoc);
            report.ListDate = this.dataGridView1.Rows;
            report.Execute();

            //reportDoc.PrintReportStaticOutputDoc(this.dataGridView1.Rows);

            //ExcelОтчет excel = new ExcelОтчет();
            
            //excel.PrintОтчетИсходящиеДокументы(report, caption, "A1", "G1");


        }

        //private void DisplayDataGrid()
        //{
        //    this.dataGridView1.Columns["id_получателя"].Visible = false;
        //    //this.dataGridView1.Columns["id_получателя"].Visible = false;
        //}


         //"SELECT [id_Адресата] " +
         //                  " ,id_получателя " +
         //                 " ,[Адресат] " +
         //                 " ,[ДатаИсходящая] " +
         //                 " ,[НомерКомитета] " +
         //                 " ,[НомерНоменклатурный] " +
         //                 " ,[НомерПодразделения] " +
         //                 " ,[НомерПорядковый] " +
         //                 " ,[Содержание] " +
         //                 " ,[ОписаниеПолучателя] " +
         

          
    }
}