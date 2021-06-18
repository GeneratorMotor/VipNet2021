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
    public partial class FormViewInputDoc : Form
    {
        private RangeDate rd;

        //private DataTable rez;

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

        public FormViewInputDoc()
        {
            InitializeComponent();
        }

        private void FormViewInputDoc_Load(object sender, EventArgs e)
        {
            this.lblPeriod.Text = "Отчет за период с " + rd.DataStart.ToShortDateString() + " по " + rd.DataEnd.ToShortDateString();

            LoadForm();

            this.comboBox1.Enabled = false;
            this.comboBox2.Enabled = false;

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void DisplayDataGrid()
        {
            this.dataGridView1.Columns["id_корреспондента"].Visible = false;
            this.dataGridView1.Columns["id_получателя"].Visible = false;
        }

       

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            if (this.radioButton4.Checked == true)
            {
                int idCorr = (int)this.comboBox1.SelectedValue;
                int idPerson = (int)this.comboBox2.SelectedValue;

                ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);
                this.dataGridView1.DataSource = view.GetДокументыIdКорреспондентIdИсполнитель(idCorr, idPerson);

                DisplayDataGrid();

            }
        }

        /// <summary>
        /// Загрузка данных при загрузке формы.
        /// </summary>
        private void LoadForm()
        {
            ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);
            this.dataGridView1.DataSource = view.GetTablePeriodDate();

            DisplayDataGrid();

            // Отобразим раскрывающейся список.
            this.comboBox1.DataSource = view.GetКорреспонденты();
            this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
            this.comboBox1.ValueMember = "id_корреспондента";

            // Выведим исполнителей.
            this.comboBox2.DataSource = view.GetИсполнители();
            this.comboBox2.DisplayMember = "ОписаниеПолучателя";
            this.comboBox2.ValueMember = "id_получателя";
        }

        private void LoadDataВыборкаПоКорреспондентам()
        {
            int id = (int)this.comboBox1.SelectedValue;

            ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);
            this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

            DisplayDataGrid();

            // Выведим исполнителей.
            this.comboBox2.DataSource = view.GetИсполнителиIdCorr(id);
            this.comboBox2.DisplayMember = "ОписаниеПолучателя";
            this.comboBox2.ValueMember = "id_получателя";
        }

        /// <summary>
        /// Выборка всех исполнителей по выбранным корреспондентам.
        /// </summary>
        private void LoadDataВыборкаИсполнителейПоКорреспонденту()
        {
            int id = (int)this.comboBox1.SelectedValue;

            ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);
            //this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

            //DisplayDataGrid();

            // Выведим исполнителей.
            this.comboBox2.DataSource = null;
            this.comboBox2.DataSource = view.GetИсполнителиIdCorr(id);
            this.comboBox2.DisplayMember = "ОписаниеПолучателя";
            this.comboBox2.ValueMember = "id_получателя";

            this.dataGridView1.DataSource = null;
        }

        /// <summary>
        /// Выбор корреспондента по исполнителю.
        /// </summary>
        private void LoadDataВыборкаКорреспондентПоИсполнителю()
        {
            int id = (int)this.comboBox2.SelectedValue;

            ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);
            //this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

            //DisplayDataGrid();

            // Выведим исполнителей.
            this.comboBox1.DataSource = null;
            this.comboBox1.DataSource = view.GetКорреспондентIdИсполнитель(id);
            this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
            this.comboBox1.ValueMember = "id_корреспондента";

            this.dataGridView1.DataSource = null;
        }

        private void LoadDataВыборкаПоИсполнителям()
        {
            int id = (int)this.comboBox2.SelectedValue;

            ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);

            // Отобразим список корреспондентов.
            this.comboBox1.DataSource = view.GetКорреспондентIdИсполнитель(id);
            this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
            this.comboBox1.ValueMember = "id_корреспондента";

            // Отобразим список документов.
            this.dataGridView1.DataSource = view.GetДокументыIdИсполнители(id);
            DisplayDataGrid();
            
        }
                

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            if (this.radioButton1.Checked == true)
            {
                LoadForm();
            }
            if (this.radioButton2.Checked == true)
            {
                LoadDataВыборкаПоКорреспондентам();
            }

            if (this.radioButton3.Checked == true)
            {
                LoadDataВыборкаПоИсполнителям();
            }

            if (this.radioButton4.Checked == true)
            {
                if (this.chkFiltrPerson.Checked == true)
                {
                    // Выборка всех исполнителей за текущий период.
                    LoadDataВыборкаИсполнителейПоКорреспонденту();
                }
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            // Выборка все.
            if (this.radioButton1.Checked == true)
            {
                LoadForm();

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
                //int id = (int)this.comboBox1.SelectedValue;

                ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);
                this.comboBox1.DataSource = view.GetКорреспонденты();
                this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
                this.comboBox1.ValueMember = "id_корреспондента";

                LoadDataВыборкаПоКорреспондентам();
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            // Выборка по исполнителям.
            if (this.radioButton3.Checked == true)
            {
                // Выведим исполнителей.
                ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);
                this.comboBox2.DataSource = view.GetИсполнители();
                this.comboBox2.DisplayMember = "ОписаниеПолучателя";
                this.comboBox2.ValueMember = "id_получателя";

                LoadDataВыборкаПоИсполнителям();


            }

        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (this.radioButton3.Checked == true)
            {
                LoadDataВыборкаПоИсполнителям();
            }
            if (this.radioButton4.Checked == true)
            {
                if (this.chkCorr.Checked == true)
                {
                    LoadDataВыборкаКорреспондентПоИсполнителю();
                }
            }

        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            // Выборка по исполнителям.
            if (this.radioButton4.Checked == true)
            {
                // Выведим исполнителей.
                ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);

                // Выведим всех исполнителей.
                this.comboBox2.DataSource = view.GetИсполнители();
                this.comboBox2.DisplayMember = "ОписаниеПолучателя";
                this.comboBox2.ValueMember = "id_получателя";

                // Выведем всех корреспондентов.
                this.comboBox1.DataSource = view.GetКорреспонденты();
                this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
                this.comboBox1.ValueMember = "id_корреспондента";

                // Выведим список документов за текущий период которые 
                int idCorr = (int)this.comboBox1.SelectedValue;

                int idPers = (int)this.comboBox2.SelectedValue;

                // первоначальная загрузка документов при обращении к данному разделу.
                this.dataGridView1.DataSource = view.GetДокументыIdКорреспондентIdИсполнитель(idCorr, idPers);
                DisplayDataGrid();

                this.btnDisplay.Enabled = true;

                this.chkFiltrPerson.Checked = true;
                this.chkFiltrPerson.Enabled = true;
                this.chkCorr.Enabled = true;
                this.chkCorr.Checked = false;
            }
            else
            {
                this.btnDisplay.Enabled = false;
                this.chkFiltrPerson.Checked = false;
                this.chkFiltrPerson.Enabled = false;
                this.chkCorr.Enabled = false;
                this.chkCorr.Checked = false;
            }
        }

        private void chkFiltrPerson_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkFiltrPerson.Checked == true)
            {
                // Скрыть выбор по корреспонденту.
                this.chkCorr.Checked = false;

                // Отобразим всех за текущий период всех исполнителей.
                // Выведим исполнителей.
                ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);

                // Выведим всех исполнителей.
                this.comboBox2.DataSource = view.GetИсполнители();
                this.comboBox2.DisplayMember = "ОписаниеПолучателя";
                this.comboBox2.ValueMember = "id_получателя";

                // Выведем всех корреспондентов.
                this.comboBox1.DataSource = view.GetКорреспонденты();
                this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
                this.comboBox1.ValueMember = "id_корреспондента";
            }
        }

        private void chkCorr_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkCorr.Checked == true)
            {
                // Скрыть выбо по исполнителям.
                this.chkFiltrPerson.Checked = false;

                // Выведим исполнителей.
                ОтчетВходящихДокументов view = new ОтчетВходящихДокументов(rd);

                // Выведим всех исполнителей.
                this.comboBox2.DataSource = view.GetИсполнители();
                this.comboBox2.DisplayMember = "ОписаниеПолучателя";
                this.comboBox2.ValueMember = "id_получателя";

                // Выведем всех корреспондентов.
                this.comboBox1.DataSource = view.GetКорреспонденты();
                this.comboBox1.DisplayMember = "ОписаниеКорреспондента";
                this.comboBox1.ValueMember = "id_корреспондента";
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            string caption = "Отчет о входящих документах за период с " + rd.DataStart.ToShortDateString() + " по " + rd.DataEnd.ToShortDateString();
            ReportОтчетВходящихДокументов printDate = new ReportОтчетВходящихДокументов(caption);
            printDate.SetDate = this.dataGridView1.Rows;

            PrintReport printPaper = new PrintReport();
            printPaper.SetCommand(printDate);
            printPaper.Execute();

            //ОтчетОВходДокументах report = new ОтчетОВходДокументах();
            //report.DataGridView1 = this.dataGridView1;

            //int rCount = report.DataGridView1.RowCount;

            //WordPrint word = new WordPrint("Отчет о входящих документах за период с " + rd.DataStart.ToShortDateString() + " по " + rd.DataEnd.ToShortDateString());
            //word.Print(report);


            //ExcelОтчет excel = new ExcelОтчет();
            //string caption = "Отчет о входящих документах за период с " + rd.DataStart.ToShortDateString() + " по " + rd.DataEnd.ToShortDateString();
            //excel.PrintОтчетOfDataTable(report, caption.Trim(), "A1", "J1");

            this.Close();
        }

     
    }
}