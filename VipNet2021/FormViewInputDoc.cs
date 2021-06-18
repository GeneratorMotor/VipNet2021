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
        /// ������ �������� ���.
        /// </summary>
        public RangeDate �����������
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
            this.lblPeriod.Text = "����� �� ������ � " + rd.DataStart.ToShortDateString() + " �� " + rd.DataEnd.ToShortDateString();

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
            this.dataGridView1.Columns["id_��������������"].Visible = false;
            this.dataGridView1.Columns["id_����������"].Visible = false;
        }

       

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            if (this.radioButton4.Checked == true)
            {
                int idCorr = (int)this.comboBox1.SelectedValue;
                int idPerson = (int)this.comboBox2.SelectedValue;

                ����������������������� view = new �����������������������(rd);
                this.dataGridView1.DataSource = view.Get���������Id�������������Id�����������(idCorr, idPerson);

                DisplayDataGrid();

            }
        }

        /// <summary>
        /// �������� ������ ��� �������� �����.
        /// </summary>
        private void LoadForm()
        {
            ����������������������� view = new �����������������������(rd);
            this.dataGridView1.DataSource = view.GetTablePeriodDate();

            DisplayDataGrid();

            // ��������� �������������� ������.
            this.comboBox1.DataSource = view.Get��������������();
            this.comboBox1.DisplayMember = "����������������������";
            this.comboBox1.ValueMember = "id_��������������";

            // ������� ������������.
            this.comboBox2.DataSource = view.Get�����������();
            this.comboBox2.DisplayMember = "������������������";
            this.comboBox2.ValueMember = "id_����������";
        }

        private void LoadData������������������������()
        {
            int id = (int)this.comboBox1.SelectedValue;

            ����������������������� view = new �����������������������(rd);
            this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

            DisplayDataGrid();

            // ������� ������������.
            this.comboBox2.DataSource = view.Get�����������IdCorr(id);
            this.comboBox2.DisplayMember = "������������������";
            this.comboBox2.ValueMember = "id_����������";
        }

        /// <summary>
        /// ������� ���� ������������ �� ��������� ���������������.
        /// </summary>
        private void LoadData�����������������������������������()
        {
            int id = (int)this.comboBox1.SelectedValue;

            ����������������������� view = new �����������������������(rd);
            //this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

            //DisplayDataGrid();

            // ������� ������������.
            this.comboBox2.DataSource = null;
            this.comboBox2.DataSource = view.Get�����������IdCorr(id);
            this.comboBox2.DisplayMember = "������������������";
            this.comboBox2.ValueMember = "id_����������";

            this.dataGridView1.DataSource = null;
        }

        /// <summary>
        /// ����� �������������� �� �����������.
        /// </summary>
        private void LoadData���������������������������������()
        {
            int id = (int)this.comboBox2.SelectedValue;

            ����������������������� view = new �����������������������(rd);
            //this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

            //DisplayDataGrid();

            // ������� ������������.
            this.comboBox1.DataSource = null;
            this.comboBox1.DataSource = view.Get�������������Id�����������(id);
            this.comboBox1.DisplayMember = "����������������������";
            this.comboBox1.ValueMember = "id_��������������";

            this.dataGridView1.DataSource = null;
        }

        private void LoadData���������������������()
        {
            int id = (int)this.comboBox2.SelectedValue;

            ����������������������� view = new �����������������������(rd);

            // ��������� ������ ���������������.
            this.comboBox1.DataSource = view.Get�������������Id�����������(id);
            this.comboBox1.DisplayMember = "����������������������";
            this.comboBox1.ValueMember = "id_��������������";

            // ��������� ������ ����������.
            this.dataGridView1.DataSource = view.Get���������Id�����������(id);
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
                LoadData������������������������();
            }

            if (this.radioButton3.Checked == true)
            {
                LoadData���������������������();
            }

            if (this.radioButton4.Checked == true)
            {
                if (this.chkFiltrPerson.Checked == true)
                {
                    // ������� ���� ������������ �� ������� ������.
                    LoadData�����������������������������������();
                }
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            // ������� ���.
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
            
            // ������� �� ���������������.
            if (this.radioButton2.Checked == true)
            {
                //int id = (int)this.comboBox1.SelectedValue;

                ����������������������� view = new �����������������������(rd);
                this.comboBox1.DataSource = view.Get��������������();
                this.comboBox1.DisplayMember = "����������������������";
                this.comboBox1.ValueMember = "id_��������������";

                LoadData������������������������();
            }
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            // ������� �� ������������.
            if (this.radioButton3.Checked == true)
            {
                // ������� ������������.
                ����������������������� view = new �����������������������(rd);
                this.comboBox2.DataSource = view.Get�����������();
                this.comboBox2.DisplayMember = "������������������";
                this.comboBox2.ValueMember = "id_����������";

                LoadData���������������������();


            }

        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (this.radioButton3.Checked == true)
            {
                LoadData���������������������();
            }
            if (this.radioButton4.Checked == true)
            {
                if (this.chkCorr.Checked == true)
                {
                    LoadData���������������������������������();
                }
            }

        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            // ������� �� ������������.
            if (this.radioButton4.Checked == true)
            {
                // ������� ������������.
                ����������������������� view = new �����������������������(rd);

                // ������� ���� ������������.
                this.comboBox2.DataSource = view.Get�����������();
                this.comboBox2.DisplayMember = "������������������";
                this.comboBox2.ValueMember = "id_����������";

                // ������� ���� ���������������.
                this.comboBox1.DataSource = view.Get��������������();
                this.comboBox1.DisplayMember = "����������������������";
                this.comboBox1.ValueMember = "id_��������������";

                // ������� ������ ���������� �� ������� ������ ������� 
                int idCorr = (int)this.comboBox1.SelectedValue;

                int idPers = (int)this.comboBox2.SelectedValue;

                // �������������� �������� ���������� ��� ��������� � ������� �������.
                this.dataGridView1.DataSource = view.Get���������Id�������������Id�����������(idCorr, idPers);
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
                // ������ ����� �� ��������������.
                this.chkCorr.Checked = false;

                // ��������� ���� �� ������� ������ ���� ������������.
                // ������� ������������.
                ����������������������� view = new �����������������������(rd);

                // ������� ���� ������������.
                this.comboBox2.DataSource = view.Get�����������();
                this.comboBox2.DisplayMember = "������������������";
                this.comboBox2.ValueMember = "id_����������";

                // ������� ���� ���������������.
                this.comboBox1.DataSource = view.Get��������������();
                this.comboBox1.DisplayMember = "����������������������";
                this.comboBox1.ValueMember = "id_��������������";
            }
        }

        private void chkCorr_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkCorr.Checked == true)
            {
                // ������ ���� �� ������������.
                this.chkFiltrPerson.Checked = false;

                // ������� ������������.
                ����������������������� view = new �����������������������(rd);

                // ������� ���� ������������.
                this.comboBox2.DataSource = view.Get�����������();
                this.comboBox2.DisplayMember = "������������������";
                this.comboBox2.ValueMember = "id_����������";

                // ������� ���� ���������������.
                this.comboBox1.DataSource = view.Get��������������();
                this.comboBox1.DisplayMember = "����������������������";
                this.comboBox1.ValueMember = "id_��������������";
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            string caption = "����� � �������� ���������� �� ������ � " + rd.DataStart.ToShortDateString() + " �� " + rd.DataEnd.ToShortDateString();
            Report����������������������� printDate = new Report�����������������������(caption);
            printDate.SetDate = this.dataGridView1.Rows;

            PrintReport printPaper = new PrintReport();
            printPaper.SetCommand(printDate);
            printPaper.Execute();

            //�������������������� report = new ��������������������();
            //report.DataGridView1 = this.dataGridView1;

            //int rCount = report.DataGridView1.RowCount;

            //WordPrint word = new WordPrint("����� � �������� ���������� �� ������ � " + rd.DataStart.ToShortDateString() + " �� " + rd.DataEnd.ToShortDateString());
            //word.Print(report);


            //Excel����� excel = new Excel�����();
            //string caption = "����� � �������� ���������� �� ������ � " + rd.DataStart.ToShortDateString() + " �� " + rd.DataEnd.ToShortDateString();
            //excel.Print�����OfDataTable(report, caption.Trim(), "A1", "J1");

            this.Close();
        }

     
    }
}