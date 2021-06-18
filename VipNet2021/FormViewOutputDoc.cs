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

        public FormViewOutputDoc()
        {
            InitializeComponent();
        }

       
        private void LoadData()
        {
            this.lblPeriod.Text = "����� �� ������ � " + rd.DataStart.ToShortDateString() + " �� " + rd.DataEnd.ToShortDateString();

            ������������������������ view = new ������������������������(rd);
            this.dataGridView1.DataSource = view.��������������������();

            DisplayDataGrid();

            // ��������� �������������� ������.
            this.comboBox1.DataSource = view.Get��������();
            this.comboBox1.DisplayMember = "�������";
            this.comboBox1.ValueMember = "id_��������";

            //// ������� ������������.
            this.comboBox2.DataSource = view.Get�����������();
            this.comboBox2.DisplayMember = "������������������";
            this.comboBox2.ValueMember = "id_����������";
        }

        private void FormViewOutputDoc_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            // ������� ���.
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
            // ������� �� ���������������.
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
            // ������� �� ���������������.
            if (this.radioButton2.Checked == true)
            {
                ViewDocCorrespondent();
            }

            // ������� �� ������������.
            if (this.radioButton3.Checked == true)
            {
                LoadData���������������������();
            }

            // ������� � �� ������������ � �� ���������������.
            if (this.radioButton4.Checked == true)
            {
                if (this.chkFiltrPerson.Checked == true)
                {
                    // id ���������.
                    int idPerson = (int)this.comboBox1.SelectedValue;

                    // ������� ���� ������������ �� ������� ������.
                    LoadData�����������������������������������(idPerson);
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        private void ViewDocCorrespondent()
        {
            int id = (int)this.comboBox1.SelectedValue;

            ������������������������ view = new ������������������������(rd);
            this.comboBox1.DataSource = view.Get��������();
            this.comboBox1.DisplayMember = "����������������������";
            this.comboBox1.ValueMember = "id_��������";

            this.dataGridView1.DataSource = view.���������������������������������(id);

            DisplayDataGrid();
        }

        /// <summary>
        /// ������� �� ������������.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            int id = (int)this.comboBox2.SelectedValue;

            ������������������������ view = new ������������������������(rd);
            this.dataGridView1.DataSource = view.������������������������������(id);

            DisplayDataGrid();

            // ��������� �������������� ������.
            this.comboBox1.DataSource = view.Get��������();
            this.comboBox1.DisplayMember = "�������";
            this.comboBox1.ValueMember = "id_��������";

            //// ������� ������������.
            this.comboBox2.DataSource = view.Get�����������();
            this.comboBox2.DisplayMember = "������������������";
            this.comboBox2.ValueMember = "id_����������";
        }

        /// <summary>
        /// ������ �� ������������.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if (this.radioButton3.Checked == true)
            {
                int id = (int)this.comboBox2.SelectedValue;

                ������������������������ view = new ������������������������(rd);

                // ��������� ������ ���������������.
                this.comboBox1.DataSource = view.Get����������������������(id);
                this.comboBox1.DisplayMember = "�������";
                this.comboBox1.ValueMember = "id_��������";

                // ��������� ������ ����������.
                this.dataGridView1.DataSource = view.������������������������������(id);

                DisplayDataGrid();

            }

            if (this.radioButton4.Checked == true)
            {
                if (this.chkCorr.Checked == true)
                {
                    // �������� ��������� � �������������� � ����������� �� ���������� �����������.
                    int id = (int)this.comboBox2.SelectedValue;

                    ������������������������ view = new ������������������������(rd);

                    this.comboBox1.DataSource = view.Get����������������������(id);
                    this.comboBox1.DisplayMember = "�������";
                    this.comboBox1.ValueMember = "id_��������";

                    this.dataGridView1.DataSource = null;
                }
            }

            //DisplayDataGrid();
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            // ������� �� ��������������� � ������������.
            if (this.radioButton4.Checked == true)
            {
                // ������� ������������.
                ������������������������ view = new ������������������������(rd);

                // ������� ���� ������������.
                this.comboBox2.DataSource = view.Get�����������();
                this.comboBox2.DisplayMember = "������������������";
                this.comboBox2.ValueMember = "id_����������";

                // ������� ���� ���������������.
                this.comboBox1.DataSource = view.Get��������();
                this.comboBox1.DisplayMember = "����������������������";
                this.comboBox1.ValueMember = "id_��������";

                // ������� ������ ���������� �� ������� ������ ������� 
                int idCorr = (int)this.comboBox1.SelectedValue;

                int idPers = (int)this.comboBox2.SelectedValue;

                //this.dataGridView1.DataSource = view.���������������������������������������(idPers, idCorr);
                //DisplayDataGrid();

                this.btnDisplay.Enabled = true;

                this.chkFiltrPerson.Checked = true;
                this.chkFiltrPerson.Enabled = true;
                this.chkCorr.Enabled = true;
                this.chkCorr.Checked = false;

                this.dataGridView1.DataSource = view.���������������������������������������(idPers, idCorr);
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
        /// �������� �� ������������ ������ ����������.
        /// </summary>
        private void DisplayDataGrid()
        {
            this.dataGridView1.Columns["id_��������"].Visible = false;
            this.dataGridView1.Columns["id_����������"].Visible = false;
        }

        //"SELECT [id_��������] " +
        //                  " ,id_���������� " +
        //                 " ,[�������] " +
        //                 " ,[�������������] " +
        //                 " ,[�������������] " +
        //                 " ,[�������������������] " +
        //                 " ,[������������������] " +
        //                 " ,[���������������] " +
        //                 " ,[����������] " +
        //                 " ,[������������������] " +

        private void chkFiltrPerson_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkFiltrPerson.Checked == true)
            {
                // ������ ����� �� ��������������.
                this.chkCorr.Checked = false;

                // ��������� ���� �� ������� ������ ���� ������������.
                // ������� ������������.
                ������������������������ view = new ������������������������(rd);

                // ������� ���� ������������.
                this.comboBox2.DataSource = view.Get�����������();
                this.comboBox2.DisplayMember = "������������������";
                this.comboBox2.ValueMember = "id_����������";

                // ������� ���� ���������������.
                this.comboBox1.DataSource = view.Get��������();
                this.comboBox1.DisplayMember = "����������������������";
                this.comboBox1.ValueMember = "id_��������";
            }
        }

        private void chkCorr_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chkCorr.Checked == true)
            {
                // ������ ���� �� ������������.
                this.chkFiltrPerson.Checked = false;

                // ������� ������������.
                ������������������������ view = new ������������������������(rd);

                // ������� ���� ������������.
                this.comboBox2.DataSource = view.Get�����������();
                this.comboBox2.DisplayMember = "������������������";
                this.comboBox2.ValueMember = "id_����������";

                // ������� ���� ���������������.
                this.comboBox1.DataSource = view.Get��������();
                this.comboBox1.DisplayMember = "����������������������";
                this.comboBox1.ValueMember = "id_��������";
            }
        }

        private void LoadData���������������������()
        {
            int id = (int)this.comboBox2.SelectedValue;

            ������������������������ view = new ������������������������(rd);

            // ��������� ������ ���������������.
            this.comboBox1.DataSource = view.Get����������������������(id);
            this.comboBox1.DisplayMember = "����������������������";
            this.comboBox1.ValueMember = "id_��������";

            // ��������� ������ ����������.
            this.dataGridView1.DataSource = view.������������������������������(id);
            DisplayDataGrid();
        }


        /// <summary>
        /// ������� ���� ������������ �� ��������� ���������������.
        /// </summary>
        private void LoadData�����������������������������������(int idPerson)
        {
            if (this.radioButton4.Checked == true)
            {

                //int id = (int)this.comboBox1.SelectedValue;

                ������������������������ view = new ������������������������(rd);
                //this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

                //DisplayDataGrid();

                // ������� ������������ �� ��������������.
                this.comboBox2.DataSource = null;
                this.comboBox2.DataSource = view.Get���������������������������(idPerson);
                this.comboBox2.DisplayMember = "������������������";
                this.comboBox2.ValueMember = "id_����������";

                this.dataGridView1.DataSource = null;
                //this.dataGridView1.DataSource = view.���������������������������������������(idPers, idCorr);
            }

        }

        /// <summary>
        /// ������� ���� ������������ �� ��������� ���������������.
        /// </summary>
        private void LoadData�����������������������������������(int idPerson)
        {
            if (this.radioButton4.Checked == true)
            {

                //int id = (int)this.comboBox1.SelectedValue;

                ������������������������ view = new ������������������������(rd);
                //this.dataGridView1.DataSource = view.GetTablePeriodDateIdCorr(id);

                //DisplayDataGrid();

                // ������� ������������ �� ��������������.
                this.comboBox1.DataSource = null;
                this.comboBox1.DataSource = view.Get���������������������������(idPerson);
                this.comboBox1.DisplayMember = "������������������";
                this.comboBox1.ValueMember = "id_��������";

                this.dataGridView1.DataSource = null;
                //this.dataGridView1.DataSource = view.���������������������������������������(idPers, idCorr);
            }

        }

        private void ���������������������������������_Id�����������()
        {
            // ������� ���������������.
            int idPers = (int)this.comboBox2.SelectedValue;
            int idCorr = (int)this.comboBox1.SelectedValue;

            ������������������������ view = new ������������������������(rd);
            //DisplayDataGrid();

            // ������� ��������������� �� �����������.
            this.comboBox1.DataSource = null;
            this.comboBox1.DataSource = view.Get����������������������(idPers);
            this.comboBox1.DisplayMember = "������������������";
            this.comboBox1.ValueMember = "id_����������";

            this.dataGridView1.DataSource = null;

            // ������� ��������� � ����������� �� ����������� � ��������������.
            this.dataGridView1.DataSource = view.���������������������������������������(idPers, idCorr);

            DisplayDataGrid();

        }

        private void btnDisplay_Click(object sender, EventArgs e)
        {
            if (this.radioButton4.Checked == true)
            {
                //���������� ��������� � ����������� �� ���������� �������������� � �����������.
                // ������� ���������������.
                int idPers = (int)this.comboBox2.SelectedValue;
                int idCorr = (int)this.comboBox1.SelectedValue;

                ������������������������ view = new ������������������������(rd);
                //DisplayDataGrid();

                //// ������� ��������������� �� �����������.
                //this.comboBox1.DataSource = null;
                //this.comboBox1.DataSource = view.Get����������������������(idPers);
                //this.comboBox1.DisplayMember = "������������������";
                //this.comboBox1.ValueMember = "id_����������";

                this.dataGridView1.DataSource = null;

                // ������� ��������� � ����������� �� ����������� � ��������������.
                this.dataGridView1.DataSource = view.���������������������������������������(idPers, idCorr);

                DisplayDataGrid();
            }

        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // ����������� ���������� DatagridView.
            //�������������������� report = new ��������������������();
            //report.DataGridView1 = this.dataGridView1;

            string caption = "����� �� ��������� ���������� �� ������ � " + rd.DataStart.ToShortDateString() + " �� " + rd.DataEnd.ToShortDateString();

            Report����������������������� reportDoc = new Report�����������������������(caption);

            Report������������������������� report = new Report�������������������������(reportDoc);
            report.ListDate = this.dataGridView1.Rows;
            report.Execute();

            //reportDoc.PrintReportStaticOutputDoc(this.dataGridView1.Rows);

            //Excel����� excel = new Excel�����();
            
            //excel.Print�����������������������(report, caption, "A1", "G1");


        }

        //private void DisplayDataGrid()
        //{
        //    this.dataGridView1.Columns["id_����������"].Visible = false;
        //    //this.dataGridView1.Columns["id_����������"].Visible = false;
        //}


         //"SELECT [id_��������] " +
         //                  " ,id_���������� " +
         //                 " ,[�������] " +
         //                 " ,[�������������] " +
         //                 " ,[�������������] " +
         //                 " ,[�������������������] " +
         //                 " ,[������������������] " +
         //                 " ,[���������������] " +
         //                 " ,[����������] " +
         //                 " ,[������������������] " +
         

          
    }
}