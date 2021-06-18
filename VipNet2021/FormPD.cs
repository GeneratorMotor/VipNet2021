using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;


namespace RegKor
{
    public partial class FormPD : Form
    {
        private bool flagEdit = false;

        private int id = 0;

        public FormPD()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormPD_Load(object sender, EventArgs e)
        {
            LoadDate();            
        }

        private void btsSave_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text != "")
            {
                // ���������� ������.
                this.btsSave.Enabled = true;
                string query = string.Empty;

                // ���������� ������ �������� ������.
                this.btnDelete.Enabled = true;

                if (flagEdit == false)
                {
                    // ������� ����� ������.
                    query = "INSERT INTO [�����������������] VALUES ('" + this.textBox1.Text.Trim() + "')";
                }
                else
                {
                    // ������ ��������� � ������ ������������������������.
                    query = "update ����������������� " +
                            "set ����������������� = '" + this.textBox1.Text.Trim() + "' " +
                            "where id_����������������� = " + id + " ";
                }

                ������������ connection = new ������������();

                string sCon = connection.�����������������();

                using(SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();
                    SqlCommand com = new SqlCommand(query, con);

                    com.ExecuteNonQuery();
                }
                // ������� ���������� ��� �������� id.
                id = 0;

                this.textBox1.Text = "";

                // ����������� ������ ���������� � ������� ������.
                this.btsSave.Enabled = false;
                this.btnDelete.Enabled = false;

                flagEdit = false;

                //SqlCommand com = new SqlCommand(insertQuery, sCon);
                //com.EndExecuteNonQuery();

                LoadDate();
            }

            // ������� ����� ��� ���������� ������������.
            this.Close();
        }

        /// <summary>
        /// ��������� �������� ������.
        /// </summary>
        private void LoadDate()
        {
            string query = "select * from �����������������";

            GetDataTable getTable = new GetDataTable(query);

            DataTable tabPD = getTable.DataTable();

            this.dataGridView1.DataSource = tabPD;
            this.dataGridView1.Columns[0].Visible = false;
            this.dataGridView1.Columns[1].Width = 400;
            
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (this.textBox1.Text != "")
            {
                this.btsSave.Enabled = true;
                this.btnDelete.Enabled = true;
            }
            else
            {
                this.btsSave.Enabled = false;
                this.btnDelete.Enabled = false;
            }

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // ������� ������� ���������� ������.
            DataGridViewRow row = this.dataGridView1.CurrentRow;

            // ������� ��������.
            this.textBox1.Text = row.Cells[1].Value.ToString().Trim();

            flagEdit = true;

            // ������� id ������ ��������������.
            if (row.Cells[0].Value != DBNull.Value)
            {
                id = Convert.ToInt32(row.Cells[0].Value);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text != "")
            {
                // ���������� ������.
                this.btsSave.Enabled = true;
                string query = string.Empty;

                // ���������� ������ �������� ������.
                this.btnDelete.Enabled = true;

                query = "delete from ����������������� " +
                        "where id_����������������� = " + id + " ";

                ������������ connection = new ������������();

                string sCon = connection.�����������������();

                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();
                    SqlCommand com = new SqlCommand(query, con);

                    com.ExecuteNonQuery();
                }
                // ������� ���������� ��� �������� id.
                id = 0;

                this.textBox1.Text = "";

                // ����������� ������ ���������� � ������� ������.
                this.btsSave.Enabled = false;
                this.btnDelete.Enabled = false;

                flagEdit = false;

                //SqlCommand com = new SqlCommand(insertQuery, sCon);
                //com.EndExecuteNonQuery();

                LoadDate();
            }
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}