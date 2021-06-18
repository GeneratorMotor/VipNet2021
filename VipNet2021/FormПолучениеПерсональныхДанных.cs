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
    public partial class FormПолучениеПерсональныхДанных : Form
    {
        // Переменная для хранения id таблицы.
        private int id;

        // Флаг редактирования.
        private bool flagEdit = false;

        public FormПолучениеПерсональныхДанных()
        {
            InitializeComponent();
        }

        private void btsSave_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text != "")
            {
                // Активируем кнопку.
                this.btsSave.Enabled = true;
                string query = string.Empty;

                // Активируем кнопку удаления записи.
                this.btnDelete.Enabled = true;

                if (flagEdit == false)
                {
                    // Добавим новую запись.
                    query = "INSERT INTO [ЦельПолученияПерсональныхДанных] VALUES ('" + this.textBox1.Text.Trim() + "')";
                }
                else
                {
                    // Внесеём изменения в талицу СоставПерсональныхДанных.
                    query = "update ЦельПолученияПерсональныхДанных " +
                            "set ЦельПолученияПерсональныхДанных = '" + this.textBox1.Text.Trim() + "' " +
                            "where id_цельПолученияПерсДанных = " + id + " ";
                }

                ПодключитьБД connection = new ПодключитьБД();

                string sCon = connection.СтрокаПодключения();

                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();
                    SqlCommand com = new SqlCommand(query, con);

                    com.ExecuteNonQuery();
                }
                // Обнулим переменную для хранения id.
                id = 0;

                this.textBox1.Text = "";

                // Деакивируем кнопку добавления и уаления записи.
                this.btsSave.Enabled = false;
                this.btnDelete.Enabled = false;

                flagEdit = false;

                //SqlCommand com = new SqlCommand(insertQuery, sCon);
                //com.EndExecuteNonQuery();

                LoadDate();
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (this.textBox1.Text != "")
            {
                // Активируем кнопку.
                this.btsSave.Enabled = true;
                string query = string.Empty;

                // Активируем кнопку удаления записи.
                this.btnDelete.Enabled = true;

                query = "delete from dbo.ЦельПолученияПерсональныхДанных " +
                        "where id_цельПолученияПерсДанных = " + id + " ";

                ПодключитьБД connection = new ПодключитьБД();

                string sCon = connection.СтрокаПодключения();

                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();
                    SqlCommand com = new SqlCommand(query, con);

                    com.ExecuteNonQuery();
                }
                // Обнулим переменную для хранения id.
                id = 0;

                this.textBox1.Text = "";

                // Деакивируем кнопку добавления и уаления записи.
                this.btsSave.Enabled = false;
                this.btnDelete.Enabled = false;

                flagEdit = false;

                //SqlCommand com = new SqlCommand(insertQuery, sCon);
                //com.EndExecuteNonQuery();

                LoadDate();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // Получим текущую выделенную строку.
            DataGridViewRow row = this.dataGridView1.CurrentRow;

            // Получим название.
            this.textBox1.Text = row.Cells[1].Value.ToString().Trim();

            flagEdit = true;

            // Получим id строки редактирования.
            id = Convert.ToInt32(row.Cells[0].Value);
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

        private void LoadDate()
        {
            string query = "select * from ЦельПолученияПерсональныхДанных";

            GetDataTable getTable = new GetDataTable(query);

            DataTable tabPD = getTable.DataTable();

            this.dataGridView1.DataSource = tabPD;
            this.dataGridView1.Columns[0].Visible = false;
            this.dataGridView1.Columns[1].Width = 400;

        }

        private void FormПолучениеПерсональныхДанных_Load(object sender, EventArgs e)
        {
            LoadDate();
        }
    }
}