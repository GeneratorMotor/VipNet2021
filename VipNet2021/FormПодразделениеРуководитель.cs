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
    public partial class FormПодразделениеРуководитель : Form
    {
        string date = string.Empty;
        private int id;

        public FormПодразделениеРуководитель(string дата)
        {
            InitializeComponent();

            date = дата;
        }

        private void FormПодразделениеРуководитель_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void LoadData()
        {
            string query = "SELECT   dbo.ПодразделенияКомитета.id_подразделения,  dbo.ПодразделенияКомитета.ОписаниеПодразделения, dbo.ПодразделенияКомитета.НомерПодразделения,  " +
                          " dbo.Получатели.ОписаниеПолучателя " +
                          " FROM         dbo.ПодразделенияКомитета INNER JOIN  " +
                          " dbo.Получатели ON dbo.ПодразделенияКомитета.id_РуководителяПодразделения = dbo.Получатели.id_получателя " +
                          "where  ПодразделенияКомитета.Удален = 'False' and ПодразделенияКомитета.ФлагДействующий = 'True' ";

            DataTable tab = DataTableSql.GetDataTable(query);
            this.dataGridView1.DataSource = tab;

            if (tab.Rows.Count > 0)
            {
                id = Convert.ToInt32(this.dataGridView1.CurrentRow.Cells["id_подразделения"].Value);
            }

        }


        private void btnAdd_Click(object sender, EventArgs e)
        {
           DialogResult rezult= MessageBox.Show("Установить новую структуру подразделений? ","Внимание",MessageBoxButtons.OKCancel,MessageBoxIcon.Exclamation);
           
           if (rezult == DialogResult.OK)
           {
               string queryUpdate = "update dbo.ПодразделенияКомитета set ФлагДействующий = 'False'";

               // Сохраним данные.
               ПодключитьБД connectBD = new ПодключитьБД();
               string sCon = connectBD.СтрокаПодключения();

               // Выполним запрос на вставку (к сожалению не в единой транзакции.
               using (SqlConnection con = new SqlConnection(sCon))
               {
                   con.Open();
                   SqlCommand com = new SqlCommand(queryUpdate.ToString().Trim(), con);
                   com.ExecuteNonQuery();
               }

               LoadData();

           }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            FormСправочникПодразделенияКомитета form = new FormСправочникПодразделенияКомитета();
            form.IdПодразделения = id;
            form.FlagUpdate = true;
            form.ShowDialog();

            if (form.DialogResult == DialogResult.Yes)
            {
                LoadData();
            }
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
            // Переборсим в форму справочник подразделений комитета.
            id = Convert.ToInt32(this.dataGridView1.CurrentRow.Cells["id_подразделения"].Value);

        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
             DialogResult result = MessageBox.Show("Удалить запись", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
           
            if(result == DialogResult.Yes)
            {
                string queryDelete = " SET TRANSACTION ISOLATION LEVEL serializable begin transaction delete dbo.ПодразделенияКомитета " +
                                     "where id_подразделения = " + id + " " +
                                     "COMMIT TRANSACTION ";

                ПодключитьБД connectBD = new ПодключитьБД();
                string sCon = connectBD.СтрокаПодключения();

                // Выполним запрос на вставку (к сожалению не в единой транзакции.
                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();
                    SqlCommand com = new SqlCommand(queryDelete.ToString().Trim(), con);
                    com.ExecuteNonQuery();
                }

                LoadData();

                

            }
        }

        private void bAdd_Click(object sender, EventArgs e)
        {
            FormСправочникПодразделенияКомитета form = new FormСправочникПодразделенияКомитета();
            form.FlagUpdate = false;
            form.ShowDialog();

            if (form.DialogResult == DialogResult.OK)
            {
                LoadData();
            }
        }

        private void новаяСтруктураПодразделенияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult rezult = MessageBox.Show("Установить новую структуру подразделений? ", "Внимание", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);

            if (rezult == DialogResult.OK)
            {
                string queryUpdate = "update dbo.ПодразделенияКомитета set ФлагДействующий = 'False'";

                // Сохраним данные.
                ПодключитьБД connectBD = new ПодключитьБД();
                string sCon = connectBD.СтрокаПодключения();

                // Выполним запрос на вставку (к сожалению не в единой транзакции.
                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();
                    SqlCommand com = new SqlCommand(queryUpdate.ToString().Trim(), con);
                    com.ExecuteNonQuery();
                }

                LoadData();

            }
        }
    }
}