using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;
using System.Data;
using System.Data.SqlClient;

namespace RegKor
{
    public partial class FormСправочникПодразделенияКомитета : Form
    {
        List<DepartmentPerson> list;

        private bool flagUpdate;

        private int idПодразд;

        /// <summary>
        /// Хранит Id подразделения.
        /// </summary>
        public int IdПодразделения
        {
            get
            {
                return idПодразд;
            }
            set
            {
                idПодразд = value;
            }
        }

        /// <summary>
        /// Флаг указывает что форма работает в режиме редактирования.
        /// </summary>
        public bool FlagUpdate
        {
            get
            {
                return flagUpdate;
            }
            set
            {
                flagUpdate = value;
            }
        }

        public FormСправочникПодразделенияКомитета()
        {
            InitializeComponent();

            list = new List<DepartmentPerson>();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormСправочникПодразделенияКомитета_Load(object sender, EventArgs e)
        {

            if (this.FlagUpdate == true)
            {
                this.btnAdd.Enabled = false;


                string query = "SELECT     dbo.ПодразделенияКомитета.ОписаниеПодразделения, dbo.ПодразделенияКомитета.НомерПодразделения,dbo.ПодразделенияКомитета.БуквенноеОбозначение, " +
                               " dbo.Получатели.ОписаниеПолучателя, dbo.Получатели.id_получателя " +
                               " FROM         dbo.ПодразделенияКомитета INNER JOIN " +
                               " dbo.Получатели ON dbo.ПодразделенияКомитета.id_РуководителяПодразделения = dbo.Получатели.id_получателя " +
                               " where  ПодразделенияКомитета.Удален = 'False' and ПодразделенияКомитета.ФлагДействующий = 'True' " +
                               " and id_подразделения = "+ this.IdПодразделения +" ";

                DataTable tab = DataTableSql.GetDataTable(query);

                // Выведим в поле редактирования данные по руководителю и по подразделению.
                this.lblId.Text = tab.Rows[0]["id_получателя"].ToString();

                this.txtNamePerson.Text = tab.Rows[0]["ОписаниеПолучателя"].ToString();

                this.txtNameDepartment.Text = tab.Rows[0]["ОписаниеПодразделения"].ToString();

                this.txtNumDepartment.Text = tab.Rows[0]["НомерПодразделения"].ToString();

                this.txtLiter.Text = tab.Rows[0]["БуквенноеОбозначение"].ToString();


            }

            //this.radioButton1.Checked = true;

            ////// Отобразим список руководителей подразделения.
            //string query = " select id_получателя,ОписаниеПолучателя from dbo.Получатели " +
            //               " where Удален is null ";
        }

        private void btnAddList_Click(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FormСписокСотрудников form = new FormСписокСотрудников();
            DialogResult result = form.ShowDialog(this);

            if (result == DialogResult.OK)
            {

                string query = "SELECT [id_получателя] ,[ОписаниеПолучателя] FROM [Получатели] where Удален is null and id_получателя = "+ form.ИДСотрудника +" ";

                DataRowCollection rows = DataTableSql.GetDataTable(query).Rows;
                if (rows.Count > 0)
                {
                    this.lblId.Text = rows[0]["id_получателя"].ToString().Trim();
                    this.txtNamePerson.Text = rows[0]["ОписаниеПолучателя"].ToString().Trim();
                }

            }
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (this.txtNamePerson.Text.Length > 0 && this.txtNumDepartment.Text.Length > 0 && this.txtNameDepartment.Text.Length > 0)
            {
                string queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          "INSERT INTO [ПодразделенияКомитета] " +
                                           " ([ОписаниеПодразделения] " +
                                           ",[id_РуководителяПодразделения] " +
                                           ",[НомерПодразделения] " +
                                           ",[БуквенноеОбозначение] " +
                                           ",[Удален] " +
                                           ",[ФлагДействующий]) " +
                                           " VALUES " +
                                           "('" + this.txtNameDepartment.Text + "' " +
                                           "," + Convert.ToInt32(this.lblId.Text) + " " +
                                           ",'" + this.txtNumDepartment.Text + "' " +
                                           ",'" + this.txtLiter.Text + "' " +
                                           ",'False' " +
                                           ",'True') " +
                                          "COMMIT TRANSACTION ";

                // Сохраним данные.
                ПодключитьБД connectBD = new ПодключитьБД();
                string sCon = connectBD.СтрокаПодключения();

                // Выполним запрос на вставку (к сожалению не в единой транзакции.
                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();
                    SqlCommand com = new SqlCommand(queryInsert.ToString().Trim(), con);
                    com.ExecuteNonQuery();
                }

                //// Отобразим в DataGridView.
                //string query = "SELECT     dbo.ПодразделенияКомитета.ОписаниеПодразделения, dbo.ПодразделенияКомитета.НомерПодразделения,  " +
                //           " dbo.Получатели.ОписаниеПолучателя " +
                //           " FROM         dbo.ПодразделенияКомитета INNER JOIN  " +
                //           " dbo.Получатели ON dbo.ПодразделенияКомитета.id_РуководителяПодразделения = dbo.Получатели.id_получателя " +
                //           "where  ПодразделенияКомитета.Удален = 'False' and ПодразделенияКомитета.ФлагДействующий = 'True' ";

                //DataTable tab = DataTableSql.GetDataTable(query);
                //this.dataGridView1.DataSource = tab;


            }


        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //int id = this.dataGridView1.CurrentRow.Cells[""]
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (this.txtNamePerson.Text.Length > 0 && this.txtNumDepartment.Text.Length > 0 && this.txtNameDepartment.Text.Length > 0)
            {
                string queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          " UPDATE [ПодразделенияКомитета] " +
                                          " SET [ОписаниеПодразделения] = '" + this.txtNameDepartment.Text + "' " +
                                          " ,[id_РуководителяПодразделения] = " + Convert.ToInt32(this.lblId.Text) + " " +
                                          " ,[НомерПодразделения] = '" + this.txtNumDepartment.Text + "' " +
                                          " ,[БуквенноеОбозначение] = '" + this.txtLiter.Text + "' " +
                                          //" ,[Удален] =  'False'  " +
                                          //" ,[ФлагДействующий] = 'True' " +
                                          " ,[Дата] = '" + ДатаSQL.Дата(DateTime.Today.ToShortDateString()) + "' " +
                                          " WHERE id_подразделения = " + this.IdПодразделения + " " +
                                          "COMMIT TRANSACTION ";

                // Сохраним данные.
                ПодключитьБД connectBD = new ПодключитьБД();
                string sCon = connectBD.СтрокаПодключения();

                // Выполним запрос на вставку (к сожалению не в единой транзакции.
                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();
                    SqlCommand com = new SqlCommand(queryInsert.ToString().Trim(), con);
                    com.ExecuteNonQuery();
                }
            }
        }
    }
}