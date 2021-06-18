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
    public partial class FormСоставПерсДанных : Form
    {
        private int idКарточкиProperti;
        string queryСвязующаяУчётПерсональныхДанных = string.Empty;

        /// <summary>
        /// Свойство хранит id карточки.
        /// </summary>
        public int Idкарточки
        {
            get
            {
                return idКарточкиProperti;
            }
            set
            {
                idКарточкиProperti = value;
            }
        }

         
        /// <summary>
        /// Хранит строку запроса для внесения информации в связующую таблицу Учёт Персональных Данных.
        /// </summary>
        public string СвязующаяУчётПерсональныхДанных
        {
            get
            {
                return queryСвязующаяУчётПерсональныхДанных;
            }
            set
            {
                queryСвязующаяУчётПерсональныхДанных = value;
            }
        }

        public FormСоставПерсДанных()
        {
            InitializeComponent();
        }



        private void FormСоставПерсДанных_Load(object sender, EventArgs e)
        {
            // Поучим срдерждимое справочника персональных данных.
            ПодключитьБД coonectDB = new ПодключитьБД();
            string sConn = coonectDB.СтрокаПодключения();

            string query = "select СоставПерсональныхДанных from СоставПерсональныхДанных";

            GetDataTable getTable = new GetDataTable(query);
            DataTable tabPD = getTable.DataTable();

            //this.checkedListBox1.DataSource = tabPD;
            foreach (DataRow s in tabPD.Rows)
            {
                this.checkedListBox1.Items.Add(s[0]);
            }
            //this.dataGridView1.Columns[0].Visible = false;
            //this.dataGridView1.Columns[1].Width = 400;



           
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Переменная для хранения строки запроса.
            StringBuilder builder = new StringBuilder();

            // Сохраним изминия в БД.
            for(int i = 0; i< checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    // Получим название вида персональных данных отмеченное галочкой.
                    string val = checkedListBox1.Items[i].ToString();

                    // Получим id из таблицы состав персональных данных.
                    string queryIdПерсДанных = "select id_СоставПерсДанных  from СоставПерсональныхДанных " +
                                               "where СоставПерсональныхДанных = '"+ val +"'";

                    // Получим id персональных данных.
                    GetDataTable getTable = new GetDataTable(queryIdПерсДанных);
                    int id_СоставПерсДанных =  Convert.ToInt32(getTable.DataTable().Rows[0][0]);

                    // Получим id карточки.
                    int id_Карточки = this.Idкарточки;

                    // Сохраним в единой транзакции данные для связывающих таблицы.
                    string qInsert = "INSERT INTO [СвязующаяУчетаПерсональныхДанных] " +
                                     "([id_карточки] " +
                                     ",[id_СоставПерсДанных]) " +
                                     "VALUES " +
                                     "("+ id_Карточки +" " +
                                     "," + id_СоставПерсДанных + " ) ";

                    builder.Append(qInsert);
                }

            }


            string sQuery = builder.ToString();

            // Передадим в свойство формы запрос на внесение изменений в связующую таблицу СвязующаяУчётПерсональныхДанных.
            this.СвязующаяУчётПерсональныхДанных = sQuery;

            //// Сохраним данные.
            //ПодключитьБД connectBD = new ПодключитьБД();
            //string sCon = connectBD.СтрокаПодключения();

            //using (SqlConnection con = new SqlConnection(sCon))
            //{
            //    con.Open();
            //    SqlCommand com = new SqlCommand(sQuery, con);
            //    com.ExecuteNonQuery();
            //}


            this.Close();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }
    }
}