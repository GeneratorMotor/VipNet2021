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
    public partial class FormПерсональныеДанные : Form
    {

        private DatePersonal dPerson = new DatePersonal();
        private int idКарточки = 0;

        /// <summary>
        /// Хранит id карточки.
        /// </summary>
        public int IdКарточки
        {
            get
            {
                return idКарточки;
            }
            set
            {
                idКарточки = value;
            }
        }

        /// <summary>
        /// Содержит данные для конфигурирования персональных данных.
        /// </summary>
        public DatePersonal КонфигурированиеПерсональныхДанных
        {
            get
            {
                return dPerson;
            }
            set
            {
                dPerson = value;
            }
        }

        public FormПерсональныеДанные()
        {
            InitializeComponent();
        }

        private void FormПерсональныеДанные_Load(object sender, EventArgs e)
        {
            // Заполним поле CheckedListBox===========
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
           
            // Заполним раскрывающийся список - цель получения персональных данных.
            // Заполним раскрывающийся список: цель получения персональных данных информацией из базы данных.
            // Поучим срдерждимое справочника персональных данных.
            ПодключитьБД coonectDB2 = new ПодключитьБД();
            string sConn2 = coonectDB2.СтрокаПодключения();

            string query2 = "select [id_цельПолученияПерсДанных],[ЦельПолученияПерсональныхДанных] from ЦельПолученияПерсональныхДанных";

            GetDataTable getTable2 = new GetDataTable(query2);
            DataTable tabPD2 = getTable2.DataTable();

            this.cmbBox.DataSource = tabPD2;
            this.cmbBox.DisplayMember = "ЦельПолученияПерсональныхДанных";
            //this.cmbBox.ValueMember = "id_цельПолученияПерсДанных";



        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void radioButton1_Click(object sender, EventArgs e)
        {
            this.btnSave.Enabled = true;
            this.textBox1.Enabled = false;

            // Запишем причину отказа с коменнтарием.
            ОтметкаПередачаОтказ отметка = new ОтметкаПередачаОтказ();
            отметка.Отметка = true;
            отметка.ПричиныОтказа = "NULL";

            КонфигурированиеПерсональныхДанных.ОтметкаОтказПередача = отметка;
        }

        private void radioButton2_Click(object sender, EventArgs e)
        {
            this.btnSave.Enabled = true;
            this.textBox1.Enabled = true;

            // Запишем причину отказа с коментарием.
            ОтметкаПередачаОтказ отметка = new ОтметкаПередачаОтказ();
            отметка.Отметка = false;
            отметка.ПричиныОтказа = this.textBox1.Text.Trim();

            КонфигурированиеПерсональныхДанных.ОтметкаОтказПередача = отметка;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Запишем в свойство отконфигурированные сведения о персональных данных.
            // Переменная для хранения строки запроса.
            StringBuilder builder = new StringBuilder();

            // Сохраним изминия в БД.
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    // Получим название вида персональных данных отмеченное галочкой.
                    string val = checkedListBox1.Items[i].ToString();

                    // Получим id из таблицы состав персональных данных.
                    string queryIdПерсДанных = "select id_СоставПерсДанных  from СоставПерсональныхДанных " +
                                               "where СоставПерсональныхДанных = '" + val + "'";

                    // Получим id персональных данных.
                    GetDataTable getTable = new GetDataTable(queryIdПерсДанных);
                    int id_СоставПерсДанных = Convert.ToInt32(getTable.DataTable().Rows[0][0]);

                    // Получим id карточки.
                    int id_Карточки = this.IdКарточки;

                    // Сохраним в единой транзакции данные для связывающих таблицы.
                    string qInsert = "INSERT INTO [СвязующаяУчетаПерсональныхДанных] " +
                                     "([id_карточки] " +
                                     ",[id_СоставПерсДанных]) " +
                                     "VALUES " +
                                     "(" + id_Карточки + " " +
                                     "," + id_СоставПерсДанных + " ) ";

                    builder.Append(qInsert);
                }

            }


            string sQuery = builder.ToString();

            // Здесь пробел добавлен в конец строки для того чтобы SQL Server не сгинерил ошибку.
            КонфигурированиеПерсональныхДанных.СотавПерсональныхДанных = sQuery + " ";

            // Передадим в свойство id Цель получения данных.
            ПодключитьБД coonectDB = new ПодключитьБД();
            string sConn = coonectDB.СтрокаПодключения();

            string query = "select [id_цельПолученияПерсДанных]from ЦельПолученияПерсональныхДанных " +
                           "where [ЦельПолученияПерсональныхДанных] = '" + this.cmbBox.Text + "' ";

            GetDataTable getTable2 = new GetDataTable(query);
            DataTable tabPD = getTable2.DataTable();

            КонфигурированиеПерсональныхДанных.IdЦельПолученияПерсональныхДанных = Convert.ToInt32(tabPD.Rows[0][0]);

            // Получим Отметка отказ.
            ОтметкаПередачаОтказ отм = new ОтметкаПередачаОтказ();

            // Если дан полоэительный ответ.
            if(this.radioButton1.Checked == true)
            {
                отм.Отметка = true;
                отм.ПричиныОтказа = "NULL";
            }

            // Если дан отрицательный ответ.
            // Если дан полоэительный ответ.
            if (this.radioButton2.Checked == true)
            {
                отм.Отметка = false;
                отм.ПричиныОтказа = this.textBox1.Text.Trim();
            }

            КонфигурированиеПерсональныхДанных.ОтметкаОтказПередача = отм;
        }
    }
}