using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;
using RegKor.Classess2021;

namespace RegKor
{
    public partial class FormОснованиеПередачи : Form
    {
        private List<ОснованиеПередачи> listProperty = new List<ОснованиеПередачи>();
        private string stringQuery = string.Empty;

        /// <summary>
        /// Хранит список выбранных оснований для передачи персональных данных.
        /// </summary>
        public List<ОснованиеПередачи> ListОснованиеПередачи
        {
            get
            {
                return listProperty;
            }
            set
            {
                // Очистим коллекцию перед записью новых данных.
                //listProperty.Clear();
                listProperty = value;
            }
        }

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

        //private Dictionary<int, ОснованиеПередачи> dictionary = new Dictionary<int, ОснованиеПередачи>();

        /// <summary>
        /// Содержит SQL инструкцию на вставку.
        /// </summary>
        public string StringQuery
        {
            get
            {
                return stringQuery;
            }
            set
            {
                stringQuery = value;
            }
        }

        // Флаг указывает что форма работает с карточкой фходящей.
        private bool flagCardInpur = false;




        private bool flagEdit;

        /// <summary>
        /// Флаг указывает, что форма работает в режиме редактирования и что нужно показать основания для передачи персональных данных.
        /// </summary>
        public bool FlagEdit
        {
            get
            {
                return flagEdit;
            }

            set
            {
                flagEdit = value;
            }
        }

        private int idКарточкиВход = 0;

        // Флаг указывающий что форма в режиме добавления новой записи - false(true - режим редактирования).
        private bool flagEditPersonDate = false;

        /// <summary>
        /// Хранит id карточки входящих документов.
        /// </summary>
        //public int IdКарточкиВход
        //{
        //    get
        //    {
        //        return idКарточкиВход;
        //    }
        //    set
        //    {
        //        idКарточкиВход = value;
        //    }
        //}

    
        

        public FormОснованиеПередачи()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Получение id карточки входящей.
        /// </summary>
        /// <param name="idCardInput">id карточки входящей</param>
        public FormОснованиеПередачи(int idCardInput, bool flagEditPersonCard, bool flagCardInpur)
        {
            InitializeComponent();

            // Если id idCardInput > 0 значит работаем в режиме редактирования.
            //if (idCardInput > 0)
            //{
                this.idКарточкиВход = idCardInput;
            //}

            // Укажем что форма работает в режиме 
            this.flagEditPersonDate = flagEditPersonCard;

            // Укажем что форма работает с картчкой входящей.
            this.flagCardInpur = flagCardInpur;

            // Заполним данными DataGrid формы FormОснованиеПередачи.


        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormОснованиеПередачи_Load(object sender, EventArgs e)
        {
            // Заполним список данными.
            List<ОснованиеПередачи> list = new List<ОснованиеПередачи>();

            // Очистим свойство содержащее установленные основания для передачи перс данных.
            ListОснованиеПередачи.Clear();

            // Заполним форму данными.
            //string query = "SELECT [id_основаниеПередачи] " +
            //              ",[ОснованиеПередачи] " +
            //              "FROM [Основаниепередачи] where [ОснованиеПередачи] like '" + this.textBox1.Text.Trim() + "%'";

            // Получим SQL скрипт для заполнения данными формы ОснованиеПередачи.
            IQueryStringSQL queryBasisTrafic = new QueryОснованиеПередачи(this.textBox1.Text.Trim());
            string query = queryBasisTrafic.Query();

            // Строка подключения к SQL серверу.
            ПодключитьБД connect = new ПодключитьБД();
            string sConnect = connect.СтрокаПодключения();

            // Получим таблицу содержащую данные об онсованиях передачи.
            GetDataTable getTab = new GetDataTable(query);
            DataTable tab = getTab.DataTable("ОснованиеПередачи");

            // Преобразуем таблицу с основанием передачи данных в список.
            ConvertTableToList convert = new ConvertTableToList(tab, list);

            if (flagCardInpur == false)
            {

                //// Делаем не в единой транзакции 
                //string quer = "select * from Основаниепередачи " +
                //              "where id_основаниеПередачи in ( " +
                //              "SELECT [id_ОснованиеПередачи] " +
                //              "FROM [СвязующаяУчетаПерсональныхДанных ] " +
                //              "where id_карточки = " + IdКарточки + " )";

//                string quer = @"select * from Основаниепередачи 
//inner join СвязующаяУчетаПерсональныхДанных 
//on Основаниепередачи.id_основаниеПередачи = СвязующаяУчетаПерсональныхДанных.id_СоставПерсДанных
//where id_карточки = " + IdКарточки + " ";

                // Получаем скрипт на связь карточки с основанием передачи.
                IQueryStringSQL queryОснованиеПередачиСвязующее = new QueryСвязующаяУчетаПерсональныхДанных(IdКарточки);
                string quer = queryОснованиеПередачиСвязующее.Query();

                GetDataTable getTabC = new GetDataTable(quer);
                DataTable tabContr = getTabC.DataTable("ОснованиеПередачиКонтрольное");

                // Пройдем по списку оснований передачи и отметим выбранные пункты.
                foreach (ОснованиеПередачи item in convert.Get())
                {
                    DataRow[] rowSelect = tabContr.Select("id_ОснованиеПередачи = '" + item.Id_основаниеПередачи + "'");
                    if (rowSelect.Length != 0)
                    {
                        if (item.Id_основаниеПередачи == Convert.ToInt32(rowSelect[0]["id_ОснованиеПередачи"]))
                        {
                            item.FlagSelect = true;
                            //ListОснованиеПередачи.Add(item);
                        }
                    }
                    else
                    {
                        item.FlagSelect = false;
                    }
                }
            }
            else
            {
                // Получим SQl запрос для получения основания передачи.
                LoadQueryОснованиеПередачи loadQuery = new LoadQueryОснованиеПередачи(this.idКарточкиВход);
                string queryОснованиеПередачи = loadQuery.Query();

                GetDataTable getTabОснование = new GetDataTable(queryОснованиеПередачи);
                DataTable tabОснование = getTabОснование.DataTable("ОснованиеПередачи");

                foreach (ОснованиеПередачи item in convert.Get())
                {
                    DataRow[] rowSelect = tabОснование.Select("id_ОснованиеПередачи = '" + item.Id_основаниеПередачи + "'");
                    if (rowSelect.Length != 0)
                    {
                        if (item.Id_основаниеПередачи == Convert.ToInt32(rowSelect[0]["id_ОснованиеПередачи"]))
                        {
                            item.FlagSelect = true;
                            //ListОснованиеПередачи.Add(item);
                        }
                    }
                    else
                    {
                        item.FlagSelect = false;
                    }
                }
            }
            

            this.dataGridView1.DataSource = list;
            this.dataGridView1.Columns["Id_основаниеПередачи"].Visible = false;
            this.dataGridView1.Columns["Основание"].Width = 420;
            this.dataGridView1.Columns["FlagSelect"].Width = 150;
            this.dataGridView1.Columns["FlagSelect"].HeaderText = "Выбрать";

        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            //if (this.flagEditPersonDate == false)
            //{
                //List<ОснованиеПередачи> list = new List<ОснованиеПередачи>();

                // Строка для хранения SQL инструкции на вставку данных.
                StringBuilder build = new StringBuilder();

                // Заполним данными список.
                foreach (DataGridViewRow row in this.dataGridView1.Rows)
                {

                    string sT = row.Cells["Основание"].Value.ToString();
                    bool fl = row.Cells["FlagSelect"].Selected;

                    DataGridViewCheckBoxCell chk = (DataGridViewCheckBoxCell)row.Cells["FlagSelect"];
                    if (Convert.ToBoolean(row.Cells[2].Value) == true)
                    //if (chk.Value == true)
                    {
                        ОснованиеПередачи item = new ОснованиеПередачи();
                        item.Id_основаниеПередачи = Convert.ToInt32(row.Cells["Id_основаниеПередачи"].Value);
                        item.Основание = row.Cells["Основание"].Value.ToString().Trim();

                        // Поместим в список выбранное основание передачи персональных данных.
                        ListОснованиеПередачи.Add(item);

                        if (this.flagEditPersonDate == false && this.flagCardInpur == true)
                        {

                            string query = "INSERT INTO СвязующаяЦельПолучениперсональныхДанных " +
                                           "(id_карточки " +
                                           ",id_ОснованиеПередачи) " +
                                           "VALUES " +
                                           "( {0} " +
                                            ", " + Convert.ToInt32(row.Cells["Id_основаниеПередачи"].Value) + " )";

                            build.Append(query);
                        }
                       
                    }
                }

                StringQuery = build.ToString().Trim();
            //}

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            // Заполним форму данными.
            string query = "SELECT [id_основаниеПередачи] " +
                          ",[ОснованиеПередачи] " +
                          "FROM [Основаниепередачи] where [ОснованиеПередачи] like '%" + this.textBox1.Text.Trim() + "%'";

            ПодключитьБД connect = new ПодключитьБД();
            string sConnect = connect.СтрокаПодключения();

            // Получим таблицу содержащую данные об онсованиях передачи.
            GetDataTable getTab = new GetDataTable(query);
            DataTable tab = getTab.DataTable("ОснованиеПередачи");

          
            string quer = @"select * from Основаниепередачи 
                            inner join СвязующаяУчетаПерсональныхДанных 
                            on Основаниепередачи.id_основаниеПередачи = СвязующаяУчетаПерсональныхДанных.id_СоставПерсДанных
                            where id_карточки = " + IdКарточки + " ";

            GetDataTable getTabC = new GetDataTable(quer);
            DataTable tabContr = getTabC.DataTable("ОснованиеПередачиКонтрольное");

            // Заполним список данными.
            List<ОснованиеПередачи> list = new List<ОснованиеПередачи>();

            // Сконвертим таблицу данных в список с основаниями передач.
            ConvertTableToList convertTable = new ConvertTableToList(tab, list);

            // Список с основаниями передачи даных.
            List<ОснованиеПередачи> listSorted = convertTable.Get();

            // Пройдемся по списку содержащиму основания для передачи персональных данных.
            foreach (DataRow row in tab.Rows)
            {
                foreach (ОснованиеПередачи item in listSorted)
                {
                        // Узнаем отмеченные галочкой основания для передачи персональных данных.
                        DataRow[] rowSelect = tabContr.Select("id_ОснованиеПередачи = '" + item.Id_основаниеПередачи + "'");

                        // Если в таблице связей Карточки и ОснованиеПередачиДанных есть соответсвующие записи.
                        if (rowSelect.Length > 0)
                        {
                            // Если id основания передач равны.
                            if (item.Id_основаниеПередачи == Convert.ToInt32(rowSelect[0]["id_ОснованиеПередачи"]))
                            {
                                // Пометим Флаг выбора как True.
                                item.FlagSelect = true;
                            }
                        }
                        else
                        {
                            // Если связанных записей нет помечаем как False.
                            item.FlagSelect = false;
                        }
                }
            }


                // Пройдемся по строкам DataGridView.
                foreach (DataGridViewRow row in this.dataGridView1.Rows)
                {
                    // Сравним строки коллекции со строками DataGridView.
                    foreach (ОснованиеПередачи item in listSorted)
                    {
                        if (Convert.ToInt32(row.Cells["id_ОснованиеПередачи"].Value) == item.Id_основаниеПередачи)
                        {
                            // Пометим отмеченные галочкой в DataGridView элементы списка.
                            if (Convert.ToBoolean(row.Cells["FlagSelect"].Value) == true)
                            {
                                item.FlagSelect = true;
                            }
                            else
                            {
                                item.FlagSelect = false;
                            }
                        }
                    }
                }

            // Присвоим коллекцию как источник данных.
            this.dataGridView1.DataSource = list;

            // Отформатируем внешний вид DataGridView формы.
            this.dataGridView1.Columns["Id_основаниеПередачи"].Visible = false;
            this.dataGridView1.Columns["Основание"].Width = 420;
            this.dataGridView1.Columns["FlagSelect"].Width = 150;
            this.dataGridView1.Columns["FlagSelect"].HeaderText = "Выбрать";

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            ////отследим изменения в DatatGridView
            //DataGridViewCell col = this.dataGridView1.CurrentCell;

            //bool flag = Convert.ToBoolean(col.Value);

            ////Определим тип выбранной ячейки
            //Type t = col.GetType();

            ////если пользователь нажал checkbox 
            ////if (t.ToString() == "System.Windows.Forms.DataGridViewCheckBoxCell")
            //if (Convert.ToBoolean(this.dataGridView1.CurrentRow.Cells[2].Value) == true)
            ////if (flag == true)
            //{
            //    //================================пометсим полученное значение в словарь

            //    //получим id записи
            //    int id = Convert.ToInt32(this.dataGridView1.CurrentRow.Cells["id_основаниеПередачи"].Value);

            //    ОснованиеПередачи item = new ОснованиеПередачи();
            //    item.Id_основаниеПередачи = Convert.ToInt32(this.dataGridView1.CurrentRow.Cells[0].Value);
            //    item.Основание = this.dataGridView1.CurrentRow.Cells[1].Value.ToString().Trim();
            //    item.FlagSelect = flag;

            //    //запишем в словарь
            //    try
            //    {
            //        dictionary.Add(id, item);
            //    }
            //    catch
            //    {
            //        //Если пользователь снял ранее введённое значение значит мы удаляем ранее введённую строку 
            //        dictionary.Remove(id);

            //        // Введём повтороно данные.
            //        dictionary.Add(id, item);
            //    }

            //}
            //else
            //{
            //    //получим id записи
            //    int id = Convert.ToInt32(this.dataGridView1.CurrentRow.Cells["id_основаниеПередачи"].Value);

            //    // Так как мы сняли флаг с существующего элемента, то предпологается, что мы отказываемся от выбранного элемента.
            //    dictionary.Remove(id);

            //}

         
        }
    }
}