using System;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.IO;
using System.Collections.Generic;
using System.Configuration;
using System.ServiceProcess;
using System.Data.SqlClient;
using RegKor.Classess2021;

using RegKor.Classess;

namespace RegKor
{
    /// <summary>
    /// Summary description for FormКарточка.
    /// </summary>
    public class FormКарточка : System.Windows.Forms.Form
    {
        public RegKor.DS1.КарточкаRow строкаКарточки;
        private System.Windows.Forms.Label labelКорреспID;
        private System.Windows.Forms.Label labelРезультатВыполнения;
        private System.Windows.Forms.Label labelРезолюция;
        private System.Windows.Forms.Label labelСодержание;
        private System.Windows.Forms.Label labelДокумент;
        private System.Windows.Forms.Label labelКорреспондент;
        public System.Windows.Forms.CheckBox checkBoxВДеле;
        private System.Windows.Forms.Button buttonОтмена;
        private System.Windows.Forms.Button buttonСохранить;
        private System.Windows.Forms.GroupBox groupBoxОтправлено;
        private System.Windows.Forms.Label labelНомерИсходящий;
        private System.Windows.Forms.Label labelДатаОтправления;
        public System.Windows.Forms.DateTimePicker dateTimeДатаОтправления;
        public System.Windows.Forms.TextBox textBoxНомерИсходящий;
        public System.Windows.Forms.TextBox textBoxРезолюция;
        public System.Windows.Forms.TextBox textBoxСодержание;
        public System.Windows.Forms.DateTimePicker dateTimeКонтроль;
        public System.Windows.Forms.CheckBox checkBoxКонтроль;
        public System.Windows.Forms.ComboBox comboДокумент;
        public System.Windows.Forms.ComboBox comboКорреспондент;
        private System.Windows.Forms.GroupBox groupBoxПоступило;
        private System.Windows.Forms.Label labelНомерВходящий;
        private System.Windows.Forms.Label labelДатаПоступления;
        public System.Windows.Forms.DateTimePicker dateTimeДатаПоступления;
        private RegKor.DS1 ds1;
        public System.Windows.Forms.TextBox textBoxРезультатВыполнения;
        private System.Windows.Forms.Panel panelKontrol;
        private string старыйТекстРезультата = "";
        private System.Windows.Forms.Button buttonДобавитьПолучателей;
        private System.Windows.Forms.ToolTip toolTip1;
        private Panel panel1;
        private System.ComponentModel.IContainer components;
        private Label label1;

        // Текущий год.
        private int currentYear;

        /// <summary>
        /// Номер порядковый, который должен быть вставлен в базу
        /// </summary>
        private int следНомерПП;

        /// <summary>
        /// Номер порядковый для сохранения
        /// </summary>
        private int номерПП = 0;
        private Button btnElementPS;

        /// <summary>
        /// Это новый документ или изменяется существующий
        /// </summary>
        private bool новыйДокумент = false;

        private int idКарточкиProperti;

        /// <summary>
        /// Хранит выбранный год.
        /// </summary>
        public int CurrentYear
        {
            get
            {
                return currentYear;
            }
            set
            {
                currentYear = value;
            }
        }

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
        private string queryStringProperty = string.Empty;

        /// <summary>
        /// Свойство хранит строку запроса для заполнения связующей таблицы УчётПерсональныхДанных.
        /// </summary>
        public string QueryStringУчётПерсДанных
        {
            get
            {
                return queryStringProperty;
            }
            set
            {
                queryStringProperty = value;
            }
        }

        int цель = 0;

        /// <summary>
        /// Хранит id таблицы Цель получения персональных данных.
        /// </summary>
        public int IdЦельПолученияПерсональныхДанных
        {
            get
            {
                return цель;
            }
            set
            {
                цель = value;
            }
        }

        bool _отметка;

        /// <summary>
        ///  Хранит отметку о передаче или отказе в передаче персональных данных.
        /// </summary>
        public bool ОтметкаПередачаОтказ
        {
            get
            {
                return _отметка;
            }
            set
            {
                _отметка = value;
            }
        }

        private ОтметкаПередачаОтказ передачаОтказ;
        private CheckBox chBoxRepet;

        /// <summary>
        /// Хранит передачу или отказ с указанием причины отказа.
        /// </summary>
        public ОтметкаПередачаОтказ ПередачаОтказ
        {
            get
            {
                return передачаОтказ;
            }
            set
            {
                передачаОтказ = value;
            }
        }

        private DatePersonal dPerson = new DatePersonal();
        private Label label2;
        private TextBox txtPeriod;

        /// <summary>
        /// Хранит сконфигурированные персональные данные.
        /// </summary>
        public DatePersonal ConfigDatePerosnal
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

        private bool flagRecordRepet;
        private CheckBox chekDocServer;

        /// <summary>
        /// Хранит состояние указывающее что запись будет иметь повторяющейя ответ.
        /// </summary>
        public bool FlagRecordRepeet
        {
            get
            {
                return flagRecordRepet;
            }
            set
            {
                flagRecordRepet = value;
            }
        }

        private int increment;
        private Button btnTie;

        /// <summary>
        /// Хранит приращение даты.
        /// </summary>
        public int IncrementDate
        {
            get
            {
                return increment;
            }
            set
            {
                increment = value;
            }
        }
        private bool saveDocServer;
        /// <summary>
        /// Флаг указывает что сканкопия документа сохраняется на сервере.
        /// </summary>
        public bool SaveDocServer
        {
            get
            {
                return saveDocServer;
            }
            set
            {
                saveDocServer = value;
            }

        }

        //Переменная для хранения пути к папке с документами.
        private string pathFileServer = string.Empty;

        // Переменная для хранения пути к титульному листу.
        private string pathFileServerTitlePage = string.Empty;

        /// <summary>
        /// Хранит путь к папке с документами на локальной машине.
        /// </summary>
        public string PathFileServer
        {
            get
            {
                return pathFileServer;
            }
            set
            {
                pathFileServer = value;
            }
        }

        // Переменные для хранения имени файла который копируется на сервер.
        private string fileName = string.Empty;
        private string fileNameCopy = string.Empty;

        /// <summary>
        /// Содержит имя файла который будем архивировать.
        /// </summary>
        public string FileName
        {
            get
            {
                return fileName;
            }
            set
            {
                fileName = value;
            }
        }


        // Переменная для хранения следующего номера документа.
        private int следующийНомерДокумента;
        private LinkLabel linkLabel1;

        // Переменная которая хранит следующий порядковый номер документа.
        private string lastNumberDoc = string.Empty;

        // Флаг указывает что используется подключение нормальной нумерации документа.
        private bool flagLastNumberDoc = false;

        // Экземпляр класса описывающий следующий номер документа.
        НомерДокумента numDoc = new НомерДокумента();

        /// <summary>
        /// Хранит номер вновь зарегистрированного документа.
        /// </summary>
        public НомерДокумента СледующийНомерДокумента
        {
            get
            {
                return numDoc;
            }
            set
            {
                numDoc = value;
            }
        }

        private string имяДокумена = string.Empty;

        /// <summary>
        /// Хранит номер документа прописанного в БД.
        /// </summary>
        public string ИмяДокумента
        {
            get
            {
                return имяДокумена;
            }
            set
            {
                имяДокумена = value;
            }
        }

        private string имяТитульногоЛиста = string.Empty;

        /// <summary>
        /// Хранит титульный лист.
        /// </summary>
        public string ИмяТитульногоЛиста
        {
            get
            {
                return имяТитульногоЛиста;
            }
            set
            {
                имяТитульногоЛиста = value;
            }
        }

        // Переменная для хранения пути к серверу для копирования файла на клиент.
        private string patchServerSave = string.Empty;

        // Переменная для хранения имени файла на сервере.
        private string fileNameServer = string.Empty;

        private bool флагЗаписиАрхива;
        private MaskedTextBox textBoxНомерВходящий;
        private CheckBox chcDop;

        /// <summary>
        /// Свойство указывает, что архив с документом можно записывать на сервер.
        /// </summary>
        public bool ФлагЗаписиАрхива
        {
            get
            {
                return флагЗаписиАрхива;
            }
            set
            {
                флагЗаписиАрхива = value;
            }
        }


        private bool _flagUpdateRecord;

        /// <summary>
        /// Свойство обновления записи.
        /// </summary>
        public bool FlagUpdateDocument
        {
            get
            {
                return _flagUpdateRecord;
            }
            set
            {
                _flagUpdateRecord = value;
            }
        }

        private bool flagAddDoc = false;


        /// <summary>
        /// Свойство определяет что в документ в будущем возможно добавление листов.
        /// </summary>
        public bool FlagAddDoc
        {
            get
            {
                return flagAddDoc;
            }
            set
            {
                flagAddDoc = value;
            }
        }

        // Переменная для хранения выбранного способа получения документа.
        private ItemСпособПоступленияДокумента item;
        private LinkLabel linkLabel2;

        /// <summary>
        /// Возвращает выбранный спосб поступления документа.
        /// </summary>
        public ItemСпособПоступленияДокумента СпособПоступления
        {
            get
            {
                return item;
            }
            set
            {
                item = value;
            }
        }

        private List<PersonRecepient> listPerson;

        /// <summary>
        /// Свойство хранит список льготников которые выбраны для получения отписанного документа.
        /// </summary>
        public List<PersonRecepient> ListPerson
        {
            get
            {
                return listPerson;
            }
            set
            {
                listPerson = value;
            }
        }

        private byte[] fileByteArray;

        private TextBox textBoxНомерВходящий2;

        // Переменная указывает что счётчик номеров документов нужно остановить.
        private bool flagAutoNumberDocStoip = false;
        private CheckBox chboxDsp;

        // Переменная для хранения номера входящего документа.
        private string sNumStart = string.Empty;

        private string flagDsp = "False";

        /// <summary>
        /// Хранит состояние является документ ДСП или нет.
        /// </summary>
        public string FlagDsp
        {
            get
            {
                return flagDsp;
            }
            set
            {
                flagDsp = value;
            }
        }

        private Dictionary<string, string> numbersDepartment = new Dictionary<string, string>();
        private RadioButton rb04;
        private RadioButton rb02;
        private Button btnLastNumber;

        // Перемення для хранения id основания передачи персональных данных.
        private int idPersonDate = 0;

        // Свойство для хранения id основания передачи персональных данных.
        public int IdPersonDate
        {
            get
            {
                return idPersonDate;
            }
            set
            {
                idPersonDate = value;
            }
        }

        // Переменная для хранения строки запроса к персональным данным для карточки входящей.
        private string queryPersonDateForCardInput = string.Empty;

        /// <summary>
        /// Переменная для хранения строки запроса к персональным данным для карточки входящей.
        /// </summary>
        public string QueryPersonDateForCardInput
        {
            get
            {
                return queryPersonDateForCardInput;
            }
            set
            {
                queryPersonDateForCardInput = value;
            }
        }

       



        

        /// <summary>
        /// Переменная для хранения полного номер карточки.
        /// </summary>
        private string _НомерДокументаПовторПривязка = string.Empty;

        /// <summary>
        /// Конструктор формы. В качестве параметра принимает датасет. Используется для создания новой записи
        /// </summary>
        /// <param name="ds">Датасет</param>
        public FormКарточка(RegKor.DS1 ds, string выбранныйГод, bool flagAutoStop)
        {
            InitializeComponent();

            // Запретим или разрешим автоматическое увеличение счётчика номеров документов.
            flagAutoNumberDocStoip = flagAutoStop;

            this.ds1 = ds;
            строкаКарточки = ds1.Карточка.NewКарточкаRow();

            // Получим id карточки.
            int id_карточки = строкаКарточки.id_карточки;

            // Передадим id катрочки в свойство формы.
            this.Idкарточки = id_карточки;

            новыйДокумент = true;

            string префиксДокумента = string.Empty;

            // Загрузим список номеров подразделений комитета.
            LoadNumberDepartments();

            
            // Здесь я не мудрил поставил порядковый номер НомерПП в таблице Карточка в ручную 1 сразу после нового года.

            // Получение максимального номераПП из таблицы КарточкаИсходящая
            //DataRow[] dr = ds.Карточка.Select("ДатаПоступ >='01.12." + выбранныйГод + "'", "НомерПП DESC");

            //string query = "declare @numDoc int  " +
            //               " select top 1 @numDoc = номерПП from Карточка " +
            //                " where ДатаИсхода >= '" + выбранныйГод + "0101' and ДатаИсхода <= '" + выбранныйГод + "1231' and " +
            //                " id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
            //                " where FlagAuto is null) " +
            //                " order by id_карточки desc ";

            string query = " select top 1 номерПП from Карточка " +
                " where ДатаПоступ <= '" + выбранныйГод + "1231' and ДатаПоступ >= '" + (Convert.ToInt32(выбранныйГод) - 1).ToString().Trim() + "1231' and FlagAuto is null " +
                  "order by id_карточки desc ";
                //" where ДатаИсхода <= '" + выбранныйГод + "1231' and  FlagAuto is null " +
                          // "where ДатаПоступ <= '" + выбранныйГод + "1231' and FlagAuto is null " +
                //" id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                //" where FlagAuto is null) " +
                //" order by id_карточки asc ";
                         

            DataRow[] dr = DataTableSql.GetDataTableRows(query);

            if (dr.Length > 0)
            {
                if (flagAutoStop == false)
                {
                    следНомерПП = 1 +(int)dr[0]["НомерПП"];
                    label1.Text = "След. номер п\\п " + (следНомерПП);
                    //textBoxНомерВходящий.Text = следНомерПП.ToString() + "/12-02-0";

                    if (this.rb02.Checked == true)
                    {
                        //textBoxНомерВходящий.Text = "12-02-";
                        textBoxНомерВходящий.Text = "02-";
                    }
                    else if (this.rb04.Checked == true)
                    {
                        textBoxНомерВходящий.Text = "04-";
                    }

                    // Сохраняем следующий номер документа.
                    следующийНомерДокумента = следНомерПП;

                    // Замена поля TextBox на MasckTextBox.
                    //префиксДокумента = "12-02-" + textBoxНомерВходящий.Text.Trim();//0";
                    префиксДокумента = textBoxНомерВходящий.Text.Trim();//0";

                    // Скроем поле ввода номера.
                    this.textBoxНомерВходящий2.Visible = false;
                }
                else
                {
                    // Сроем маску ввода номера.
                    this.textBoxНомерВходящий.Visible = false;

                    следНомерПП = 1 + (int)dr[0]["НомерПП"];
                    label1.Text = "След. номер п\\п " + (следНомерПП);
                    textBoxНомерВходящий2.Text = следНомерПП.ToString() + "/12-02-0";

                    //this.maskedTextBox1.Visible = false;
                }
            }
            else
            {
                следНомерПП = 1;
                label1.Text = "След. номер п\\п " + (следНомерПП);

                    //textBoxНомерВходящий.Text = следНомерПП.ToString() + "/12-02-0";

                    // Замена поля TextBox на MasckTextBox.
                    //textBoxНомерВходящий.Text = "";
                    textBoxНомерВходящий.Text = "12-02-";

                    // Сохраняем следующий номер документа.
                    следующийНомерДокумента = следНомерПП;

                    // Замена поля TextBox на MasckTextBox.
                    //префиксДокумента = "12-02-" + textBoxНомерВходящий.Text.Trim();
                    префиксДокумента = textBoxНомерВходящий.Text.Trim();

                    // Test.
                    this.textBoxНомерВходящий2.Visible = true;

                    this.textBoxНомерВходящий.Visible = false;
               

            }

            // Упакуем номер вновь создаваемого документа в класс.
            НомерДокумента doc = new НомерДокумента();
            doc.Номер = следующийНомерДокумента;
            doc.Префикс = префиксДокумента;

            // Сохраним номер документа в свойстве формы.
            СледующийНомерДокумента = doc;

            comboКорреспондент.DataSource = ds1.Корреспонденты;
            comboКорреспондент.DisplayMember = ds1.Корреспонденты.Columns["ОписаниеКорреспондента"].ToString();
            comboКорреспондент.ValueMember = ds1.Корреспонденты.Columns["id_корреспондента"].ToString();
            comboКорреспондент.Text = "";

            comboДокумент.DataSource = ds1.Документы;
            comboДокумент.DisplayMember = ds1.Документы.Columns["ОписаниеДокумента"].ToString();
            comboДокумент.ValueMember = ds1.Документы.Columns["id_документа"].ToString();

            comboКорреспондент.Focus();


            if (this.chBoxRepet.Checked == true)
            {
                this.txtPeriod.Enabled = true;
            }
            else
            {
                this.txtPeriod.Enabled = false;
            }
        }

        /// <summary>
        /// Конструктор формы. В качестве параметров принимает Датасет, и идентификатор строки для изменения
        /// </summary>
        /// <param name="ds">Датасет</param>
        /// <param name="idКарточки">Идентификатор строки для изменения</param>
        public FormКарточка(RegKor.DS1 ds, int idКарточки, string выбранныйГод)
        {
            //// Скроем поле редактирования.
            //this.textBoxНомерВходящий2.Visible = false;
            bool flagDsp = false;

            InitializeComponent();

            // Скроем поле редактирования.
            this.textBoxНомерВходящий2.Visible = false;

            // Загрузим список номеров подразделений комитета.
            LoadNumberDepartments();

            this.ds1 = ds;
            DataRow[] dr = ds1.Карточка.Select("id_карточки=" + idКарточки);
            DataRow[] dr2 = ds1.Выборка.Select("id_карточки=" + idКарточки);

            // Получим карточку.
            string queryCard = "select * from Карточка where id_карточки = "+ idКарточки +" ";

            DataRow rowCurrCard = DataTableSql.GetDataTable(queryCard).Rows[0];

            if (rowCurrCard["ДСП"] != DBNull.Value)
            {
                // Проверим является ли документ документом ДСП.
                if (Convert.ToBoolean(rowCurrCard["ДСП"]) == true)
                {
                    // Значит карточка является ДСП.
                    flagDsp = true;
                }
            }

            // Передадим в свойство формы id карточки.
            this.Idкарточки = idКарточки;

            строкаКарточки = (DS1.КарточкаRow)dr[0];


            // Переменная для хранения префикса документов.
            string префиксДокумента = string.Empty;

            НомерДокумента doc = new НомерДокумента();

            // Получение максимального номераПП из таблицы КарточкаИсходящая
            DataRow[] dr3 = ds.Карточка.Select("ДатаПоступ >='01.12." + выбранныйГод + "'", "НомерПП DESC");
            if (dr3.Length > 0)
            {
                следНомерПП = 1 + (int)dr3[0]["НомерПП"];
                label1.Text = "След. номер п\\п " + (следНомерПП);
            }
            else
            {
                следНомерПП = 1;
                label1.Text = "След. номер п\\п " + (следНомерПП);
            }

            comboКорреспондент.DataSource = ds1.Корреспонденты;
            comboКорреспондент.DisplayMember = ds1.Корреспонденты.Columns["ОписаниеКорреспондента"].ToString();
            comboКорреспондент.ValueMember = ds1.Корреспонденты.Columns["id_корреспондента"].ToString();
            comboКорреспондент.SelectedItem = строкаКарточки["id_корреспондента"];

            comboДокумент.DataSource = ds1.Документы;
            comboДокумент.DisplayMember = ds1.Документы.Columns["ОписаниеДокумента"].ToString();
            comboДокумент.ValueMember = ds1.Документы.Columns["id_документа"].ToString();

            comboДокумент.SelectedValue = (int)строкаКарточки["id_документа"];
            comboКорреспондент.SelectedValue = (int)строкаКарточки["id_корреспондента"];
            checkBoxВДеле.Checked = (Boolean)строкаКарточки["ВДело"];
            dateTimeДатаОтправления.Value = Convert.ToDateTime(строкаКарточки["ДатаИсхода"]);
            dateTimeДатаПоступления.Value = Convert.ToDateTime(строкаКарточки["ДатаПоступ"]);
            textBoxСодержание.Text = (string)строкаКарточки["КраткоеСодержание"];

            if (flagDsp == false)
            {
                _НомерДокументаПовторПривязка = dr2[0]["НомерВход"].ToString();
                textBoxНомерВходящий.Text = dr2[0]["НомерВход"].ToString().Split('/')[1].ToString().Trim();
            }
            else
            {
                _НомерДокументаПовторПривязка = dr2[0]["НомерВход"].ToString() + "дсп";
                this.chboxDsp.Checked = true;
                textBoxНомерВходящий.Text = dr2[0]["НомерВход"].ToString().Split('/')[1].ToString().Trim() + "дсп";
            }

            //textBoxНомерВходящий.Text = dr2[0]["НомерВход"].ToString().Split('/')[1].ToString().Trim();
           

            string[] nums = dr2[0]["НомерВход"].ToString().Split('/');

            doc.Номер = Convert.ToInt32(nums[0]); //следНомерПП;
            doc.Префикс = nums[1].ToString();

            СледующийНомерДокумента = doc;


            textBoxНомерИсходящий.Text = (string)строкаКарточки["НомерИсход"];
            textBoxРезолюция.Text = (string)строкаКарточки["Резолюция"];

            // Если вдруг результат выполнения сталравен null по покане понятной причиине.
            if (строкаКарточки["РезультатВыполнения"] == DBNull.Value)
            {
                строкаКарточки["РезультатВыполнения"] = "";
            }

            if ((string)строкаКарточки["РезультатВыполнения"] != "")
            {
                textBoxРезультатВыполнения.Enabled = true;
                textBoxРезультатВыполнения.Text = (string)строкаКарточки["РезультатВыполнения"];
            }
            checkBoxКонтроль.Checked = (Boolean)строкаКарточки["НаКонтроле"];
            if (строкаКарточки["СрокВыполнения"] != System.DBNull.Value)
            {
                dateTimeКонтроль.Value = Convert.ToDateTime(строкаКарточки["СрокВыполнения"]);
            }
            if (checkBoxКонтроль.Checked)
            {
                dateTimeКонтроль.Enabled = true;
            }

            chBoxRepet.Checked = (Boolean)строкаКарточки["FlagCardRepeet"];

            if (this.chBoxRepet.Checked == true)
            {
                this.txtPeriod.Enabled = true;
            }
            else
            {
                this.txtPeriod.Enabled = false;
            }

            if (ДокументооборотConfig.ВключитьДокументооборот() == true)
            {
                string queryDoc = "select FileDate,FileDateTitlePage from КарточкиДокументы " +
                                  "where id_карточки = " + idКарточки + " ";
                GetDataTable tab = new GetDataTable(queryDoc);
                DataTable tabRow = tab.DataTable();

                // Обнулим переменную для хранения путь к карточке документа VipNet.
                pathFileServerTitlePage = null;

                //if (tabRow.Rows[0]["FileDate"].ToString() != "" || tabRow.Rows[0]["FileDate"] != null)
                if (tabRow.Rows.Count > 0)
                {
                    if (tabRow.Rows[0]["FileDate"] != DBNull.Value)
                    {
                        this.linkLabel1.Text = "Просмотреть файл документа";

                        ////this.lblFile.Text = "Файл документа - " + tabRow.Rows[0]["NameFileDocument"].ToString();
                        //this.linkLabel1.Text = "Файл документа - " + tabRow.Rows[0]["NameFileDocument"].ToString().Split('_')[0].Trim();

                        //if (tabRow.Rows[0]["GuidName"].ToString().Trim() != "")
                        //{
                        //    pathFileServer = tabRow.Rows[0]["NameFileDocument"].ToString().Trim();// +"_" + tabRow.Rows[0]["GuidName"].ToString().Trim();
                        //}
                        //else
                        //{
                        //    pathFileServer = tabRow.Rows[0]["NameFileDocument"].ToString().Trim();
                        //}


                        //ИмяДокумента = null;
                        //ИмяДокумента = pathFileServer;

                    }
                    else
                    {
                        //this.lblFile.Text = "";

                        ИмяДокумента = null;
                    }
                }
                else
                {
                    ИмяДокумента = null;
                }



                //if (tabRow.Rows[0]["MD5"].ToString() == "md5")
                //{
                //    this.chcDop.Checked = true;
                //}
                //else
                //{
                //    this.chcDop.Checked = false;
                //}

                //// Получим путь к серверу.
                //string patchServerQuery = "select PatchServer from СерверПуть";

                //GetDataTable tabServer = new GetDataTable(patchServerQuery);
                //DataTable tabServFile = tabServer.DataTable();

                //patchServerSave = tabServFile.Rows[0]["PatchServer"].ToString().Trim();

                //// Установим флаг FlagUpdateDocument в true, в связи с тем что форма открыта для изменения.
                //FlagUpdateDocument = true;

                // Отобразим ссылку на титульный лист.
                //if (tabRow.Rows[0]["FileDateTitlePage"].ToString() != "" || tabRow.Rows[0]["FileDateTitlePage"].ToString() != null)
                if (tabRow.Rows.Count > 0)
                {
                    if (tabRow.Rows[0]["FileDateTitlePage"] != DBNull.Value)
                    {
                        this.linkLabel2.Text = "Просмотреть файл титульного листа";

                        //this.linkLabel2.Text = "Просмотреть aайл титульного листа - " + tabRow.Rows[0]["NameFileDocumentVipNetEmailTitlePage"].ToString().Split('_')[0].Trim();

                        //if (tabRow.Rows[0]["GuidName"].ToString().Trim() != "")
                        //{
                        //    pathFileServerTitlePage = tabRow.Rows[0]["NameFileDocumentVipNetEmailTitlePage"].ToString().Trim();// +"_" + tabRow.Rows[0]["GuidName"].ToString().Trim();
                        //}
                        //else
                        //{
                        //    pathFileServerTitlePage = tabRow.Rows[0]["NameFileDocumentVipNetEmailTitlePage"].ToString().Trim();
                        //}



                        //pathFileServerTitlePage = pathFileServerTitlePage;
                    }
                }
            }
        


        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormКарточка));
            this.labelКорреспID = new System.Windows.Forms.Label();
            this.labelРезультатВыполнения = new System.Windows.Forms.Label();
            this.labelРезолюция = new System.Windows.Forms.Label();
            this.labelСодержание = new System.Windows.Forms.Label();
            this.labelДокумент = new System.Windows.Forms.Label();
            this.labelКорреспондент = new System.Windows.Forms.Label();
            this.checkBoxВДеле = new System.Windows.Forms.CheckBox();
            this.buttonОтмена = new System.Windows.Forms.Button();
            this.buttonСохранить = new System.Windows.Forms.Button();
            this.groupBoxОтправлено = new System.Windows.Forms.GroupBox();
            this.labelНомерИсходящий = new System.Windows.Forms.Label();
            this.labelДатаОтправления = new System.Windows.Forms.Label();
            this.dateTimeДатаОтправления = new System.Windows.Forms.DateTimePicker();
            this.textBoxНомерИсходящий = new System.Windows.Forms.TextBox();
            this.textBoxРезолюция = new System.Windows.Forms.TextBox();
            this.textBoxСодержание = new System.Windows.Forms.TextBox();
            this.dateTimeКонтроль = new System.Windows.Forms.DateTimePicker();
            this.checkBoxКонтроль = new System.Windows.Forms.CheckBox();
            this.comboДокумент = new System.Windows.Forms.ComboBox();
            this.comboКорреспондент = new System.Windows.Forms.ComboBox();
            this.groupBoxПоступило = new System.Windows.Forms.GroupBox();
            this.rb04 = new System.Windows.Forms.RadioButton();
            this.rb02 = new System.Windows.Forms.RadioButton();
            this.chboxDsp = new System.Windows.Forms.CheckBox();
            this.textBoxНомерВходящий2 = new System.Windows.Forms.TextBox();
            this.textBoxНомерВходящий = new System.Windows.Forms.MaskedTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.labelНомерВходящий = new System.Windows.Forms.Label();
            this.labelДатаПоступления = new System.Windows.Forms.Label();
            this.dateTimeДатаПоступления = new System.Windows.Forms.DateTimePicker();
            this.ds1 = new RegKor.DS1();
            this.textBoxРезультатВыполнения = new System.Windows.Forms.TextBox();
            this.panelKontrol = new System.Windows.Forms.Panel();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.buttonДобавитьПолучателей = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnTie = new System.Windows.Forms.Button();
            this.btnElementPS = new System.Windows.Forms.Button();
            this.chBoxRepet = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPeriod = new System.Windows.Forms.TextBox();
            this.chekDocServer = new System.Windows.Forms.CheckBox();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.chcDop = new System.Windows.Forms.CheckBox();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.btnLastNumber = new System.Windows.Forms.Button();
            this.groupBoxОтправлено.SuspendLayout();
            this.groupBoxПоступило.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ds1)).BeginInit();
            this.panelKontrol.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelКорреспID
            // 
            this.labelКорреспID.Location = new System.Drawing.Point(516, 2);
            this.labelКорреспID.Name = "labelКорреспID";
            this.labelКорреспID.Size = new System.Drawing.Size(26, 24);
            this.labelКорреспID.TabIndex = 0;
            // 
            // labelРезультатВыполнения
            // 
            this.labelРезультатВыполнения.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelРезультатВыполнения.Location = new System.Drawing.Point(4, 404);
            this.labelРезультатВыполнения.Name = "labelРезультатВыполнения";
            this.labelРезультатВыполнения.Size = new System.Drawing.Size(172, 14);
            this.labelРезультатВыполнения.TabIndex = 0;
            this.labelРезультатВыполнения.Text = "Результат выполнения";
            this.labelРезультатВыполнения.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // labelРезолюция
            // 
            this.labelРезолюция.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelРезолюция.Location = new System.Drawing.Point(4, 344);
            this.labelРезолюция.Name = "labelРезолюция";
            this.labelРезолюция.Size = new System.Drawing.Size(172, 13);
            this.labelРезолюция.TabIndex = 0;
            this.labelРезолюция.Text = "Резолюция";
            this.labelРезолюция.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // labelСодержание
            // 
            this.labelСодержание.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelСодержание.Location = new System.Drawing.Point(4, 284);
            this.labelСодержание.Name = "labelСодержание";
            this.labelСодержание.Size = new System.Drawing.Size(172, 14);
            this.labelСодержание.TabIndex = 0;
            this.labelСодержание.Text = "Краткое содержание";
            this.labelСодержание.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // labelДокумент
            // 
            this.labelДокумент.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelДокумент.Location = new System.Drawing.Point(8, 32);
            this.labelДокумент.Name = "labelДокумент";
            this.labelДокумент.Size = new System.Drawing.Size(108, 22);
            this.labelДокумент.TabIndex = 0;
            this.labelДокумент.Text = "Вид документа";
            this.labelДокумент.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelКорреспондент
            // 
            this.labelКорреспондент.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelКорреспондент.Location = new System.Drawing.Point(10, 3);
            this.labelКорреспондент.Name = "labelКорреспондент";
            this.labelКорреспондент.Size = new System.Drawing.Size(108, 24);
            this.labelКорреспондент.TabIndex = 0;
            this.labelКорреспондент.Text = "Корреспондент";
            this.labelКорреспондент.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // checkBoxВДеле
            // 
            this.checkBoxВДеле.Checked = true;
            this.checkBoxВДеле.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxВДеле.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBoxВДеле.Location = new System.Drawing.Point(354, 215);
            this.checkBoxВДеле.Name = "checkBoxВДеле";
            this.checkBoxВДеле.Size = new System.Drawing.Size(118, 24);
            this.checkBoxВДеле.TabIndex = 12;
            this.checkBoxВДеле.Text = "В деле";
            this.checkBoxВДеле.CheckedChanged += new System.EventHandler(this.checkBoxВДеле_CheckedChanged);
            // 
            // buttonОтмена
            // 
            this.buttonОтмена.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonОтмена.Dock = System.Windows.Forms.DockStyle.Right;
            this.buttonОтмена.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonОтмена.Location = new System.Drawing.Point(370, 0);
            this.buttonОтмена.Name = "buttonОтмена";
            this.buttonОтмена.Size = new System.Drawing.Size(178, 28);
            this.buttonОтмена.TabIndex = 17;
            this.buttonОтмена.Text = "Отмена";
            this.toolTip1.SetToolTip(this.buttonОтмена, "Закрыть окно без сохранения изменений");
            this.buttonОтмена.Click += new System.EventHandler(this.buttonОтмена_Click);
            // 
            // buttonСохранить
            // 
            this.buttonСохранить.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonСохранить.Dock = System.Windows.Forms.DockStyle.Left;
            this.buttonСохранить.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonСохранить.Location = new System.Drawing.Point(0, 0);
            this.buttonСохранить.Name = "buttonСохранить";
            this.buttonСохранить.Size = new System.Drawing.Size(178, 28);
            this.buttonСохранить.TabIndex = 16;
            this.buttonСохранить.Text = "Сохранить";
            this.toolTip1.SetToolTip(this.buttonСохранить, "Сохранить изменения и закрыть окно");
            this.buttonСохранить.Click += new System.EventHandler(this.buttonСохранить_Click);
            // 
            // groupBoxОтправлено
            // 
            this.groupBoxОтправлено.Controls.Add(this.labelНомерИсходящий);
            this.groupBoxОтправлено.Controls.Add(this.labelДатаОтправления);
            this.groupBoxОтправлено.Controls.Add(this.dateTimeДатаОтправления);
            this.groupBoxОтправлено.Controls.Add(this.textBoxНомерИсходящий);
            this.groupBoxОтправлено.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBoxОтправлено.Location = new System.Drawing.Point(8, 56);
            this.groupBoxОтправлено.Name = "groupBoxОтправлено";
            this.groupBoxОтправлено.Size = new System.Drawing.Size(258, 103);
            this.groupBoxОтправлено.TabIndex = 3;
            this.groupBoxОтправлено.TabStop = false;
            this.groupBoxОтправлено.Text = "Отправлено";
            // 
            // labelНомерИсходящий
            // 
            this.labelНомерИсходящий.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelНомерИсходящий.Location = new System.Drawing.Point(10, 56);
            this.labelНомерИсходящий.Name = "labelНомерИсходящий";
            this.labelНомерИсходящий.Size = new System.Drawing.Size(94, 20);
            this.labelНомерИсходящий.TabIndex = 0;
            this.labelНомерИсходящий.Text = "Исходящий №";
            this.labelНомерИсходящий.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelДатаОтправления
            // 
            this.labelДатаОтправления.Location = new System.Drawing.Point(10, 22);
            this.labelДатаОтправления.Name = "labelДатаОтправления";
            this.labelДатаОтправления.Size = new System.Drawing.Size(92, 20);
            this.labelДатаОтправления.TabIndex = 0;
            this.labelДатаОтправления.Text = "Дата";
            this.labelДатаОтправления.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dateTimeДатаОтправления
            // 
            this.dateTimeДатаОтправления.CalendarTrailingForeColor = System.Drawing.SystemColors.Control;
            this.dateTimeДатаОтправления.Location = new System.Drawing.Point(106, 22);
            this.dateTimeДатаОтправления.Name = "dateTimeДатаОтправления";
            this.dateTimeДатаОтправления.Size = new System.Drawing.Size(144, 22);
            this.dateTimeДатаОтправления.TabIndex = 4;
            // 
            // textBoxНомерИсходящий
            // 
            this.textBoxНомерИсходящий.BackColor = System.Drawing.SystemColors.HighlightText;
            this.textBoxНомерИсходящий.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxНомерИсходящий.Location = new System.Drawing.Point(106, 56);
            this.textBoxНомерИсходящий.MaxLength = 20;
            this.textBoxНомерИсходящий.Name = "textBoxНомерИсходящий";
            this.textBoxНомерИсходящий.Size = new System.Drawing.Size(144, 22);
            this.textBoxНомерИсходящий.TabIndex = 5;
            this.textBoxНомерИсходящий.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBoxРезолюция
            // 
            this.textBoxРезолюция.BackColor = System.Drawing.SystemColors.HighlightText;
            this.textBoxРезолюция.Enabled = false;
            this.textBoxРезолюция.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxРезолюция.Location = new System.Drawing.Point(2, 359);
            this.textBoxРезолюция.MaxLength = 250;
            this.textBoxРезолюция.Multiline = true;
            this.textBoxРезолюция.Name = "textBoxРезолюция";
            this.textBoxРезолюция.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxРезолюция.Size = new System.Drawing.Size(512, 42);
            this.textBoxРезолюция.TabIndex = 0;
            this.textBoxРезолюция.TabStop = false;
            // 
            // textBoxСодержание
            // 
            this.textBoxСодержание.BackColor = System.Drawing.SystemColors.HighlightText;
            this.textBoxСодержание.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxСодержание.Location = new System.Drawing.Point(2, 300);
            this.textBoxСодержание.MaxLength = 250;
            this.textBoxСодержание.Multiline = true;
            this.textBoxСодержание.Name = "textBoxСодержание";
            this.textBoxСодержание.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxСодержание.Size = new System.Drawing.Size(540, 42);
            this.textBoxСодержание.TabIndex = 13;
            // 
            // dateTimeКонтроль
            // 
            this.dateTimeКонтроль.Enabled = false;
            this.dateTimeКонтроль.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dateTimeКонтроль.Location = new System.Drawing.Point(136, 4);
            this.dateTimeКонтроль.Name = "dateTimeКонтроль";
            this.dateTimeКонтроль.Size = new System.Drawing.Size(144, 22);
            this.dateTimeКонтроль.TabIndex = 11;
            // 
            // checkBoxКонтроль
            // 
            this.checkBoxКонтроль.BackColor = System.Drawing.SystemColors.Control;
            this.checkBoxКонтроль.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBoxКонтроль.Location = new System.Drawing.Point(4, 4);
            this.checkBoxКонтроль.Name = "checkBoxКонтроль";
            this.checkBoxКонтроль.Size = new System.Drawing.Size(126, 24);
            this.checkBoxКонтроль.TabIndex = 10;
            this.checkBoxКонтроль.Text = "На контроле до";
            this.checkBoxКонтроль.UseVisualStyleBackColor = false;
            this.checkBoxКонтроль.CheckedChanged += new System.EventHandler(this.checkBoxКонтроль_CheckedChanged);
            // 
            // comboДокумент
            // 
            this.comboДокумент.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboДокумент.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboДокумент.DisplayMember = "id_документа";
            this.comboДокумент.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboДокумент.Location = new System.Drawing.Point(124, 30);
            this.comboДокумент.MaxDropDownItems = 40;
            this.comboДокумент.Name = "comboДокумент";
            this.comboДокумент.Size = new System.Drawing.Size(372, 24);
            this.comboДокумент.TabIndex = 2;
            this.comboДокумент.ValueMember = "id_документа";
            // 
            // comboКорреспондент
            // 
            this.comboКорреспондент.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboКорреспондент.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboКорреспондент.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboКорреспондент.Location = new System.Drawing.Point(124, 1);
            this.comboКорреспондент.MaxDropDownItems = 40;
            this.comboКорреспондент.Name = "comboКорреспондент";
            this.comboКорреспондент.Size = new System.Drawing.Size(372, 24);
            this.comboКорреспондент.TabIndex = 1;
            // 
            // groupBoxПоступило
            // 
            this.groupBoxПоступило.Controls.Add(this.rb04);
            this.groupBoxПоступило.Controls.Add(this.rb02);
            this.groupBoxПоступило.Controls.Add(this.chboxDsp);
            this.groupBoxПоступило.Controls.Add(this.textBoxНомерВходящий2);
            this.groupBoxПоступило.Controls.Add(this.textBoxНомерВходящий);
            this.groupBoxПоступило.Controls.Add(this.label1);
            this.groupBoxПоступило.Controls.Add(this.labelНомерВходящий);
            this.groupBoxПоступило.Controls.Add(this.labelДатаПоступления);
            this.groupBoxПоступило.Controls.Add(this.dateTimeДатаПоступления);
            this.groupBoxПоступило.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBoxПоступило.Location = new System.Drawing.Point(284, 60);
            this.groupBoxПоступило.Name = "groupBoxПоступило";
            this.groupBoxПоступило.Size = new System.Drawing.Size(258, 145);
            this.groupBoxПоступило.TabIndex = 6;
            this.groupBoxПоступило.TabStop = false;
            this.groupBoxПоступило.Text = "Поступило";
            // 
            // rb04
            // 
            this.rb04.AutoSize = true;
            this.rb04.Location = new System.Drawing.Point(112, 49);
            this.rb04.Name = "rb04";
            this.rb04.Size = new System.Drawing.Size(40, 20);
            this.rb04.TabIndex = 13;
            this.rb04.Text = "04";
            this.rb04.UseVisualStyleBackColor = true;
            this.rb04.Visible = false;
            this.rb04.CheckedChanged += new System.EventHandler(this.rb04_CheckedChanged);
            // 
            // rb02
            // 
            this.rb02.AutoSize = true;
            this.rb02.Checked = true;
            this.rb02.Location = new System.Drawing.Point(15, 49);
            this.rb02.Name = "rb02";
            this.rb02.Size = new System.Drawing.Size(40, 20);
            this.rb02.TabIndex = 13;
            this.rb02.TabStop = true;
            this.rb02.Text = "02";
            this.rb02.UseVisualStyleBackColor = true;
            this.rb02.CheckedChanged += new System.EventHandler(this.rb02_CheckedChanged);
            // 
            // chboxDsp
            // 
            this.chboxDsp.AutoSize = true;
            this.chboxDsp.Location = new System.Drawing.Point(15, 120);
            this.chboxDsp.Name = "chboxDsp";
            this.chboxDsp.Size = new System.Drawing.Size(55, 20);
            this.chboxDsp.TabIndex = 12;
            this.chboxDsp.Text = "ДСП";
            this.chboxDsp.UseVisualStyleBackColor = true;
            this.chboxDsp.CheckedChanged += new System.EventHandler(this.chboxDsp_CheckedChanged);
            // 
            // textBoxНомерВходящий2
            // 
            this.textBoxНомерВходящий2.Location = new System.Drawing.Point(112, 75);
            this.textBoxНомерВходящий2.Name = "textBoxНомерВходящий2";
            this.textBoxНомерВходящий2.Size = new System.Drawing.Size(133, 22);
            this.textBoxНомерВходящий2.TabIndex = 11;
            this.textBoxНомерВходящий2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBoxНомерВходящий
            // 
            this.textBoxНомерВходящий.Location = new System.Drawing.Point(107, 75);
            this.textBoxНомерВходящий.Mask = "00-00-00";
            this.textBoxНомерВходящий.Name = "textBoxНомерВходящий";
            this.textBoxНомерВходящий.Size = new System.Drawing.Size(123, 22);
            this.textBoxНомерВходящий.TabIndex = 10;
            this.textBoxНомерВходящий.TextChanged += new System.EventHandler(this.textBoxНомерВходящий_TextChanged);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(6, 102);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(239, 18);
            this.label1.TabIndex = 9;
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // labelНомерВходящий
            // 
            this.labelНомерВходящий.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelНомерВходящий.Location = new System.Drawing.Point(12, 76);
            this.labelНомерВходящий.Name = "labelНомерВходящий";
            this.labelНомерВходящий.Size = new System.Drawing.Size(88, 20);
            this.labelНомерВходящий.TabIndex = 0;
            this.labelНомерВходящий.Text = "Входящий №";
            this.labelНомерВходящий.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelДатаПоступления
            // 
            this.labelДатаПоступления.Location = new System.Drawing.Point(12, 22);
            this.labelДатаПоступления.Name = "labelДатаПоступления";
            this.labelДатаПоступления.Size = new System.Drawing.Size(88, 20);
            this.labelДатаПоступления.TabIndex = 0;
            this.labelДатаПоступления.Text = "Дата";
            this.labelДатаПоступления.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dateTimeДатаПоступления
            // 
            this.dateTimeДатаПоступления.CalendarTrailingForeColor = System.Drawing.SystemColors.Control;
            this.dateTimeДатаПоступления.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.ds1, "Выборка.ДатаПоступ", true));
            this.dateTimeДатаПоступления.Location = new System.Drawing.Point(102, 22);
            this.dateTimeДатаПоступления.Name = "dateTimeДатаПоступления";
            this.dateTimeДатаПоступления.Size = new System.Drawing.Size(144, 22);
            this.dateTimeДатаПоступления.TabIndex = 7;
            // 
            // ds1
            // 
            this.ds1.DataSetName = "DS";
            this.ds1.Locale = new System.Globalization.CultureInfo("ru-RU");
            this.ds1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // textBoxРезультатВыполнения
            // 
            this.textBoxРезультатВыполнения.Enabled = false;
            this.textBoxРезультатВыполнения.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxРезультатВыполнения.Location = new System.Drawing.Point(2, 419);
            this.textBoxРезультатВыполнения.MaxLength = 250;
            this.textBoxРезультатВыполнения.Multiline = true;
            this.textBoxРезультатВыполнения.Name = "textBoxРезультатВыполнения";
            this.textBoxРезультатВыполнения.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxРезультатВыполнения.Size = new System.Drawing.Size(512, 42);
            this.textBoxРезультатВыполнения.TabIndex = 15;
            // 
            // panelKontrol
            // 
            this.panelKontrol.BackColor = System.Drawing.SystemColors.Control;
            this.panelKontrol.Controls.Add(this.dateTimeКонтроль);
            this.panelKontrol.Controls.Add(this.checkBoxКонтроль);
            this.panelKontrol.Location = new System.Drawing.Point(6, 211);
            this.panelKontrol.Name = "panelKontrol";
            this.panelKontrol.Size = new System.Drawing.Size(290, 30);
            this.panelKontrol.TabIndex = 9;
            // 
            // buttonДобавитьПолучателей
            // 
            this.buttonДобавитьПолучателей.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonДобавитьПолучателей.Image = global::RegKor.Properties.Resources.add;
            this.buttonДобавитьПолучателей.Location = new System.Drawing.Point(520, 367);
            this.buttonДобавитьПолучателей.Name = "buttonДобавитьПолучателей";
            this.buttonДобавитьПолучателей.Size = new System.Drawing.Size(22, 22);
            this.buttonДобавитьПолучателей.TabIndex = 14;
            this.toolTip1.SetToolTip(this.buttonДобавитьПолучателей, "Добавить лиц, которым направлен документ");
            this.buttonДобавитьПолучателей.Click += new System.EventHandler(this.buttonДобавитьПолучателей_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.btnTie);
            this.panel1.Controls.Add(this.buttonСохранить);
            this.panel1.Controls.Add(this.buttonОтмена);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 513);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(552, 32);
            this.panel1.TabIndex = 18;
            // 
            // btnTie
            // 
            this.btnTie.Location = new System.Drawing.Point(184, 0);
            this.btnTie.Name = "btnTie";
            this.btnTie.Size = new System.Drawing.Size(178, 28);
            this.btnTie.TabIndex = 18;
            this.btnTie.Text = "Повторно привязать";
            this.btnTie.UseVisualStyleBackColor = true;
            this.btnTie.Visible = false;
            this.btnTie.Click += new System.EventHandler(this.btnTie_Click);
            // 
            // btnElementPS
            // 
            this.btnElementPS.Location = new System.Drawing.Point(318, 275);
            this.btnElementPS.Name = "btnElementPS";
            this.btnElementPS.Size = new System.Drawing.Size(211, 23);
            this.btnElementPS.TabIndex = 19;
            this.btnElementPS.Text = "Состав персональных данных";
            this.btnElementPS.UseVisualStyleBackColor = true;
            this.btnElementPS.Visible = false;
            this.btnElementPS.Click += new System.EventHandler(this.button1_Click);
            // 
            // chBoxRepet
            // 
            this.chBoxRepet.AutoSize = true;
            this.chBoxRepet.Location = new System.Drawing.Point(354, 257);
            this.chBoxRepet.Name = "chBoxRepet";
            this.chBoxRepet.Size = new System.Drawing.Size(152, 17);
            this.chBoxRepet.TabIndex = 20;
            this.chBoxRepet.Text = "Ответ с периодичностью";
            this.chBoxRepet.UseVisualStyleBackColor = true;
            this.chBoxRepet.CheckedChanged += new System.EventHandler(this.chBoxRepet_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 257);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(112, 13);
            this.label2.TabIndex = 21;
            this.label2.Text = "Периодичность дней";
            // 
            // txtPeriod
            // 
            this.txtPeriod.Enabled = false;
            this.txtPeriod.Location = new System.Drawing.Point(142, 254);
            this.txtPeriod.Name = "txtPeriod";
            this.txtPeriod.Size = new System.Drawing.Size(100, 20);
            this.txtPeriod.TabIndex = 22;
            this.txtPeriod.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtPeriod_KeyPress);
            // 
            // chekDocServer
            // 
            this.chekDocServer.AutoSize = true;
            this.chekDocServer.Enabled = false;
            this.chekDocServer.Location = new System.Drawing.Point(10, 468);
            this.chekDocServer.Name = "chekDocServer";
            this.chekDocServer.Size = new System.Drawing.Size(171, 17);
            this.chekDocServer.TabIndex = 23;
            this.chekDocServer.Text = "Сохранить копию документа";
            this.chekDocServer.UseVisualStyleBackColor = true;
            this.chekDocServer.Visible = false;
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(206, 468);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(0, 13);
            this.linkLabel1.TabIndex = 24;
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // chcDop
            // 
            this.chcDop.AutoSize = true;
            this.chcDop.Location = new System.Drawing.Point(438, 219);
            this.chcDop.Name = "chcDop";
            this.chcDop.Size = new System.Drawing.Size(76, 17);
            this.chcDop.TabIndex = 25;
            this.chcDop.Text = "Добавить";
            this.chcDop.UseVisualStyleBackColor = true;
            this.chcDop.Visible = false;
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Location = new System.Drawing.Point(206, 492);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(0, 13);
            this.linkLabel2.TabIndex = 27;
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // btnLastNumber
            // 
            this.btnLastNumber.Location = new System.Drawing.Point(12, 176);
            this.btnLastNumber.Name = "btnLastNumber";
            this.btnLastNumber.Size = new System.Drawing.Size(230, 23);
            this.btnLastNumber.TabIndex = 29;
            this.btnLastNumber.Text = "Основание передачи";
            this.btnLastNumber.UseVisualStyleBackColor = true;
            this.btnLastNumber.Click += new System.EventHandler(this.btnLastNumber_Click);
            // 
            // FormКарточка
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.buttonОтмена;
            this.ClientSize = new System.Drawing.Size(552, 545);
            this.ControlBox = false;
            this.Controls.Add(this.btnLastNumber);
            this.Controls.Add(this.linkLabel2);
            this.Controls.Add(this.chcDop);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.chekDocServer);
            this.Controls.Add(this.txtPeriod);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.chBoxRepet);
            this.Controls.Add(this.btnElementPS);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.buttonДобавитьПолучателей);
            this.Controls.Add(this.panelKontrol);
            this.Controls.Add(this.textBoxРезультатВыполнения);
            this.Controls.Add(this.labelКорреспID);
            this.Controls.Add(this.labelРезультатВыполнения);
            this.Controls.Add(this.labelРезолюция);
            this.Controls.Add(this.labelСодержание);
            this.Controls.Add(this.labelДокумент);
            this.Controls.Add(this.labelКорреспондент);
            this.Controls.Add(this.checkBoxВДеле);
            this.Controls.Add(this.groupBoxОтправлено);
            this.Controls.Add(this.textBoxРезолюция);
            this.Controls.Add(this.textBoxСодержание);
            this.Controls.Add(this.comboДокумент);
            this.Controls.Add(this.comboКорреспондент);
            this.Controls.Add(this.groupBoxПоступило);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(554, 436);
            this.Name = "FormКарточка";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Карточка документа";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormКарточка_FormClosed);
            this.Shown += new System.EventHandler(this.FormКарточка_Shown);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormКарточка_FormClosing);
            this.Load += new System.EventHandler(this.FormКарточка_Load);
            this.groupBoxОтправлено.ResumeLayout(false);
            this.groupBoxОтправлено.PerformLayout();
            this.groupBoxПоступило.ResumeLayout(false);
            this.groupBoxПоступило.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ds1)).EndInit();
            this.panelKontrol.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion
        private void checkBoxКонтроль_CheckedChanged(object sender, System.EventArgs e)
        {
            if (checkBoxКонтроль.Checked)
            {
                dateTimeКонтроль.Enabled = true;
            }
            else
            {
                dateTimeКонтроль.Enabled = false;
            }
        }

        private void checkBoxВДеле_CheckedChanged(object sender, System.EventArgs e)
        {
            if (checkBoxВДеле.Checked)
            {
                textBoxРезультатВыполнения.Enabled = true;
                textBoxРезультатВыполнения.Text = старыйТекстРезультата;
                panelKontrol.Enabled = false;
            }
            else
            {
                textBoxРезультатВыполнения.Enabled = false;
                старыйТекстРезультата = textBoxРезультатВыполнения.Text;
                textBoxРезультатВыполнения.Text = "";
                panelKontrol.Enabled = true;
            }
        }

        private void buttonДобавитьПолучателей_Click(object sender, System.EventArgs e)
        {
            FormРезолюция form = new FormРезолюция();
            DialogResult result = form.ShowDialog(this);

            if (result == DialogResult.OK)
            {
                textBoxРезолюция.Text = form.строкаРезолюции;

                // Получим список начальников отделов и управлений которым отписан документ.
                this.ListPerson = form.ListPerson;

                string sTest = "";
            }
        }

        private void FormКарточка_Shown(object sender, EventArgs e)
        {
            comboКорреспондент.Focus();
        }


        /// <summary>
        /// Открывает справочник документов, 
        /// если в комбике есть текст то справочник открывается на добавление, 
        /// а если комбик документов пустой, то и справочник открывается на просмотр, 
        /// а дальше как юзер поступит
        /// </summary>
        private void СправочникДокументов()
        {
            FormДокументы form;
            string новыйДокумент = comboДокумент.Text;

            if (новыйДокумент != "")
            {
                form = new FormДокументы(новыйДокумент);
                form.Text = "Справочник \"Документы\"";
            }
            else
            {
                form = new FormДокументы();
                form.Text = "Справочник \"Документы\"";
            }

            form.ShowDialog(this);
            ds1.Документы.Clear();
            DS1TableAdapters.ДокументыTableAdapter adapter = new RegKor.DS1TableAdapters.ДокументыTableAdapter();
            adapter.Fill(ds1.Документы);
            comboДокумент.DataSource = null;
            comboДокумент.DisplayMember = "";
            comboДокумент.ValueMember = "";
            comboДокумент.DataSource = ds1.Документы;
            comboДокумент.DisplayMember = ds1.Документы.Columns["ОписаниеДокумента"].ToString();
            comboДокумент.ValueMember = ds1.Документы.Columns["id_документа"].ToString();
            comboДокумент.Text = новыйДокумент;
            if (comboДокумент.Text != новыйДокумент)
            {
                comboДокумент.Text = "";
                comboДокумент.SelectedText = новыйДокумент;
            }
        }

        /// <summary>
        /// Открывает справочник корреспондентов, 
        /// если в комбике есть текст то справочник открывается на добавление, 
        /// а если комбик корреспондентов пустой, 
        /// то и справочник открывается на просмотр, 
        /// а дальше как юзер поступит
        /// </summary>
        private void СправочникКорреспондентов()
        {
            FormКорреспонденты form;
            string новыйКорреспондент = comboКорреспондент.Text;

            if (новыйКорреспондент != "")
            {
                form = new FormКорреспонденты(новыйКорреспондент);
                form.Text = "Справочник \"Корреспонденты\"";
            }
            else
            {
                form = new FormКорреспонденты();
                form.Text = "Справочник \"Корреспонденты\"";
            }

            form.ShowDialog(this);
            ds1.Корреспонденты.Clear();
            DS1TableAdapters.КорреспондентыTableAdapter adapter = new RegKor.DS1TableAdapters.КорреспондентыTableAdapter();
            adapter.Fill(ds1.Корреспонденты);
            comboКорреспондент.DataSource = null;
            comboКорреспондент.DisplayMember = "";
            comboКорреспондент.ValueMember = "";
            comboКорреспондент.DataSource = ds1.Корреспонденты;
            comboКорреспондент.DisplayMember = ds1.Корреспонденты.Columns["ОписаниеКорреспондента"].ToString();
            comboКорреспондент.ValueMember = ds1.Корреспонденты.Columns["id_корреспондента"].ToString();
            comboКорреспондент.Text = новыйКорреспондент;
            if (comboКорреспондент.Text != новыйКорреспондент)
            {
                comboКорреспондент.Text = "";
                comboКорреспондент.SelectedText = новыйКорреспондент;
            }
        }

        private void textBoxНомерВходящий_Enter(object sender, EventArgs e)
        {
            textBoxНомерВходящий.Select(textBoxНомерВходящий.Text.Length, 0);
        }


        private void buttonОтмена_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        private void buttonСохранить_Click(object sender, System.EventArgs e)
        {
            #region Документ
            // Устанавливаем id документа:
            DataRow[] rows = ds1.Документы.Select("ОписаниеДокумента='" + comboДокумент.Text.Trim() + "'");
            if (rows.Length > 0)
            {
                строкаКарточки["id_документа"] = (int)comboДокумент.SelectedValue;
            }
            else if (comboДокумент.Text != "")
            {
                DialogResult result = MessageBox.Show(this,
                    "Вы указали документ, который не зарегистрирован в справочнике \"Документы\". Будем добавлять его в справочник или нет?",
                    "Неизвестный тип документа",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);
                if (result == DialogResult.No)
                {// Если юзер сказал нет, прерываем сохранение и выходим из процедуры
                    this.DialogResult = DialogResult.None;
                    return;
                }
                if (result == DialogResult.Yes)
                {// Если юзер сказал да, открываем справочник, прерываем сохранение и выходим из процедуры
                    СправочникДокументов();
                    this.DialogResult = DialogResult.None;
                    return;
                }
            }
            else if (comboДокумент.Text.Trim() == "")
            {
                MessageBox.Show(this,
                "Вы не указали тип документа",
                "Тип документа",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.None;
                return;
            }
            #endregion

            #region Корреспондент
            // Устанавливаем id корреспондента:
            DataRow[] rows2 = ds1.Корреспонденты.Select("ОписаниеКорреспондента='" + comboКорреспондент.Text.Trim() + "'");
            if (rows2.Length > 0)
            {
                строкаКарточки["id_корреспондента"] = (int)comboКорреспондент.SelectedValue;
            }
            else if (comboКорреспондент.Text != "")
            {
                DialogResult result = MessageBox.Show(this,
                    "Вы указали корреспондента, который не зарегистрирован в справочнике \"Корреспонденты\". Будем добавлять его в справочник или нет?",
                    "Неизвестный корреспондент",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);
                if (result == DialogResult.No)
                {// Если юзер сказал нет, прерываем сохранение и выходим из процедуры
                    this.DialogResult = DialogResult.None;
                    return;
                }
                if (result == DialogResult.Yes)
                {// Если юзер сказал да, открываем справочник, прерываем сохранение и выходим из процедуры
                    СправочникКорреспондентов();
                    this.DialogResult = DialogResult.None;
                    return;
                }
            }
            else if (comboКорреспондент.Text.Trim() == "")
            {
                MessageBox.Show(this,
                "Вы не указали корреспондента",
                "Корреспондент",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.None;
                return;
            }
            #endregion

            #region В Дело, На контроле

            строкаКарточки["ВДело"] = (Boolean)checkBoxВДеле.Checked;
            строкаКарточки["СрокВыполнения"] = dateTimeКонтроль.Value.ToShortDateString();
            строкаКарточки["НаКонтроле"] = checkBoxКонтроль.Checked;

            if ((checkBoxВДеле.Checked == true && checkBoxКонтроль.Checked == true && textBoxРезультатВыполнения.Text.Trim().Length > 0) || (checkBoxВДеле.Checked == true && checkBoxКонтроль.Checked == false && textBoxРезультатВыполнения.Text.Trim().Length > 0))
            {
                строкаКарточки["ВДело"] = true;
                строкаКарточки["НаКонтроле"] = true;
            }

            #endregion

            #region Номер Входящий
            if (textBoxНомерВходящий.Text == "б/н" || textBoxНомерВходящий.Text == "б.н" || textBoxНомерВходящий.Text == "б.н." || textBoxНомерВходящий.Text == "бн")
            {
                строкаКарточки["НомерВход"] = "б/н";
            }
            else
            {
                // Проверим корректность введённого номера.
                string[] arr = textBoxНомерВходящий.Text.Split('-');

                bool flagErrorNuber = true;
                foreach (string sKey in this.numbersDepartment.Keys)
                {
                    //if(arr[3].Trim() == sKey.Trim())
                    if (arr[2].Trim() == sKey.Trim())
                    {
                        // Установим флаг ошибки в true.
                        flagErrorNuber = false;
                    }
                }

                // Если ошибка то сообщим об этом пользователю.
                if (flagErrorNuber == true)
                {
                    MessageBox.Show(this,
                           "Неверно указан номер входящий",
                           "Ошибка номера",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error);
                    this.DialogResult = DialogResult.None;
                    return;
                }

                // Старая реализация генерации номера документа, оставим вдруг понадобиться.
                //string[] arr = textBoxНомерВходящий.Text.Split('/');


                //if (arr.Length != 2)
                //{
                //    MessageBox.Show(
                //                        this,
                //                       "Неверно указан номер входящий",
                //                       "Ошибка номера",
                //                       MessageBoxButtons.OK,
                //                       MessageBoxIcon.Error
                //                   );
                //    this.DialogResult = DialogResult.None;
                //    return;
                //}
                //else
                //{
                //    if (Information.IsNumeric(arr[0]))
                //    {
                //        if (Convert.ToInt32(arr[0]) > следНомерПП)
                //        {
                //            MessageBox.Show(this,
                //               "Неверно указан порядковый входящий номер. Вы можете указать число не больше чем " + следНомерПП,
                //               "Ошибка номера",
                //               MessageBoxButtons.OK,
                //               MessageBoxIcon.Error);
                //            this.DialogResult = DialogResult.None;
                //            return;

                //        }
                //        else if (Convert.ToInt32(arr[0]) < следНомерПП && новыйДокумент)
                //        {
                //            DialogResult result = MessageBox.Show(this,
                //                "Вы указали порядковый входящий номер, который не соответствует рекомендуемому.\nЕсли вы оставите введенный номер, возможно дублирование номеров в базе данных",
                //                "Возможно дублирование номеров",
                //                MessageBoxButtons.YesNo,
                //                MessageBoxIcon.Warning,
                //                MessageBoxDefaultButton.Button2);
                //            if (result == DialogResult.No)
                //            {
                //                this.DialogResult = DialogResult.None;
                //                return;
                //            }
                //        }


                //        //СледующийНомерДокумента
                //        строкаКарточки["НомерПП"] = arr[0];
                //        строкаКарточки["НомерВход"] = arr[1];
                //    }
                //    else
                //    {
                //        MessageBox.Show(this,
                //           "Неверно указан номер входящий",
                //           "Ошибка номера",
                //           MessageBoxButtons.OK,
                //           MessageBoxIcon.Error);
                //        this.DialogResult = DialogResult.None;
                //        return;
                //    }
                //}

                if (FlagUpdateDocument == false)
                {
                    if (flagAutoNumberDocStoip == false)
                    {

                        // Проверим флаг ввода номера документа вручную.
                        //if (this.flagNumberDoc.Checked == true)
                        //{
                        //    FormNumberDoc fNumDoc = new FormNumberDoc();
                        //    fNumDoc.ShowDialog();

                        //    if (fNumDoc.DialogResult == DialogResult.OK)
                        //    {
                        //        строкаКарточки["НомерПП"] = fNumDoc.NumberDoc;
                        //    }
                        //    else if (fNumDoc.DialogResult == DialogResult.Cancel)
                        //    {
                        //        return;
                        //    }
                        //}
                        //else
                        //{
                        //    строкаКарточки["НомерПП"] = СледующийНомерДокумента.Номер;
                        //}

                        строкаКарточки["НомерПП"] = СледующийНомерДокумента.Номер;
                        //строкаКарточки["НомерВход"] = СледующийНомерДокумента.Префикс + textBoxНомерВходящий.Text.Trim();
                        строкаКарточки["НомерВход"] = textBoxНомерВходящий.Text.Trim();

                        // Поместим в свойство формы номер нового документа.
                        НомерДокумента nextNumDoc = this.СледующийНомерДокумента;

                        //if (this.flagNumberDoc.Checked == true)
                        //{
                        //    FormNumberDoc fNumDoc = new FormNumberDoc();

                        //    if (this.flagLastNumberDoc == true)
                        //    {
                        //        fNumDoc.NumberDoc = this.lastNumberDoc;
                        //    }
                        //    fNumDoc.ShowDialog();

                        //    if (fNumDoc.DialogResult == DialogResult.OK)
                        //    {
                        //        // Пользователь установил ручной ввод нумерации документов.
                        //        if (this.flagLastNumberDoc == false)
                        //        {
                        //            nextNumDoc.Номер = Convert.ToInt16(fNumDoc.NumberDoc);
                        //            строкаКарточки["НомерПП"] = Convert.ToInt16(fNumDoc.NumberDoc);
                        //        }
                        //        else
                        //        {
                        //            // Восстановление автоматической нумерации документов после ручного ввода.
                        //            fNumDoc.NumberDoc = string.Empty;

                        //            // Установим нормальную нумерациию документиа.
                        //            fNumDoc.NumberDoc = this.lastNumberDoc;

                        //            // Отобразим последующую.


                        //            nextNumDoc.Номер = Convert.ToInt16(fNumDoc.NumberDoc);
                        //            строкаКарточки["НомерПП"] = Convert.ToInt16(fNumDoc.NumberDoc);
                        //        }
                        //    }
                        //}
                        //else
                        //{
                        //    nextNumDoc.Номер = СледующийНомерДокумента.Номер;
                        //}

                        //nextNumDoc.Префикс = СледующийНомерДокумента.Префикс + textBoxНомерВходящий.Text.Trim();
                        nextNumDoc.Префикс = textBoxНомерВходящий.Text.Trim();

                        // Заполним для отображения пользователю следующий номер документа.
                        СледующийНомерДокумента = nextNumDoc;
                    }
                    else
                    {
                        // Поместим в свойство формы номер нового документа.
                        НомерДокумента nextNumDoc = this.СледующийНомерДокумента;

                        // Узнаем номер и префикс документа.
                        this.textBoxНомерВходящий2.Text.Split('/')[0].Trim();

                        string[] arrayNum  = this.textBoxНомерВходящий2.Text.Split('/');

                        nextNumDoc.Номер = Convert.ToInt32(arrayNum[0]);

                        //nextNumDoc.Префикс = СледующийНомерДокумента.Префикс + textBoxНомерВходящий.Text.Trim();
                        nextNumDoc.Префикс = arrayNum[1].Trim();

                        // Заполним для отображения пользователю следующий номер документа.
                        СледующийНомерДокумента = nextNumDoc;
                    }
                }
                else
                {
                    строкаКарточки["НомерПП"] = СледующийНомерДокумента.Номер;

                    string[] arry = textBoxНомерВходящий.Text.Trim().Split('/');

                    // Старая реализация.
                    //строкаКарточки["НомерВход"] = arry[1].Trim();
                    строкаКарточки["НомерВход"] = arry[0].Trim();

                    string iTest = строкаКарточки["НомерВход"].ToString().Trim();


                    // Поместим в свойство формы номер нового документа.
                    НомерДокумента nextNumDoc = new НомерДокумента();
                    nextNumDoc.FlagUpdate = true;
                    nextNumDoc.Номер = СледующийНомерДокумента.Номер;
                    nextNumDoc.Префикс = textBoxНомерВходящий.Text.Trim();

                    // Заполним для отображения пользователю следующий номер документа.
                    СледующийНомерДокумента = nextNumDoc;
                }
            }
            #endregion

            #region Номер Исходящий
            if (textBoxНомерИсходящий.Text == "")
            {
                строкаКарточки["НомерИсход"] = "б/н";
            }
            else
            {
                строкаКарточки["НомерИсход"] = textBoxНомерИсходящий.Text;
            }
            #endregion

            строкаКарточки["ДатаИсхода"] = dateTimeДатаОтправления.Value.ToShortDateString();

            строкаКарточки["ДатаПоступ"] = dateTimeДатаПоступления.Value.ToShortDateString();

            строкаКарточки["КраткоеСодержание"] = textBoxСодержание.Text;

            строкаКарточки["Резолюция"] = textBoxРезолюция.Text;

            string myTest = textBoxРезультатВыполнения.Text;

            строкаКарточки["РезультатВыполнения"] = textBoxРезультатВыполнения.Text;

            строкаКарточки["FlagPersonData"] = false;

            строкаКарточки.FlagCardRepeet = this.chBoxRepet.Checked;

            //// Передадим в свойство id Цель получения данных.
            //ПодключитьБД coonectDB = new ПодключитьБД();
            //string sConn = coonectDB.СтрокаПодключения();

            //string query = "select [id_цельПолученияПерсДанных]from ЦельПолученияПерсональныхДанных " +
            //               "where [ЦельПолученияПерсональныхДанных] = '" + this.cmbBox.Text + "' ";

            //GetDataTable getTable = new GetDataTable(query);
            //DataTable tabPD = getTable.DataTable();

            //this.IdЦельПолученияПерсональныхДанных = Convert.ToInt32(tabPD.Rows[0][0]);

            // Сохраним в форме значение флага указывающего, что данная запись будет иметь повторяющиеся исходящие записи.


            string ttt = dateTimeКонтроль.Value.ToShortDateString();

            this.FlagRecordRepeet = this.chBoxRepet.Checked;

            if (this.chBoxRepet.Checked == true)
            {
                this.IncrementDate = Convert.ToInt32(txtPeriod.Text.Trim()) - 2;

                // Прибавим количество дней.
                строкаКарточки["СрокВыполнения"] = dateTimeКонтроль.Value.AddDays(this.IncrementDate);

                string sTestDate = Convert.ToDateTime(строкаКарточки["СрокВыполнения"]).ToShortDateString();
            }

            // Если уазано что документ со временем понадобиться дополнять.
            if (this.chcDop.Checked == true)
            {
                this.FlagAddDoc = true;
            }
            else
            {
                this.FlagAddDoc = false;
            }

            if (this.chboxDsp.Checked == true)
            {
                this.FlagDsp = "True";
            }
            else
            {
                this.FlagDsp = "False";
            }

            // Выведим окно которое содержит список видов поступления документов.
            FormTypeCompanyDocument formType = new FormTypeCompanyDocument();
            formType.ShowDialog();

            if (formType.DialogResult == DialogResult.OK)
            {
                // Передадим в форму способ поступления документа.
                СпособПоступления = formType.СпособПоступления;
            }

            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //FormСоставПерсДанных form = new FormСоставПерсДанных();
            //form.Idкарточки = this.Idкарточки;
            //form.TopMost = true;
            //form.ShowDialog();

            //if (form.DialogResult == DialogResult.OK)
            //{
            //    this.QueryStringУчётПерсДанных = form.СвязующаяУчётПерсональныхДанных;
            //}

            FormПерсональныеДанные form = new FormПерсональныеДанные();
            form.IdКарточки = this.Idкарточки;
            form.TopMost = true;
            form.ShowDialog();

            if (form.DialogResult == DialogResult.OK)
            {
                this.ConfigDatePerosnal = form.КонфигурированиеПерсональныхДанных;
            }

        }

        private void FormКарточка_Load(object sender, EventArgs e)
        {
            //// Заполним раскрывающийся список: цель получения персональных данных информацией из базы данных.
            // // Поучим срдерждимое справочника персональных данных.
            //ПодключитьБД coonectDB = new ПодключитьБД();
            //string sConn = coonectDB.СтрокаПодключения();

            //string query = "select [id_цельПолученияПерсДанных],[ЦельПолученияПерсональныхДанных] from ЦельПолученияПерсональныхДанных";

            //GetDataTable getTable = new GetDataTable(query);
            //DataTable tabPD = getTable.DataTable();

            //this.cmbBox.DataSource = tabPD;
            //this.cmbBox.DisplayMember = "ЦельПолученияПерсональныхДанных";
            //this.cmbBox.ValueMember = "id_цельПолученияПерсДанных";
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //ОтметкаПередачаОтказ отметка = new ОтметкаПередачаОтказ();
            //отметка.Отметка = true;
            //отметка.ПричиныОтказа = "NULL";

            //// Сохраним экземляр класса ОтметкаПередачаОтказ в свойстве.
            //this.ПередачаОтказ = отметка;

            //this.lblMarks.Text = "Передать";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //FormОтказ fo = new FormОтказ();
            //fo.ShowDialog();

            //if (fo.DialogResult == DialogResult.OK)
            //{
            //    ОтметкаПередачаОтказ отметка = new ОтметкаПередачаОтказ();
            //    отметка.Отметка = false;
            //    отметка.ПричиныОтказа = fo.ТекстОтказа;

            //    // Сохраним экземляр класса ОтметкаПередачаОтказ в свойстве.
            //    this.ПередачаОтказ = отметка;


            //    this.lblMarks.Text = "Отказать";
            //}
        }

        private void chBoxRepet_CheckedChanged(object sender, EventArgs e)
        {
            if (this.chBoxRepet.Checked == true)
            {
                this.txtPeriod.Enabled = true;
            }
            else
            {
                this.txtPeriod.Enabled = false;
            }
        }

        private void txtPeriod_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != 8 && (e.KeyChar < 48 || e.KeyChar > 57))
                e.Handled = true;

            //if (!Char.IsNumber(e.KeyChar))
            //{
            //    e.Handled = true;
            //}

            //if (!Char.IsDigit(e.KeyChar))
            //    e.Handled = true;  
        }

        private void btnTie_Click(object sender, EventArgs e)
        {
            try
            {
                // Получим путь к папке внутри которой нужно создать папку с номером документов.
                string patchDir = ConfigurationSettings.AppSettings["локальнаПапкаДокументооборот"].Trim();

                //// Название директории для хранения документа.
                string nameDir = this._НомерДокументаПовторПривязка.Trim().Replace("/", "-") + "-id" + this.Idкарточки.ToString().Trim();

                // Получим информацию о каталоге хранения.
                DirectoryInfo dirInfo = new DirectoryInfo(patchDir);

                // Создадим поддирректорию.
                dirInfo.CreateSubdirectory(nameDir);

                string query = "update Карточка " +
                                "set NameFileDocument = NULL, " +
                                "DataWriterServerDoc = NULL, " +
                                "NameFileDocumentVipNetEmailTitlePage = NULL " +
                                "where id_карточки = " + this.Idкарточки + " ";

                ExecuteQuery exq = new ExecuteQuery(query);
                exq.Excecute();

                MessageBox.Show("Пака создана");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            #region Старая реализация

            //// Укажем что можно записывать файл на сервер.
            //ФлагЗаписиАрхива = true;

            //// Установим свойство формы оботкрытии диалогового окна в true.
            //this.SaveDocServer = true;

            //string docServer = ConfigurationSettings.AppSettings["выбранныйСервер"].ToString();

            //// Откроем окно файлового диалога.
            //FolderBrowserDialog openFileDialog1 = new FolderBrowserDialog();

            ////OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    // Получим путь к файлу.
            //    fileName = openFileDialog1.SelectedPath;

            //    // Оставим код если вдруг будет решено вернуться к архивированию директории.
            //    DirectoryInfo dif = new DirectoryInfo(fileName);
            //    if (dif.GetFiles().Length == 0)
            //    {
            //        MessageBox.Show("Копируемая папка пуста!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            //        // В случае если папка с документом пустая, запись архива запрещаем.
            //        ФлагЗаписиАрхива = false;

            //        return;
            //    }

            //    // Запишем в свойство название файла.
            //    FileName = fileName.Trim();

            //    string расширение = Path.GetExtension(fileName);

            //    // Присовим новое имя файла.
            //    string newFileName = Guid.NewGuid().ToString().Trim() + расширение;

            //    // Путь файла копируемого на сервер.
            //    //this.PathFileServer = docServer + Path.GetFileName(newFileName).Trim();
            //    this.PathFileServer = newFileName;

            //}

            #endregion
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ПодключитьБД connDb = new ПодключитьБД();

            using (SqlConnection connection = new SqlConnection(connDb.СтрокаПодключения()))
            {
                connection.Open();
                string sql = "select FileDate from КарточкиДокументы where id_карточки = " + this.Idкарточки + "";
                SqlCommand command = new SqlCommand(sql, connection);
                
                SqlDataReader reader = command.ExecuteReader();


                while (reader.Read())
                {
                    fileByteArray = (byte[])reader["FileDate"];
                    //fileByteArray = (byte[])command.ExecuteScalar();
                    //fileByteArray = (byte[])reader["FileDate"];
                }

                // Массив битов данных из БД.
                //byte[] fileByteArray = (byte[])command.ExecuteScalar();

                string dir = @"d:\Recor";

                string fileName = dir + @"\TempView.zip";

                FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                BinaryWriter binWriter = new BinaryWriter(fileStream);
                binWriter.Write(fileByteArray);
                binWriter.Close();

                // Откроем архив.
                System.Diagnostics.Process.Start(fileName);
            }

            //// Получим имя файла на сервере.
            //string имяФайлНаСервере = this.PathFileServer.Trim() + ".zip";

            //try
            //{

            //    // Путь к сервреру.
            //    string путКСерверу = patchServerSave.Trim();

            //    // Получим путь к файлу на серврер.
            //    string fileServer = путКСерверу + @"\" + имяФайлНаСервере.Replace("/", "-");

            //    // Путь к файлу во временной папке.
            //    string tempPath = Path.GetTempPath();

            //    // Получим путь и имя файла которое он будет иметь после копирования во временную таблицу на клиент.
            //    string fileTo = tempPath + имяФайлНаСервере;

            //    // Скопируем архив во временную папку.
            //    File.Copy(fileServer, fileTo, true);

            //    // Получим путь к файлу во временной таблице на клиенте.
            //    string fileTemp = tempPath + @"\" + имяФайлНаСервере;

            //    // Откроем архив.
            //    System.Diagnostics.Process.Start(fileTemp);
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}

        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                ПодключитьБД connDb = new ПодключитьБД();

                using (SqlConnection connection = new SqlConnection(connDb.СтрокаПодключения()))
                {
                    connection.Open();
                    string sql = "select FileDateTitlePage from КарточкиДокументы where id_карточки = " + this.Idкарточки + "";
                    SqlCommand command = new SqlCommand(sql, connection);
                    //SqlDataReader reader = command.ExecuteReader();

                    // Массив битов данных из БД.
                    byte[] fileByteArray = (byte[])command.ExecuteScalar();

                    string dir = @"d:\Recor";

                    string fileName = dir + @"\TempView.zip";

                    FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                    BinaryWriter binWriter = new BinaryWriter(fileStream);
                    binWriter.Write(fileByteArray);
                    binWriter.Close();

                    // Откроем архив.
                    System.Diagnostics.Process.Start(fileName);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //    // Получим имя файла на сервере.
            //    string имяФайлНаСервере = this.pathFileServerTitlePage.Trim() + ".zip";

            //    try
            //    {

            //    // Путь к сервреру.
            //    string путКСерверу = patchServerSave.Trim();

            //    // Получим путь к файлу на серврер.
            //    string fileServer = путКСерверу + @"\ТитульныеКарточки\" + имяФайлНаСервере.Replace("/", "-");

            //    // Путь к файлу во временной папке.
            //    string tempPath = Path.GetTempPath();

            //    // Получим путь и имя файла которое он будет иметь после копирования во временную таблицу на клиент.
            //    string fileTo = tempPath + имяФайлНаСервере;

            //    // Скопируем архив во временную папку.
            //    File.Copy(fileServer, fileTo, true);

            //    // Получим путь к файлу во временной таблице на клиенте.
            //    string fileTemp = tempPath + @"\" + имяФайлНаСервере;

            //    // Откроем архив.
            //    System.Diagnostics.Process.Start(fileTemp);
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
        }

        private void FormКарточка_FormClosing(object sender, FormClosingEventArgs e)
        {
            string dir = @"d:\Recor";

            DirectoryInfo dirInf = new DirectoryInfo(dir);

            if (dirInf.Exists == true)
            {
                string sTest = dirInf.FullName;

                foreach (FileInfo fi in dirInf.GetFiles())
                {
                    if (fi.Name.Trim().ToLower() == "TempView.zip".Trim().ToLower())
                    {
                        fi.Delete();
                    }
                }
            }
            else
            {
                MessageBox.Show(@"Создайте папку d:\Recor ");
            }

            

        }

        private void FormКарточка_FormClosed(object sender, FormClosedEventArgs e)
        {
            string dir = @"d:\Recor";

            DirectoryInfo dirInf = new DirectoryInfo(dir);

            if (dirInf.Exists == true)
            {

                string sTest = dirInf.FullName;

                foreach (FileInfo fi in dirInf.GetFiles())
                {
                    if (fi.Name.Trim().ToLower() == "TempView.zip".Trim().ToLower())
                    {
                        fi.Delete();
                    }
                }
            }
            else
            {
                MessageBox.Show(@"Создайте папку d:\Recor");
            }
        }

        private void chboxDsp_CheckedChanged(object sender, EventArgs e)
        {
           if (this.chboxDsp.Checked == true)
           {
               this.textBoxНомерВходящий.Mask = "00-00-00-00-aaa";
               textBoxНомерВходящий.Text += "дсп";
               this.FlagDsp = "True";
           }
           else
           {
               string stest = sNumStart;
               this.textBoxНомерВходящий.Mask = "00-00-00-00";
               textBoxНомерВходящий.Text = stest;
               this.FlagDsp = "False";
           }
        }

        private void textBoxНомерВходящий_TextChanged(object sender, EventArgs e)
        {
            //if (textBoxНомерВходящий.Text.Length == 11)
            //{
            //    this.chboxDsp.Enabled = true;
            //}
            //else
            //{
            //    this.chboxDsp.Enabled = true;
            //}
            sNumStart = this.textBoxНомерВходящий.Text;
        }

        /// <summary>
        /// Загружает список отделов.
        /// </summary>
        private void LoadNumberDepartments()
        {
            string query = "select НомерПодразделения from ПодразделенияКомитета";

            foreach(DataRow row in DataTableSql.GetDataTableRows(query))
            {
                numbersDepartment.Add(row["НомерПодразделения"].ToString().Trim(), row["НомерПодразделения"].ToString().Trim());
            }
        }

        private void rb04_CheckedChanged(object sender, EventArgs e)
        {
             textBoxНомерВходящий.Text = "04-";
        }

        private void rb02_CheckedChanged(object sender, EventArgs e)
        {
            textBoxНомерВходящий.Text = "02-";
        }

        private void btnLastNumber_Click(object sender, EventArgs e)
        {
            #region Старый функционал пока оставим
            /*
            string query = "select MAX(номерПП) from Карточка " +
                           "where YEAR(ДатаПоступ) = "+ this.CurrentYear +" ";

            GetDataTable tab = new GetDataTable(query);
            string numLastDoc = tab.DataTable("SelectedYear").Rows[0][0].ToString();

            int lastNumberDocumnet = Convert.ToInt32(numLastDoc) + 1;

            numLastDoc = string.Empty;
            numLastDoc = lastNumberDocumnet.ToString();

            MessageBox.Show("Следующий номер документа  - " + numLastDoc.ToString());

            this.lastNumberDoc = numLastDoc;

            flagLastNumberDoc = true;

            label1.Text = "След. номер п\\п " + (this.lastNumberDoc);

            //this.flagNumberDoc.Checked = true;
             */

            #endregion

            bool flagEdit = новыйДокумент;

            // Флаг указывает что мы работаем с карточкой входящих документов.
            bool flagInputCard = true;

            FormОснованиеПередачи formОснование = new FormОснованиеПередачи(this.Idкарточки, flagEdit, flagInputCard);

            // Обнулим список ОснованиеПередачи перед использованием.
            formОснование.ListОснованиеПередачи.Clear();

            // Передадим в форму id карточки.
            formОснование.IdКарточки = this.Idкарточки;

            formОснование.ShowDialog();

            if (formОснование.DialogResult == DialogResult.OK)
            {
                // Т.е. добавляем новую карточку.
                if (flagEdit == true && flagInputCard == true)
                {
                    // Сформируем строку запроса кна добавления в связующую таблицу, осноавния для передачи персональных данных.
                    IQueryStringSQL queryInsert = new InsertQueryОснованиеПередачи(formОснование.ListОснованиеПередачи, this.Idкарточки);

                    // Передадим в форму: Карточка строку запроса на добавления оснований к передаче персональных данных.
                    this.QueryPersonDateForCardInput = queryInsert.Query();
                }
                else if (flagEdit == false && flagInputCard == true)
                {
                    // Сформируем строку запроса для обновления связующей таблицы.
                    IQueryStringSQL queryUpdate = new UpdateQueryОснованиеПередачи(formОснование.ListОснованиеПередачи, this.Idкарточки);

                    // Передадим в форму: Карточка строку запроса на добавления оснований к передаче персональных данных.
                    this.QueryPersonDateForCardInput = queryUpdate.Query();

                }
            }

            // Отобразим форму для выбора основания для передачи персональных данных
        }

    }
}
