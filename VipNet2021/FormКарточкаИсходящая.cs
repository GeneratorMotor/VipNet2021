/*
 * User: Денис Николаевич
 * Date: 25.03.2007
 * Time: 20:33
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */

using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using Microsoft.VisualBasic;
using System.Text;
using System.Collections.Generic;
using RegKor.Classess;
using System.Data.SqlClient;
using System.IO;
using Ionic.Zip;
using RegKor.Classess;



namespace RegKor
{
    /// <summary>
    /// Description of FormКарточкаИсходящая.
    /// </summary>
    public partial class FormКарточкаИсходящая
    {
        /// <summary>
        /// Представляет собой строку из таблицы "КарточкаИсходящая"
        /// </summary>
        public readonly RegKor.DS1.КарточкаИсходящаяRow строкаИсходящейКарточки;

        /// <summary>
        /// True если новый документ, False если изменение существующего
        /// </summary>
        private bool новыйДокумент;

        /// <summary>
        /// Номер порядковый, который должен быть вставлен в базу
        /// </summary>
        private int следНомерПП;

        /// <summary>
        /// Номер комитета для сохранения
        /// </summary>
        private string номерКомитета;// = "13-01" ;// = 7;

        /// <summary>
        /// Номер подразделения для сохранения
        /// </summary>
        private char[] номерПодразделения = new char[2];

        /// <summary>
        /// Номер номенклатурный для сохранения
        /// </summary>
        private char[] номерНоменклатурный = new char[2];

        /// <summary>
        /// Номер буквенное обозначение для сохранения
        /// </summary>
    //    private char[] буквенноеОбозначение = new char[1];

        /// <summary>
        /// Номер порядковый для сохранения
        /// </summary>
        private int номерПП = 0;

        /// <summary>
        /// Документ по которому дали ответ, для сохранения
        /// </summary>
        private int ИДВходящегоДокумента = -1;

        /// <summary>
        /// 
        /// </summary>
        private int ИДСтарогоВходящегоДокумента = -1;

        private bool flagОтветПисьмо;
        private string flagListBase = "";

        private string queryInsert = string.Empty;

        /// <summary>
        /// Хранит SQL инструкцию на вставку в таблицу [СвязующаяЦельПолучениперсональныхДанных].
        /// </summary>
        public string QueryInsert
        {
            get
            {
                return queryInsert;
            }
            set
            {
                queryInsert = value;
            }
        }
        


        /// <summary>
        /// Флаг в состоянии TRUE указывает на то что мы отправляем ответ на письмо.
        /// </summary>
        public bool FlagОтветПисьмо
        {
            get
            {
                return flagОтветПисьмо;

            }
            set
            {
                flagОтветПисьмо = value;
            }
        }

        private List<ОснованиеПередачи> listProperty = new List<ОснованиеПередачи>();

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

        /// <summary>
        /// Устанавливает что список с выбранными основаниями не пустой.
        /// </summary>
        //public string FlagListBase
        //{
        //    get
        //    {
        //        return flagListBase;
        //    }
        //    set
        //    {
        //        flagListBase = value;
        //    }
        //}


        private string адресатPropery = string.Empty;

        /// <summary>
        /// Хранит адресат пользователя которому оправляется письмо.
        /// </summary>
        public string Адресат
        {
            get
            {
                return адресатPropery;
            }
            set
            {
                адресатPropery = value;
            }
        }

        // Строка для хранения результатов выбора входящих документов.
        StringBuilder build = new StringBuilder();

        /// <summary>
        /// Хранит ответы на входящие документы.
        /// </summary>
        public StringBuilder ОтветНаВходящиеДокументы
        {
            get
            {
                return build;
            }
            set
            {
                build = value;
            }
        }

        private List<int> listId = new List<int>();

        /// <summary>
        /// Хранит список id входящих документов на которые производится ответ.
        /// </summary>
        public List<int> ListIDКарточки
        {
            get
            {
                return listId;
            }
            set
            {
                listId = value;
            }
        }

       

        private List<int> listIdСвязующаяКарточкаВходящаяИсходящая = new List<int>();

        /// <summary>
        /// Хранить id таблицы СвязующаяКарточкаВходящаяИсходящая
        /// </summary>
        public List<int> ListIDСвязующаяКарточкаВходящаяИсходящая
        {
            get
            {
                return listIdСвязующаяКарточкаВходящаяИсходящая;
            }
            set
            {
                listIdСвязующаяКарточкаВходящаяИсходящая = value;
            }
        }

        private List<int> listIdСвязующаяЦельПолучениперсональныхДанных = new List<int>();

        /// <summary>
        /// Хранить id таблрицы СвязующаяЦельПолучениперсональныхДанных
        /// </summary>
        public List<int> ListIDСвязующаяЦельПолучениперсональныхДанных
        {
            get
            {
                return listIdСвязующаяЦельПолучениперсональныхДанных;
            }
            set
            {
                listIdСвязующаяЦельПолучениперсональныхДанных = value;
            }
        }

        private int id_карточкаИсход = 0;

        // Хранит id исходящей карточки для 
        //private int idИсходящейКарточки = 0;

        /// <summary>
        /// Хранит id карточки исходящей.
        /// </summary>
        public int IdКарочкаИсходящая
        {
            get
            {
                return id_карточкаИсход;
            }
            set
            {
                id_карточкаИсход = value;
            }
        }

        private string номерИсходящий = string.Empty;

        /// <summary>
        /// Хранит нмер исходящего документа.
        /// </summary>
        public string НомерИсходящий
        {
            get
            {
                return номерИсходящий;
            }
            set
            {
                номерИсходящий = value;
            }
        }


        private string префикаНомерИсходящий = string.Empty;

        /// <summary>
        /// Номер исходящий.
        /// </summary>
        public string ПрефиксНомерИсходящий
        {
            get
            {
                return префикаНомерИсходящий;
            }
            set
            {
                префикаНомерИсходящий = value;
            }
        }

        private string strGuid = string.Empty;

        /// <summary>
        /// Уникальный номер записи.
        /// </summary>
        public string StrGuid
        {
            get
            {
                return strGuid;
            }
            set
            {
                strGuid = value;
            }
            
        }

        private string имяДокумента = string.Empty;
        
        /// <summary>
        /// Хранит имя документа находящегося в базе данных.
        /// </summary>
        public string ИмяДокумента
        {
            get
            {
                return имяДокумента;
            }
            set
            {
                имяДокумента = value;
            }
        }

        // Путь к серверу.
        private string patchServerSave = string.Empty;

        private bool flagEdit;

        /// <summary>
        /// Флаг указывает, что карточка просматиривается для редактирования.
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

        private НомерДокумента numDoc;

        /// <summary>
        /// Хранит номер следующего документа.
        /// </summary>
        public НомерДокумента НомерDoc
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

        private ItemСпособПоступленияДокумента способПоступления;

        /// <summary>
        /// Хранит способ поступления документа.
        /// </summary>
        public ItemСпособПоступленияДокумента СпособПоступления
        {
            get
            {
                return способПоступления;
            }
            set
            {
                способПоступления = value;
            }
        }

        private bool flagStopNum = false;


        /// <summary>
        /// Устанавливает реждим работы автомата
        /// </summary>
        public bool FlagNumStopDoc
        {
            get
            {
                return flagStopNum;
            }
            set
            {
                flagStopNum = value;
            }
        }

        private int numDocA = 0;

        /// <summary>
        /// Хранит номер документа.
        /// </summary>
        public int NumDocNoAutomat
        {
            get
            {
                return numDocA;
            }
            set
            {
                numDocA = value;
            }
        }

        // Переменная для хранения номера исходящего документа.
        private string sNumStart;

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


        /// <summary>
        /// Конструктор, используется при создании нового исходящего документа
        /// </summary>
        /// <param name="dataset">датасет с данными</param>
        public FormКарточкаИсходящая(DS1 dataset,string выбранныйГод,bool flagAutoNumStop)
        {
            // Установим свойство для включения (отключения) автоматической нумерации номеров документов.
            this.FlagNumStopDoc = flagAutoNumStop;

            номерПодразделения[0] = 'x';
            номерПодразделения[1] = 'x';

         //   буквенноеОбозначение[0] = 'x';

            InitializeComponent();

            this.ds11 = dataset;
            comboBoxАдресат.DataSource = ds11.Корреспонденты;
            comboBoxАдресат.DisplayMember = ds11.Корреспонденты.Columns["ОписаниеКорреспондента"].ToString();
            comboBoxАдресат.ValueMember = ds11.Корреспонденты.Columns["id_корреспондента"].ToString();

            // Это новый документ
            новыйДокумент = true;
             //maskedTextBox1.Mask = @"00-00-00-00\/09999";// "";
             maskedTextBox1.Mask = @"00-00-00\/09999";// "";

            
             //doc.Префикс = префиксДокумента;
             НомерДокумента doc = new НомерДокумента();


            // Получение максимального номераПП из таблицы КарточкаИсходящая
             //DataRow[] dr = ds11.КарточкаИсходящая.Select("Дата >='01.12.2011'", "НомерПорядковый DESC");
             
            //DataRow[] dr = ds11.КарточкаИсходящая.Select("Дата >='01.12." + выбранныйГод + "' ", "НомерПорядковый DESC");


             //string query = "SELECT * FROM [КарточкаИсходящая] " +
             //               " where Дата >= '" + (Convert.ToInt32(выбранныйГод) + 1).ToString().Trim() + "1201' and Дата <= '" + (Convert.ToInt32(выбранныйГод) + 1).ToString().Trim() + "1231' " +
             //               //" where Дата >= '" + выбранныйГод + "1231' " +
             //               //" and id_карточки in (SELECT MAX(id_карточки) FROM [КарточкаИсходящая] " +
             //               //" where FlagAutho is null) " +
             //               //" order by НомерПорядковый asc ";
             //                   "order by id_карточки desc ";

             string query = "SELECT * FROM [КарточкаИсходящая] " +
                         " where Дата >= '" + (Convert.ToInt32(выбранныйГод) + 1).ToString().Trim() + "0101' and Дата <= '" + (Convert.ToInt32(выбранныйГод) + 1).ToString().Trim() + "1231' " +
                             "order by id_карточки desc ";

             DataRow[] dr = DataTableSql.GetDataTableRows(query);

            if (dr.Length > 0)
            {
                следНомерПП = 1 + (int)dr[0]["НомерПорядковый"];
                label1.Text = "Номер п\\п " + (следНомерПП);

                doc.Номер = следНомерПП;
            }
            else
            {
                следНомерПП = 1;
                label1.Text = "Номер п\\п " + (следНомерПП);

                doc.Номер = следНомерПП;
            }

            строкаИсходящейКарточки = ds11.КарточкаИсходящая.NewКарточкаИсходящаяRow();

            // Передадим в свойство объект содержащий следующий номер.
            this.НомерDoc = doc;

            //this.IdКарочкаИсходящая = строкаИсходящейКарточки.id_карточки;

            //if (FlagListBase == "")
            //{
            //    this.buttonСохранить.Enabled = true;
            //}

        }

        /// <summary>
        /// Конструктор, используется для изменения существующего исходящего документа
        /// </summary>
        /// <param name="dataset">датасет с данными</param>
        /// <param name="ИДИсходящегоДокумента">id исходящего документа</param>
        public FormКарточкаИсходящая(DS1 dataset, DS1.КарточкаИсходящаяRow строкаДляИзменения,string выбранныйГод)
        {
            InitializeComponent();

            this.ds11 = dataset;

            строкаИсходящейКарточки = строкаДляИзменения;

            // Устанавливаем дату отправления
            dateTimeДата.Value = (DateTime)строкаИсходящейКарточки["Дата"];

            //// Устанавливаем номер подразделения
            DataRow[] dr2 = ds11.ПодразделенияКомитета.Select("id_подразделения=" + (int)строкаИсходящейКарточки["id_Подразделения"]);
            номерПодразделения = dr2[0]["НомерПодразделения"].ToString().ToCharArray();

            //// Устанавливаем номер номенклатурный
            номерНоменклатурный = строкаИсходящейКарточки["НомерНоменклатурный"].ToString().ToCharArray();

            //// Устанавливаем буквенное обозначение подразделения
        //    буквенноеОбозначение = dr2[0]["БуквенноеОбозначение"].ToString().ToCharArray();

            //// Устанавливаем номер порядковый
            номерПП = (int)строкаИсходящейКарточки["НомерПорядковый"];

            //// Получение максимального номераПП из таблицы КарточкаИсходящая
            //DataRow[ ] dr3 = ds11.КарточкаИсходящая.Select( "", "НомерПорядковый DESC" );

            // Класс для хранения номера исходящего документа.
            НомерДокумента doc = new НомерДокумента();

            // Получение максимального номераПП из таблицы КарточкаИсходящая
            //DataRow[] dr = ds11.КарточкаИсходящая.Select("Дата >='01.12.2011'", "НомерПорядковый DESC");
            //DataRow[] dr = ds11.КарточкаИсходящая.Select("Дата >='01.12." + выбранныйГод + "'", "НомерПорядковый DESC");

            //string query = "SELECT * FROM [КарточкаИсходящая] " +
            //               " where Дата >= '" + (Convert.ToInt32(выбранныйГод) + 1).ToString().Trim() + "1201' and Дата <= '" + (Convert.ToInt32(выбранныйГод) + 1).ToString().Trim() + "1231' " +
            //               //" where Дата >= '" + выбранныйГод + "1231' " +
            //               //" and id_карточки in (SELECT MAX(id_карточки) FROM [КарточкаИсходящая] " +
            //               //" where FlagAutho is null) " +
            //               //" order by НомерПорядковый asc ";
            //                   "order by id_карточки desc ";

            string query = "SELECT * FROM [КарточкаИсходящая] " +
                        " where Дата >= '" + выбранныйГод + "1201' and Дата <= '" + (Convert.ToInt32(выбранныйГод) + 1).ToString().Trim() + "1231' " +
                            "order by id_карточки desc ";

            DataRow[] dr = DataTableSql.GetDataTableRows(query);


            if (dr.Length > 0)
            {
                следНомерПП = 1 + (int)dr[0]["НомерПорядковый"];
                label1.Text = "Номер п\\п " + (следНомерПП);

                doc.Номер = следНомерПП;
            }
            else
            {
                следНомерПП = 1;
                label1.Text = "Номер п\\п " + (следНомерПП);

                doc.Номер = следНомерПП;
            }
            //следНомерПП = 1 + (int)dr3[0]["НомерПорядковый"];

            this.НомерDoc = doc;

            // Устанавливаем содержание документа
            textBoxСодержание.Text = (string)строкаИсходящейКарточки["Содержание"];

            // Устанавливаем исходящий документ
            System.Data.DataRow[] карточкаИДок = ds11.ВыборкаИсходящихДокументов.Select("id_карточки=" + строкаИсходящейКарточки["id_карточки"]);

            int flagCountRow = 0;

            // Проверим есть ли в таблице [СвязующаяКарточкаВходящаяИсходящая] связанные данные.
            ПодключитьБД connectDB = new ПодключитьБД();
            using(SqlConnection con = new SqlConnection(connectDB.СтрокаПодключения()))
            {
                string quer = "select CONVERT(nvarchar, dbo.Карточка.НомерПП) + N'/' + RTRIM(LTRIM(CONVERT(nvarchar, " +
                              "dbo.Карточка.НомерВход))) AS НомерВход from Карточка " +
                              "where id_карточки in ( " +
                              "select distinct id_карточкаВходящая from dbo.СвязующаяКарточкаВходящаяИсходящая " +
                              "where id_карточкаИсходящая = " + Convert.ToInt32(строкаИсходящейКарточки["id_карточки"]) + ") ";

                DataSet dss = new DataSet();

                SqlDataAdapter da2 = new SqlDataAdapter(quer, con);
                da2.Fill(dss, "ПроверкаКолСтрок");

                // Получим данные из таблицы которые соответсвуют 
                System.Data.DataTable tab2 = dss.Tables["ПроверкаКолСтрок"];

                StringBuilder bb = new StringBuilder();

                flagCountRow = tab2.Rows.Count;

                foreach(DataRow r in tab2.Rows)
                {
                    bb.Append(r[0].ToString().Trim() + ",");
                }

                string qq = string.Empty;

                if (flagCountRow > 1)
                {
                    qq = "select ОписаниеКорреспондента from Корреспонденты " +
                                "where id_корреспондента in ( " +
                                "select id_корреспондента from Карточка " +
                                "where id_карточки in ( " +
                                "select distinct id_карточкаВходящая from dbo.СвязующаяКарточкаВходящаяИсходящая " +
                                "where id_карточкаИсходящая = " + Convert.ToInt32(строкаИсходящейКарточки["id_карточки"]) + ")) ";

                    SqlDataAdapter da22 = new SqlDataAdapter(qq, con);
                    da22.Fill(dss, "ПроверкаКолС");

                    // Получим данные из таблицы которые соответсвуют 
                    System.Data.DataTable tab22 = dss.Tables["ПроверкаКолС"];

                    string sText = tab22.Rows[0][0].ToString().Trim() + " дан ответ на " + bb.ToString().Trim();
                    textBoxОтветНаДокумент.Text = sText.Remove(sText.Length - 1, 1);
                }
                
                if(flagCountRow == 1)
                {
                    qq = "select ОписаниеКорреспондента from Корреспонденты " +
                                "where id_корреспондента in ( " +
                                "select id_корреспондента from Карточка " +
                                "where id_карточки in ( " +
                                "select distinct id_карточкаВходящая from dbo.СвязующаяКарточкаВходящаяИсходящая " +
                                "where id_карточкаИсходящая = " + Convert.ToInt32(строкаИсходящейКарточки["id_карточки"]) + ")) ";

                    SqlDataAdapter da22 = new SqlDataAdapter(qq, con);
                    da22.Fill(dss, "ПроверкаКолС");

                    // Получим данные из таблицы которые соответсвуют 
                    System.Data.DataTable tab22 = dss.Tables["ПроверкаКолС"];

                    string sText = tab22.Rows[0][0].ToString().Trim() + " дан ответ на " + bb.ToString().Trim();
                    textBoxОтветНаДокумент.Text = sText.Remove(sText.Length - 1, 1);

                    //textBoxОтветНаДокумент.Text = bb.ToString().Trim();

                    // Отобразим файл документа связанного с карточкой исходящей.
                    int id_карточка = строкаИсходящейКарточки.id_карточки;

                }

               

            }


            //SqlDataAdapter da2 = new SqlDataAdapter(quer, con);
            //da2.Fill(ds, "Проверка");

            //// Получим данные из таблицы которые соответсвуют 
            //System.Data.DataTable tab2 = ds.Tables["Проверка"];

            if (flagCountRow == 0)
            {
                object id = карточкаИДок[0]["id_ВходящегоДокумента"];
                if (!(id == System.DBNull.Value))
                {
                    ИДВходящегоДокумента = (int)id;
                    ИДСтарогоВходящегоДокумента = (int)id;
                    textBoxОтветНаДокумент.Text = ОтветНаДокумент(ИДВходящегоДокумента);
                }
            }

            // Устанавливаем адресат документа
            comboBoxАдресат.DataSource = ds11.Корреспонденты;
            comboBoxАдресат.DisplayMember = ds11.Корреспонденты.Columns["ОписаниеКорреспондента"].ToString();
            comboBoxАдресат.ValueMember = ds11.Корреспонденты.Columns["id_корреспондента"].ToString();
            comboBoxАдресат.Text = карточкаИДок[0]["ОписаниеАдресата"].ToString();

            // Это изменение сущ. документа
            новыйДокумент = false;


      /*      char[] нп = maskedTextBox1.Text.ToCharArray();
            if (нп.Length >= 4)
            {
                if (нп[3].Equals('7'))
                {
                    // Указан район, маска требует ввода буквенного обозначения
                    maskedTextBox1.Mask = @"0-00-00<L\/09999";
                }
                else
                {
                    maskedTextBox1.Mask = @"0-00-00\/09999";
                }
            }*/
  
          //  maskedTextBox1.Mask = @"00-00-00\/09999";// "";
            DataRow[] dr4 = ds11.ВыборкаИсходящихДокументов.Select("id_карточки=" + строкаИсходящейКарточки["id_карточки"]);
            
            // Удалим пустые символы из номера документа.
            string numDoc = dr4[0]["ТекстовыйНомер"].ToString();

            string[] docsStr = numDoc.Split(' ');

            StringBuilder numberDocument = new StringBuilder();

            foreach (string st in docsStr)
            {
                if (st.Length > 0)
                {
                    numberDocument.Append(st);
                }
            }

            maskedTextBox1.Text = numberDocument.ToString();// dr4[0]["ТекстовыйНомер"].ToString();

            if (maskedTextBox1.Text[0] == '1')
            {
               
                    maskedTextBox1.Text = maskedTextBox1.Text.Replace(" ", "");
                    maskedTextBox1.Mask = @"00-00-00-00\/09999";

            }

            errorProviderНомер.SetError(this.maskedTextBox1, "");

            // Установим флаг персональных данных.
            //DataRow[] drFPD = ds11.КарточкаИсходящая.Select("id_карточки" + строкаИсходящейКарточки["id_карточки"]);
            chkFlagPersonData.Checked = (bool)строкаИсходящейКарточки["FlagPersonData"];

            // Заполним свойство ListОснованиеПередачи данными по умолчанию.
            PersonDataDefault pdd = new PersonDataDefault(строкаИсходящейКарточки.id_карточки);
            ListОснованиеПередачи = pdd.GetList();

            if (ДокументооборотConfig.ВключитьДокументооборот() == true)
            {
                // Запрос для нахождения документа связанного с карточкой.
                string query2 = "select FileData,FileDateTitlePage from КарточкаИсходящаяДокументы " +
                               "where id_карточки = " + строкаИсходящейКарточки.id_карточки + " ";

                GetDataTable tab = new GetDataTable(query2);
                DataTable tabRow = tab.DataTable();

                if (tabRow.Rows.Count > 0)
                {
                    if (tabRow.Rows[0]["FileData"] != DBNull.Value)
                    {
                        this.linkLabel1.Text = "Открыть файлы документа";
                    }

                    if (tabRow.Rows[0]["FileDateTitlePage"] != DBNull.Value)
                    {
                        this.linkLabel2.Text = "Открыть титульный лист документа";
                    }
                    //else
                    //{
                    //    ИмяДокумента = null;
                    //}

                    // Узнаем документ ДСП или нет.
                    string queryUpdate = " SELECT ДСП FROM [КарточкаИсходящая] where id_карточки = " + строкаИсходящейКарточки.id_карточки + " ";

                    if (DataTableSql.GetDataTable(queryUpdate).Rows[0][0] != DBNull.Value)
                    {
                        if (Convert.ToBoolean(DataTableSql.GetDataTable(queryUpdate).Rows[0][0]) == true)
                        {
                            this.chkBoxDsp.Checked = true;
                            this.maskedTextBox1.Text += "-ДСП";

                        }
                    }
                }
            }

            //// Получим путь к серверу.
            //string patchServerQuery = "select PatchServer from СерверПуть";

            //GetDataTable tabServer = new GetDataTable(patchServerQuery);
            //DataTable tabServFile = tabServer.DataTable();

            //patchServerSave = tabServFile.Rows[0]["PatchServer"].ToString().Trim();


        }

        private void buttonОтмена_Click(object sender, EventArgs e)
        {
            Close();
        }


        private void buttonОтветНаДокумент_Click(object sender, EventArgs e)
        {
            FormСписокВходящиеДокументы form = new FormСписокВходящиеДокументы(ds11.Выборка);
            DialogResult result = form.ShowDialog(this);

            // Строка для хранения результатов выбора входящих документов.
            StringBuilder build = new StringBuilder();

            if (result == DialogResult.OK)
            {
                строкаИсходящейКарточки["id_ВходящегоДокумента"] = form.ИДВходящегоДокумента;
                ИДВходящегоДокумента = (int)form.ИДВходящегоДокумента;

                // Запишем в свойства формы список id входящих документов.
                ListIDКарточки.Add(ИДВходящегоДокумента);

                // Старая реализация ответа на входящие доркументы.
                //textBoxОтветНаДокумент.Text = ОтветНаДокумент(form.ИДВходящегоДокумента);
                

                // Соберём в свойстве формы несколько ответов на входящие документы.
                this.ОтветНаВходящиеДокументы.Append("; " + ОтветНаДокумент(form.ИДВходящегоДокумента));
                textBoxОтветНаДокумент.Text = this.ОтветНаВходящиеДокументы.ToString().Trim();

                // 
            }
        }

        /// <summary>
        /// Принимает id входящего документа и возвращает строку с его кратким описанием
        /// </summary>
        /// <param name="ИДВДокумента">id входящего документа</param>
        /// <returns>строку кратким описанием входящего документа</returns>
        private string ОтветНаДокумент(int ИДВДокумента)
        {
            System.Data.DataRow[] карточкаВДок = ds11.Выборка.Select("id_карточки=" + ИДВДокумента);

            if (карточкаВДок.Length == 0)
            {
                MessageBox.Show("Каротчка зарезервирована");
                return "";
            }
                string описаниеИДокумента = карточкаВДок[0]["ОписаниеДокумента"].ToString();
                string описаниеИКорреспондента = карточкаВДок[0]["ОписаниеКорреспондента"].ToString();
                string номерИВход = карточкаВДок[0]["НомерВход"].ToString();
                if (!(номерИВход == null || номерИВход == ""))
                {
                    return описаниеИДокумента + " от " + описаниеИКорреспондента + ", № вход. " + номерИВход;
                }
                return "";
            
        }

        /// <summary>
        /// Событие изменения текста в поле ввода номера
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void maskedTextBox1_TextChanged(object sender, EventArgs e)
        {
            if (this.chkBoxDsp.Checked == false)
            {
                sNumStart = this.maskedTextBox1.Text;
            }

            //if (this.chkBoxDsp.Checked == true)
            //{
            //    maskedTextBox1.Mask = @"00-00-00-00<L\/09999";
            //}
            //else
            //{
            //    maskedTextBox1.Mask = @"00-00-00-00\/09999";
            //}


            //char[] нп = maskedTextBox1.Text.ToCharArray();
        /*    if (нп.Length >= 7)
            {
                if (нп[7].Equals('7'))
                {
                    // Указан район, маска требует ввода буквенного обозначения
                    maskedTextBox1.Mask = @"00-00-00-00<L\/09999";
                }
                else
                {
                    maskedTextBox1.Mask = @"00-00-00-00\/09999";
                }
            }*/


        }

        private void maskedTextBox1_Leave(object sender, EventArgs e)
        {
            //ПроверкаНомера(maskedTextBox1.Text);
        }

        /// <summary>
        /// Открывает справочник адресатов, 
        /// если в комбике есть текст то справочник открывается на добавление, 
        /// а если комбик адресатов пустой, 
        /// то и справочник открывается на просмотр, 
        /// а дальше как юзер поступит
        /// </summary>
        private void СправочникАдресатов()
        {
            FormКорреспонденты form;
            string новыйАдресат = comboBoxАдресат.Text;

            if (новыйАдресат != "")
            {
                form = new FormКорреспонденты(новыйАдресат);
            }
            else
            {
                form = new FormКорреспонденты();
            }

            form.ShowDialog(this);
            ds11.Корреспонденты.Clear();
            DS1TableAdapters.КорреспондентыTableAdapter adapter = new RegKor.DS1TableAdapters.КорреспондентыTableAdapter();
            adapter.Fill(ds11.Корреспонденты);
            comboBoxАдресат.DataSource = null;
            comboBoxАдресат.DisplayMember = "";
            comboBoxАдресат.ValueMember = "";
            comboBoxАдресат.DataSource = ds11.Корреспонденты;
            comboBoxАдресат.DisplayMember = ds11.Корреспонденты.Columns["ОписаниеКорреспондента"].ToString();
            comboBoxАдресат.ValueMember = ds11.Корреспонденты.Columns["id_корреспондента"].ToString();
            comboBoxАдресат.Text = новыйАдресат;
            if (comboBoxАдресат.Text != новыйАдресат)
            {
                comboBoxАдресат.Text = "";
                comboBoxАдресат.SelectedText = новыйАдресат;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            СправочникАдресатов();
        }

        private void maskedTextBox1_Enter(object sender, EventArgs e)
        {
            maskedTextBox1.Select(2, 0);
        }

        private void maskedTextBox1_MouseUp(object sender, MouseEventArgs e)
        {
            maskedTextBox1.Select(2, 0);
        }

        private bool ПроверкаНомера(string номер)
        {
           // int номерКомитета;
            //char[] номерПодразделения = new char[2];

            string[] numDoc = номер.Split(' ');

            StringBuilder build = new StringBuilder();

            foreach (string s in numDoc)
            {
                if (s.Length > 0)
                {
                    build.Append(s);
                }
            }

            string numberDoc = build.ToString();

            char[] массивСимволов = numberDoc.ToCharArray();// номер.ToCharArray();// номер.Replace("", string.Empty).ToCharArray();
            int длинаНомера = массивСимволов.Length;
            string ошибка = "";
            string предупреждение = "";
            int колвоОшибок = 1;
         //   bool район = false;

          
            // проверка номера комитета
            if (длинаНомера >= 1)
            {
                номерКомитета = "" + массивСимволов[0] + массивСимволов[1];// +массивСимволов[2] + массивСимволов[3];// +массивСимволов[4];

                // Разрешим номер комитета с 01.
                //if (номерКомитета != "12-02" && номерКомитета != "12-01")
                if (номерКомитета != "02")// && номерКомитета != "12-01")
                {
                    ошибка += колвоОшибок + ". Ошибка. Номер комитета должен быть 13.\n";
                    колвоОшибок++;
                }
            }

            //проверка номера подразделения
         /*   if (длинаНомера >= 4)
            {
                if (массивСимволов[6].Equals(' '))
                {
                    ошибка += колвоОшибок + ". Ошибка. Не указан номер подразделения.\n";
                    колвоОшибок++;
                }
                else if (массивСимволов[6] < 48 || массивСимволов[6] > 49)   // первая цифра меньше нуля или больше одного
                {
                    ошибка += колвоОшибок + ". Ошибка. Номер подразделения должен быть от 01 до 12.\n";
                    колвоОшибок++;
                }
                else if (массивСимволов[6] == 48 && (массивСимволов[7] < 48 || массивСимволов[8] > 57))// первя цифра ноль, вторая отличается от [0-9]
                {
                    ошибка += колвоОшибок + ". Ошибка. Номер подразделения должен быть от 01 до 12.\n";
                    колвоОшибок++;
                }
                else if (массивСимволов[6] == 49 && (массивСимволов[7] < 48 || массивСимволов[8] > 50))// первя цифра один, вторая отличается от 0, 1 или 2
                {
                    ошибка += колвоОшибок + ". Ошибка. Номер подразделения должен быть от 01 до 12.\n";
                    колвоОшибок++;
                }
                else
                {*/
                    //номерПодразделения[0] = массивСимволов[9];
                    //номерПодразделения[1] = массивСимволов[10];

                    номерПодразделения[0] = массивСимволов[6];
                    номерПодразделения[1] = массивСимволов[7];
               // }
           /*     if (массивСимволов[7].Equals('7'))
                {
                    район = true;
                }*/
         //   }
            //проверка номера номенклатурного
            if (длинаНомера >= 7)
            {
                //if (массивСимволов[9].Equals(' ') && массивСимволов[10].Equals(' '))
                if (массивСимволов[8].Equals(' ') && массивСимволов[9].Equals(' '))
                {
                    ошибка += колвоОшибок + ". Ошибка. Не указан номер номенклатурный.\n";
                    колвоОшибок++;
                }
                // else if (массивСимволов[4] < 48 || массивСимволов[4] > 57)
                //  {
                //      ошибка += колвоОшибок + ". Ошибка. Номер номенклатурный должен быть от 01 до 99.\n";
                ///      колвоОшибок++;
                // }
                //   else if (массивСимволов[6] < 48 || массивСимволов[6] > 57)
                //   {
                //       ошибка += колвоОшибок + ". Ошибка. Номер номенклатурный должен быть от 01 до 99.\n";
                //      колвоОшибок++;
                //  }
                else
                {
                    номерНоменклатурный[0] = массивСимволов[3];
                    номерНоменклатурный[1] = массивСимволов[4];
                    //    номерНоменклатурный[2] = массивСимволов[2];
                    //   номерНоменклатурный[3] = массивСимволов[3];
                    //   номерНоменклатурный[4] = массивСимволов[4];
                }
            }
            // Проверка буквенного обозначения и номера порядкового
           /* if (район)
            {
                // Проверка буквенного обозначения
                if (длинаНомера >= 8)
                {
                    char ch = массивСимволов[11];
                    if (ch.Equals(' '))
                    {
                        ошибка += колвоОшибок + ". Ошибка. Не указана буква района.\n";
                        колвоОшибок++;
                    }
                    //else if ( !( ch.Equals( 'в' ) || ch.Equals( 'з' ) || ch.Equals( 'к' ) || ch.Equals( 'л' ) || ch.Equals( 'о' ) || ch.Equals( 'ф' ) ) )
                    else if (!(ch.Equals('ц') || ch.Equals('з') || ch.Equals('л') || ch.Equals('о') || ch.Equals('в') || ch.Equals('ф') || ch.Equals('к')))
                    {
                        ошибка += колвоОшибок + ". Ошибка. Указана неправильная буква района.\n";
                        колвоОшибок++;
                    }
                    else
                    {
                        буквенноеОбозначение[0] = ch;
                    }
                }
                // Проверка номера порядкового
                if (длинаНомера > 9)
                {
                    string номерППnew = номер.Substring(13);
                    if (Convert.ToInt32(номерППnew) > следНомерПП)
                    {

                        ошибка += колвоОшибок + ". Ошибка. Указан порядковый номер больше разрешенного. Макс. значение " + следНомерПП + ".\n";
                        колвоОшибок++;
                    }
                    else if (Convert.ToInt32(номерППnew) == 0)
                    {
                        ошибка += колвоОшибок + ". Ошибка. Порядковый номер не может быть равен 0. Рекомендуемое значение " + следНомерПП + ".\n";
                        колвоОшибок++;
                    }
                    //else
                    //{
                    //    номерПП = Convert.ToInt32( номерППnew );
                    //}
                    if (Convert.ToInt32(номерППnew) < следНомерПП)
                    {
                        предупреждение += колвоОшибок + ". Предупреждение. Номер порядковый не соответствует рекомендуемому.\nРекомендуемый номер порядковый " + следНомерПП;
                    }
                }
                else
                {
                    ошибка += колвоОшибок + ". Ошибка. Не указан номер порядковый.\nРекомендуемый номер порядковый " + следНомерПП;
                    колвоОшибок++;
                }
            }
            else
            {
                буквенноеОбозначение[0] = ' ';
                if (длинаНомера > 8)
                {
                    string номерППnew = номер.Substring(12);
                    if (Convert.ToInt32(номерППnew) > следНомерПП)
                    {
                        ошибка += колвоОшибок + ". Ошибка. Указан порядковый номер больше разрешенного. Макс. значение " + следНомерПП + ".\n";
                        колвоОшибок++;
                    }
                    else if (Convert.ToInt32(номерППnew) == 0)
                    {
                        ошибка += колвоОшибок + ". Ошибка. Порядковый номер не может быть равен 0. Рекомендуемое значение " + следНомерПП + ".\n";
                        колвоОшибок++;
                    }
                    //else
                    //{
                    //    номерПП = Convert.ToInt32( номерППnew );
                    //}
                    // Проверка на соответствие правильному порядковому номеру
                    if (Convert.ToInt32(номерППnew) < следНомерПП)
                    {
                        //if ( номерПП != Convert.ToInt32(номерППnew) )
                        //{
                        предупреждение += колвоОшибок + ". Предупреждение. Номер порядковый не соответствует рекомендуемому.\nРекомендуемый номер порядковый " + следНомерПП;
                        //}
                    }
                }
                else
                {
                    ошибка += колвоОшибок + ". Ошибка. Не указан номер порядковый.\nРекомендуемый номер порядковый " + следНомерПП;
                    колвоОшибок++;
                }
            }*/
            // Отображение итогов
            if (!ошибка.Equals(""))
            {
                errorProviderНомер.SetError(this.maskedTextBox1, (ошибка + предупреждение).Trim());
                this.DialogResult = DialogResult.None;
                return false;
            }
            else
            {
                errorProviderНомер.SetError(this.maskedTextBox1, (предупреждение).Trim());
                return true;
            }
        }

        /// <summary>
        /// Проверяет правильность заполнения элементов формы
        /// </summary>
        /// <returns></returns>
        private bool ПроверкаЗаполнения()
        {
            bool result = true;
            string ошибкаДаты = "";
            int errДаты = 1;
            if (dateTimeДата.Value > DateTime.Now)
            {
                ошибкаДаты += errДаты + " . Предупреждение - дата отправления указана в будущем времени, фантастика так сказать... А может часы врут?!\n";
                errДаты++;
            }
            errorProviderДата.SetError(dateTimeДата, ошибкаДаты.Trim());

            string ошибкаАдресата = "";
            int errАдресата = 1;
            if (comboBoxАдресат.Text == "")
            {
                ошибкаАдресата += errАдресата + " . Ошибка - не указан адресат.\n";
                errАдресата++;
                result = false;
            }
            errorProviderАдресат.SetError(comboBoxАдресат, ошибкаАдресата.Trim());

            string ошибкаСодержания = "";
            int errСодержания = 1;
            if (textBoxСодержание.Text.Trim() == "")
            {
                ошибкаСодержания += errСодержания + " . Ошибка - не указано содержание документа.\n";
                errСодержания++;
                result = false;
            }
            if (textBoxСодержание.Text.Length > 0 && textBoxСодержание.Text.Length < 5)
            {
                ошибкаСодержания += errСодержания + " . Предупреждение - слишком краткое описание, как потом будете понимать, что это за документ???.\n";
            }
            errorProviderАдресат.SetError(textBoxСодержание, ошибкаСодержания.Trim());
            //ПроверкаНомера( maskedTextBox1.Text );

            if (!ПроверкаНомера(maskedTextBox1.Text))
            {
                result = false;
            }

            return result;
        }

        /// <summary>
        /// Щелчёк мыши по кнопке СОХРАНИТЬ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonСохранить_Click(object sender, EventArgs e)
        {

            if (!ПроверкаЗаполнения())
            {
                this.DialogResult = DialogResult.None;
                return;
            }

            string[] arr = maskedTextBox1.Text.Split('/');

            #region Старый алгоритм
            //if (arr.Length != 2)
            //{
            //    MessageBox.Show(this,
            //       "Неверно указан номер исходящий",
            //       "Ошибка номера",
            //       MessageBoxButtons.OK,
            //       MessageBoxIcon.Error);
            //    this.DialogResult = DialogResult.None;
            //    return;
            //}
            //else
            //{
            //    if (Information.IsNumeric(arr[1]))
            //    {
            //        if (Convert.ToInt32(arr[1]) > следНомерПП)
            //        {
            //            MessageBox.Show(this,
            //               "Неверно указан порядковый исходящий номер. Вы можете указать число не больше чем " + следНомерПП,
            //               "Ошибка номера",
            //               MessageBoxButtons.OK,
            //               MessageBoxIcon.Error);
            //            this.DialogResult = DialogResult.None;
            //            return;

            //        }
            //        else if ((Convert.ToInt32(arr[1]) < следНомерПП && новыйДокумент) || ((Convert.ToInt32(arr[1]) < следНомерПП) && !новыйДокумент && (номерПП != Convert.ToInt32(arr[1]))))
            //        {
            //            DialogResult result = MessageBox.Show(this,
            //                "Вы указали порядковый исходящий номер, который не соответствует рекомендуемому.\nЕсли вы оставите введенный номер, возможно дублирование номеров в базе данных",
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
            //        // Устанавливаем номер порядковый
            //        строкаИсходящейКарточки["НомерПорядковый"] = arr[1];
            //    }
            //    else
            //    {
            //        MessageBox.Show(this,
            //           "Неверно указан номер исходящий",
            //           "Ошибка номера",
            //           MessageBoxButtons.OK,
            //           MessageBoxIcon.Error);
            //        this.DialogResult = DialogResult.None;
            //        return;
            //    }
            //}
            #endregion

            // Устанавливаем дату отправления
            строкаИсходящейКарточки["Дата"] = dateTimeДата.Value.ToShortDateString();

            // Устанавливаем id адресата
            DataRow[] rows = ds11.Корреспонденты.Select("ОписаниеКорреспондента='" + comboBoxАдресат.Text.Trim() + "'");
            if (rows.Length > 0)
            {
                строкаИсходящейКарточки["id_Адресата"] = comboBoxАдресат.SelectedValue;
            }
            else if (comboBoxАдресат.Text != "")
            {
                DialogResult result = MessageBox.Show(
                                                this,
                                                "Вы указали адресат, который не зарегистрирован в справочнике \"Адресаты\". Будем добавлять его в справочник или нет?",
                                                "Неизвестный адресат",
                                                MessageBoxButtons.YesNo,
                                                MessageBoxIcon.Question,
                                                MessageBoxDefaultButton.Button1
                                            );
                if (result == DialogResult.No)
                {
                    this.DialogResult = DialogResult.None;
                    return;
                }
                if (result == DialogResult.Yes)
                {
                    СправочникАдресатов();
                    this.DialogResult = DialogResult.None;
                    return;
                }
            }


            // Устанавливаем содержание
            строкаИсходящейКарточки["Содержание"] = textBoxСодержание.Text;

            // Устанавливаем номер комитета
            строкаИсходящейКарточки["НомерКомитета"] = номерКомитета; //номерКомитета;

            string querySelect = "SELECT [id_подразделения] " +
                     " FROM [ПодразделенияКомитета] " +
                     " where НомерПодразделения = '"  +номерПодразделения[0] + номерПодразделения[1]  +"' and ФлагДействующий = 'True' and Удален = 'False' ";

            DataTable tabRows = DataTableSql.GetDataTable(querySelect);

            //строкаИсходящейКарточки["id_Подразделения"] = (int)row[0]["id_Подразделения"];
            строкаИсходящейКарточки["id_Подразделения"] = (int)tabRows.Rows[0]["id_Подразделения"];
            // Устанавливаем номер подразделения
        /*    
            
            if (row.Length < 1)
            {
                MessageBox.Show("Не найдено указанное подразделение. Проверьте справочник \"Подразделения комитета\"");
                this.DialogResult = DialogResult.None;
                return;
            }
            */

            // Устанавливаем номер номенклатурный
            строкаИсходящейКарточки["НомерНоменклатурный"] = номерНоменклатурный[0] + "" + номерНоменклатурный[1];

            строкаИсходящейКарточки["FlagPersonData"] = chkFlagPersonData.Checked;

            // Установим адресат пользователя.
            Адресат = this.comboBoxАдресат.Text.Trim();

            // Списываем исходящий документ по которому дан ответ в "дело"
            if (ИДВходящегоДокумента != -1)
            {

                // Установим flag в положение true.
                FlagОтветПисьмо = true;

                // Обнулим Адресат.
                Адресат = string.Empty;

                DataRow[] списание = ds11.Карточка.Select("id_карточки=" + ИДВходящегоДокумента);
                if (списание.Length > 0)
                {
                    списание[0]["ВДело"] = true;
                    списание[0]["РезультатВыполнения"] = "Дан ответ. № исх. документа " + maskedTextBox1.Text;
                    строкаИсходящейКарточки["id_ВходящегоДокумента"] = ИДВходящегоДокумента;


                    if (chkFlagPersonData.Checked == true)
                    {
                        // Укажем что письмо ответ на персональные данные.
                        списание[0]["FlagPersonData"] = true;
                    }
                }// и отменяем списание предыдущего документа:
                if (!новыйДокумент && (ИДСтарогоВходящегоДокумента != ИДВходящегоДокумента) && (ИДСтарогоВходящегоДокумента != -1))
                {
                    DataRow[] отмена = ds11.Карточка.Select("id_карточки=" + ИДСтарогоВходящегоДокумента);
                    отмена[0]["ВДело"] = false;
                    отмена[0]["РезультатВыполнения"] = "";
                }
            }

            // Проверим если карточка создается по новой тогда сразу выходим из формы.
            if (IdКарочкаИсходящая > 0)
            {
                // Список id связующих таблиц.
                ПодключитьБД strConn = new ПодключитьБД();
                string sConn = strConn.СтрокаПодключения();

                // Выполним в единой транзакции.
                using (SqlConnection con = new SqlConnection(sConn))
                {
                    con.Open();

                    string strQueryInsert = this.QueryInsert.Replace("{0}", IdКарочкаИсходящая.ToString().Trim());

                    // Start a local transaction.
                    //SqlTransaction sqlTran = con.BeginTransaction();

                    // Запишем в таблице СвязующаяЦельПолучениперсональныхДанных id основания для передачи персональных данных.
                    //GetDataTable tabОснование = new GetDataTable(this.QueryInsert);
                    //tabОснование.DataTableSqlTransaction(con, sqlTran);

                    // ID карточки исходящей.
                    int iTestId = IdКарочкаИсходящая;

                    int id_карточкиИсход = Convert.ToInt32(строкаИсходящейКарточки["id_карточки"]);
                    string quer = "select id from СвязующаяКарточкаВходящаяИсходящая " +
                                      "where id_карточкаИсходящая = " + IdКарочкаИсходящая + " ";

                    GetDataTable tabКартВходИсход = new GetDataTable(quer);
                    DataTable tabСвязующаяКарточкаВходящаяИсходящая = tabКартВходИсход.DataTable("СвязующаяКарточкаВходящаяИсходящая");
                    //DataTable tabСвязующаяКарточкаВходящаяИсходящая = tabКартВходИсход.DataTableToConnect("СвязующаяКарточкаВходящаяИсходящая", con);

                    string query2 = "select id from dbo.СвязующаяЦельПолучениперсональныхДанных " +
                                    "where id_карточки = " + IdКарочкаИсходящая + " ";

                    GetDataTable tabСвязЦель = new GetDataTable(query2);
                    DataTable tabСвязующаяЦельПолучениперсональныхДанных = tabСвязЦель.DataTable("СвязующаяЦельПолучениперсональныхДанных");


                    if (tabСвязующаяКарточкаВходящаяИсходящая.Rows.Count > 0)
                    {
                        // Заполним список id карточки.
                        foreach (DataRow r in tabСвязующаяКарточкаВходящаяИсходящая.Rows)
                        {
                            int id = Convert.ToInt32(r[0]);
                            ListIDСвязующаяКарточкаВходящаяИсходящая.Add(id);
                        }
                    }

                    if (tabСвязующаяЦельПолучениперсональныхДанных.Rows.Count > 0)
                    {
                        // Заполним список id карточки.
                        foreach (DataRow r in tabСвязующаяЦельПолучениперсональныхДанных.Rows)
                        {
                            int id = Convert.ToInt32(r[0]);
                            ListIDСвязующаяЦельПолучениперсональныхДанных.Add(id);
                        }
                    }
                }
            }

            // Префикс номера
            ПрефиксНомерИсходящий = maskedTextBox1.Text.Trim().Split('/')[0];

            // Запишем номер исходящего документа.
            НомерИсходящий = "Дан ответ. № исх. документа " + maskedTextBox1.Text.Trim();

            // Передадим в свойство GUID и будем использовать его для идентификации записи при прочтении номер документа.
            this.StrGuid = Guid.NewGuid().ToString().Trim();

            // Выведим окно которое содержит список видов поступления документов.
            FormTypeCompanyDocument formType = new FormTypeCompanyDocument();
            formType.ShowDialog();

            if (formType.DialogResult == DialogResult.OK)
            {
                // Передадим в форму способ поступления документа.
                СпособПоступления = formType.СпособПоступления;
            }

            if (this.FlagNumStopDoc == true)
            {
                
                this.NumDocNoAutomat = Convert.ToInt32(this.maskedTextBox1.Text.Split('/')[1]);
            }

            this.Close();
        }

        private void maskedTextBox1_MaskInputRejected(object sender, MaskInputRejectedEventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            FormPD formPD = new FormPD();
            formPD.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FormОснованиеПередачи formОснование = new FormОснованиеПередачи();

            // Если родительская форма раьботает в режиме изменения записи.
            if (this.FlagEdit == true)
            {
                // Отобразить основания для передачи персональных данных.
                formОснование.FlagEdit = true;
            }

            // Установим флаг оснований в положение запись.
            //FlagListBase = "ВыборОснований";

            //if (FlagListBase == "ВыборОснований")
            //{
            //    this.buttonСохранить.Enabled = false;
            //}

            // Обнулим список ОснованиеПередачи перед использованием.
            formОснование.ListОснованиеПередачи.Clear();

            // Передадим в форму id карточки.
            formОснование.IdКарточки = this.IdКарочкаИсходящая;

            formОснование.ShowDialog();

            if (formОснование.DialogResult == DialogResult.OK)
            {
                // Обнулим свойство.
                this.QueryInsert = string.Empty;

                // Передадим SQL инструкцию.
                this.QueryInsert = formОснование.StringQuery;

                this.ListОснованиеПередачи = formОснование.ListОснованиеПередачи;


                //List<ОснованиеПередачи> ListOP = new List<ОснованиеПередачи>();
                //ListOP.Clear();

                //foreach (ОснованиеПередачи item in formОснование.ListОснованиеПередачи)
                //{
                //    ListOP.Add(item);
                //}
                    

                //// Проверим список на наличие выбранных записей.
                //if (ListOP.Count > 0)
                //{
                //    FlagListBase = "СписокОснование";
                //}
                //else
                //{
                //    this.label2.Text = "Основания передачи не выбраны!!!";
                //    this.label2.ForeColor = Color.Red;
                //    this.label2.Size = new System.Drawing.Size(0, 18);
                //}

                //// Обнулим список.
                //ListОснованиеПередачи.Clear();

                //// Добавим новый список в свойства формы карточка исходящая.
                //foreach(ОснованиеПередачи item in ListOP)
                //{
                //    ListОснованиеПередачи.Add(item);
                //}

                //if (FlagListBase == "СписокОснование")
                //{
                //    this.buttonСохранить.Enabled = false;

                //    // Список для хранения данных.
                //    StringBuilder build = new StringBuilder();
                //    this.label2.Text = "";
                //    this.label2.ForeColor = Color.Green;
                //    this.label2.Size = new System.Drawing.Size(0, 13);

                //    this.buttonСохранить.Enabled = true;
                   
                //}
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ПодключитьБД connDb = new ПодключитьБД();

            using (SqlConnection connection = new SqlConnection(connDb.СтрокаПодключения()))
            {
                connection.Open();
                string sql = "select FileData from КарточкаИсходящаяДокументы where id_карточки = " + this.IdКарочкаИсходящая + "";
                SqlCommand command = new SqlCommand(sql, connection);
                //SqlDataReader reader = command.ExecuteReader();

                // Массив битов данных из БД.
                byte[] fileByteArray = (byte[])command.ExecuteScalar();

                string dir = @"d:\Recor";

                string fileName = dir + @"\TempViewOutputDoc.zip";

                FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                BinaryWriter binWriter = new BinaryWriter(fileStream);
                binWriter.Write(fileByteArray);
                binWriter.Close();

                // Откроем архив.
                System.Diagnostics.Process.Start(fileName);
            }

            //// Получим имя файла на сервере.
            //string имяФайлНаСервере = this.ИмяДокумента.Trim(); //+".zip";

            //// Путь к сервреру.
            //string путКСерверу = patchServerSave.Trim();

            //// Получим путь к файлу на серврер.
            //string fileServer = путКСерверу + @"\" + имяФайлНаСервере.Replace("/", "-");

            //// Путь к файлу во временной папке.
            //string tempPath = Path.GetTempPath();

            //// Получим путь и имя файла которое он будет иметь после копирования во временную таблицу на клиент.
            //string fileTo = tempPath + имяФайлНаСервере;

            //// Скопируем архив во временную папку.
            //File.Copy(fileServer, fileTo, true);

            //// Получим путь к файлу во временной таблице на клиенте.
            //string fileTemp = tempPath + @"\" + имяФайлНаСервере;

            //// Откроем архив.
            //System.Diagnostics.Process.Start(fileTemp);

        }

        private void FormКарточкаИсходящая_FormClosing(object sender, FormClosingEventArgs e)
        {
            string dir = @"d:\Recor";

            DirectoryInfo dirInf = new DirectoryInfo(dir);

            string sTest = dirInf.FullName;

            foreach (FileInfo fi in dirInf.GetFiles())
            {
                if (fi.Name.Trim().ToLower() == "TempViewOutputDoc.zip".Trim().ToLower())
                {
                    fi.Delete();
                }
            }
        }

        private void FormКарточкаИсходящая_FormClosed(object sender, FormClosedEventArgs e)
        {
            string dir = @"d:\Recor";

            DirectoryInfo dirInf = new DirectoryInfo(dir);

            string sTest = dirInf.FullName;

            foreach (FileInfo fi in dirInf.GetFiles())
            {
                if (fi.Name.Trim().ToLower() == "TempViewOutputDoc.zip".Trim().ToLower())
                {
                    fi.Delete();
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ПодключитьБД connDb = new ПодключитьБД();

            using (SqlConnection connection = new SqlConnection(connDb.СтрокаПодключения()))
            {
                connection.Open();
                string sql = "select FileDateTitlePage from КарточкаИсходящаяДокументы where id_карточки = " + this.IdКарочкаИсходящая + "";
                SqlCommand command = new SqlCommand(sql, connection);
                //SqlDataReader reader = command.ExecuteReader();

                // Массив битов данных из БД.
                byte[] fileByteArray = (byte[])command.ExecuteScalar();

                string dir = @"d:\Recor";

                string fileName = dir + @"\TempViewOutputDoc.zip";

                FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                BinaryWriter binWriter = new BinaryWriter(fileStream);
                binWriter.Write(fileByteArray);
                binWriter.Close();

                // Откроем архив.
                System.Diagnostics.Process.Start(fileName);
            }
        }

        private void chkBoxDsp_CheckedChanged(object sender, EventArgs e)
        {
            string stest2 = sNumStart;
            
            if (this.chkBoxDsp.Checked == true)
            {
                this.maskedTextBox1.Mask = maskedTextBox1.Mask = @"00-00-00\/09999-ДСП";
                this.maskedTextBox1.Text += stest2 + "-ДСП";
                this.FlagDsp = "True";
            }
            else
            {


                string stest = sNumStart;
                this.maskedTextBox1.Mask = @"00-00-00\/09999";
                this.maskedTextBox1.Text.Replace("-ДСП", string.Empty);
                this.FlagDsp = "False";
            }
        }
       
    }
}