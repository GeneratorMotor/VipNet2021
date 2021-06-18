using System;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
//using CrystalDecisions.ReportAppServer.
using System.Globalization;
using System.Data.SqlClient;
using System.Configuration;
using System.Text;
using System.Collections.Generic;
using System.IO;

using RegKor.Classess;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace RegKor
{
    /// <summary>
    /// Summary description for FormГлавная.
    /// </summary>
    public class FormГлавная : System.Windows.Forms.Form
    {
        #region Переменные

        private System.Windows.Forms.DataGridTableStyle dataGridTableStyle2;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumnДокумент;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumnКорреспондент;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumnДатаОтправ;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumnДатаПоступ;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumnНомерИсход;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumnНомерВход;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumnСодержание;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumnКонтроль;
        private System.Windows.Forms.DataGridBoolColumn dataGridBoolColumnВДеле;
        private System.Windows.Forms.DataGrid dataGridРабочиеДокументы;
        private System.Windows.Forms.DataGrid dataGridДокументыВДеле;
        private System.Windows.Forms.DataGrid dataGridИсходящиеДокументы;
        private System.Windows.Forms.DataGridTableStyle dataGridTableStyleРабочиеДокументы;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn1;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn2;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn3;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn4;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn5;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn6;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn7;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn8;
        private System.Windows.Forms.DataGridTableStyle dataGridTableStyleДокументыВДеле;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn9;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn10;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn11;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn12;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn13;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn14;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn15;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn16;
        private System.Windows.Forms.Panel panel1Tab1;
        private System.Windows.Forms.Panel panel4Tab1;
        private System.Windows.Forms.Panel panel1Tab2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Panel panel5;
        private System.Windows.Forms.Panel panel6;
        private System.Windows.Forms.Panel panel7;
        private System.Windows.Forms.RichTextBox labelИнфоTab1;
        private System.Windows.Forms.RichTextBox labelИнфоTab2;
        private System.Windows.Forms.TextBox textBoxСтрокаПоискаTab2;
        private System.Windows.Forms.TextBox textBoxСтрокаПоискаTab1;
        private System.Windows.Forms.Button buttonОчиститьСтрокуПоискаTab1;
        private System.Windows.Forms.Button buttonОчиститьСтрокуПоискаTab2;
        private System.Windows.Forms.Label labelОтобраноДокументовПоискомTab2;
        private System.Windows.Forms.Label labelОтобраноДокументовПоискомTab1;
        private System.Windows.Forms.MainMenu mainMenu1;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem menuItem2;
        private System.Windows.Forms.MenuItem menuItem3;
        private System.Windows.Forms.MenuItem menuItem4;
        private System.Windows.Forms.MenuItem menuItemСохранитьВФайл;
        private System.Windows.Forms.MenuItem menuItemContextПечатьКарточки;
        private System.Windows.Forms.MenuItem menuItemСправочникиДокументы;
        private System.Windows.Forms.MenuItem menuItemСправочникиКорреспонденты;
        private System.Windows.Forms.MenuItem menuItemСправочникиПолучатели;
        private System.Windows.Forms.MenuItem menuItemКонтрольныеУведомления;
        private System.Windows.Forms.MenuItem menuItemПросрочДокументы;
        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.CheckBox checkBoxKontrolFilter;
        private System.ComponentModel.IContainer components;
        System.Windows.Forms.TreeNode treeNodeЯнварь;
        System.Windows.Forms.TreeNode treeNodeФевраль;
        System.Windows.Forms.TreeNode treeNodeМарт;
        System.Windows.Forms.TreeNode treeNodeАпрель;
        System.Windows.Forms.TreeNode treeNodeМай;
        System.Windows.Forms.TreeNode treeNodeИюнь;
        System.Windows.Forms.TreeNode treeNodeИюль;
        System.Windows.Forms.TreeNode treeNodeАвгуст;
        System.Windows.Forms.TreeNode treeNodeСентябрь;
        System.Windows.Forms.TreeNode treeNodeОктябрь;
        System.Windows.Forms.TreeNode treeNodeНоябрь;
        System.Windows.Forms.TreeNode treeNodeДекабрь;
        System.Windows.Forms.TreeNode treeNodeГод;
        /// <summary>
        /// Таб-контрол со входящими документами
        /// </summary>
        private System.Windows.Forms.TabControl tabControlВходящиеДокументы;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;

        /// <summary>
        /// Таб-контрол с типами документов
        /// </summary>
        private System.Windows.Forms.TabControl tabControlТипыДокументов;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.TabPage tabPage3;

        /// <summary>
        /// Представление для "рабочих документов"
        /// </summary>
        private System.Data.DataView dataViewВыборкаРабДокументы;

        /// <summary>
        /// Представление для "документов в деле"
        /// </summary>
        private System.Data.DataView dataViewВыборкаДокументыВДеле;

        /// <summary>
        /// Представление для исходящих документов
        /// </summary>
        private System.Data.DataView dataViewИсходящиеДокументы;

        /// <summary>
        /// Считывает параметры конфигурации из файла App.config
        /// </summary>
        System.Configuration.AppSettingsReader configReader;

        /// <summary>
        /// Строка подключения к источнику данных
        /// </summary>
        string строкаПодключения = "";
        
        /// <summary>
        /// Массив временного интервала разделенных "-"
        /// </summary>
        string[] TimeInterval;
        
        /// <summary>
        /// Подключение к источнику данных
        /// </summary>
        SqlConnection подключение;

        /// <summary>
        /// Осуществляет операции обновления, вставки, удаления и выборки над источником данных
        /// </summary>
        SqlDataAdapter датаАдаптер;

        /// <summary>
        /// Строка с именем источника данных
        /// </summary>
        string источникДанных;

        /// <summary>
        /// Глобальный объект для проверка на pegecr второй копии программы
        /// </summary>
        static System.Threading.Mutex mutex;

        /// <summary>
        /// Запускает окно с рисунком ожидания
        /// </summary>
        public System.Threading.Thread потокОжидания;

        private DataGridTableStyle dataGridTableStyleИсходящиеДокументы;
        private DataGridTextBoxColumn dataGridTextBoxColumnИсхДокДатаИсхода;
        private DataGridTextBoxColumn dataGridTextBoxColumnИсхДокНомер;
        private DataGridTextBoxColumn dataGridTextBoxColumnИсхДокОписаниеАдресата;
        private DataGridTextBoxColumn dataGridTextBoxColumnИсхДокСодержание;
        private DataGridTextBoxColumn dataGridTextBoxColumnИсхДокНомерВходДокта;
        private SplitContainer splitContainer1;
        private SplitContainer splitContainer2;
        private RichTextBox labelИнфоTab3;
        private TextBox textBoxСтрокаПоискаИсходящихДокументов;
        private Label labelОтобраноДокументовПоискомИсходящихДокументов;
        private Button buttonОчиститьСтрокуПоискаИсходящихДокументов;
        private MenuItem menuItemСправочникиПодразделения;
        private DS1 ds11;
        private SplitContainer splitContainer3;
        private TableLayoutPanel tableLayoutPanel2;
        private CheckBox checkBox1;
        private CheckBox checkBox2;
        private Label label2;
        private MenuItem menuItem5;
        private ComboBox comboBoxФильтрИДПоДате;
        private CheckBox checkBoxКорреспонденты;
        private ComboBox comboBoxКорреспонденты;
        private MenuItem menuItem6;
        private MenuItem menuItem7;
        private MenuItem menuItem8;
        private MenuItem menuItem9;

        //Хранит выбранный год
        private int selectedYear;
        private string выбраннаяДата;
        
        private string выбранныйГод;
        private string следующаяДата;
        private MenuItem menuItem10;
        private MenuItem menuItem11;
        private MenuItem menuItem12;
        private MenuItem menuItem13;

        // Переменные для хранения имени файла который копируется на сервер.
        private string fileName = string.Empty;
        private string fileNameCopy = string.Empty;

        // Флаг указывает, что список с просроченными документами отпечатывается 1-ый раз.
        private bool flagFirstLoad = false;

        // Переменная для хранения префикас карточки исходящей.
        private string numberPrefix = string.Empty;

        /// <summary>
        /// Представляет координаты мыши
        /// </summary>
        private struct координатыМыши
        {
            public int X;
            public int Y;
        }

        /// <summary>
        /// Структура с координатами мыши
        /// </summary>
        координатыМыши мышь = new координатыМыши();
        private MenuItem menuItem14;
        private MenuItem menuItem15;

        /// <summary>
        /// Путь с указанием папки копирования на сервер.
        /// </summary>
        private string patchServerFile = string.Empty;

        /// <summary>
        /// Флаг указывает, что документ помечен для записи на сервер.
        /// </summary>
        private bool flagInsertCopyDoc = false;

        #endregion
        private MenuItem menuItem16;
        private MenuItem menuItem17;
        private MenuItem menuItem18;
        private MenuItem menuItem19;
        private MenuItem menuItem20;
        private MenuItem menuItem21;
        private MenuItem menuItem24;
        private MenuItem menuItem22;
        private MenuItem menuItem23;
        private MenuItem menuItem25;
        private MenuItem menuItem26;
        private MenuItem menuItem27;
        private MenuItem menuItem28;
        private MenuItem menuItem29;
        private MenuItem menuItem30;
        private MenuItem menuItem31;
        private MenuItem menuItem32;
        private MenuItem menuItem33;

        private List<PersonRecepient> listPerson;

        /// <summary>
        /// Свойство хранит список начальников отделов и управлений которые выбраны для получения отписанного документа.
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



        /// <summary>
        /// Конструктор формы
        /// </summary>
        public FormГлавная()
        {
            InitializeComponent();

            // считываем путь к источнику данных:
            configReader = new AppSettingsReader();
            строкаПодключения = (string)configReader.GetValue("строкаПодключения", typeof(System.String));
            //Временной период выборки из базы данных 
            TimeInterval = configReader.GetValue("ВременнойИнтервал", typeof(System.String)).ToString().Split('-');

            //Выберим год
            SelectYearForm selectYearForm = new SelectYearForm();
            selectYearForm.ShowDialog();

            //если пользователь нажал ОК запоминаем выбранный год
            if (selectYearForm.DialogResult == DialogResult.OK)
            {
                //получим выбранный год
                selectedYear = selectYearForm.SelectedYear;
                выбраннаяДата = ДатаРаботыПрограммы.ДатаНастройкиПрограммы(selectedYear);

                //получим год с которого идёт фильтрация данных работает программа
                выбранныйГод = ДатаРаботыПрограммы.ВыбранныйГод(selectedYear); //selectedYear.ToString();

                //получим следующую дату
                следующаяДата = ДатаРаботыПрограммы.ДатаСледующийГод(selectedYear);
            }

            //если пользователь нажал Отмена выходим из приложения
            if (selectYearForm.DialogResult == DialogResult.Cancel)
            {
                this.Close();
                Environment.Exit(0);
            }
            

            // создаем подключение к источнику данных:
            подключение = new SqlConnection(строкаПодключения);

            // создаем датаадаптер для операций над источником данных:
            датаАдаптер = new SqlDataAdapter("", подключение);

            источникДанных = подключение.DataSource.ToString();

            string str = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Environment.CurrentDirectory + "\\RegKor.exe").FileVersion;

            this.Text = "Регистрация корреспонденции. Версия: " + str + ". SQL Server: " + подключение.DataSource;

            // заполняем данными датасет и отображаем их на форме:
            ПодключитьсяПолучитьДанные();

            // Проверяем, есть ли просроченные документы
            //DataRow[] rows = ds11.Выборка.Select("СрокВыполнения<'" + DateTime.Now.ToString() + "' AND ДатаПоступ >='01.12.2011' AND НаКонтроле=True AND ВДело=False");

            List<ПросроченныеДокументы> list = new List<ПросроченныеДокументы>();

            string sTest = DateTime.Now.ToString();

            string querySelect = "SELECT * FROM [Выборка] " +
                                 "where СрокВыполнения<'"+ ДатаSQL.Дата(DateTime.Today.ToShortDateString()) +"' AND ДатаПоступ >= '"+ выбранныйГод +"0112' AND НаКонтроле='True' AND ВДело='False'";

            GetDataTable getTable = new GetDataTable(querySelect);
            DataTable tab = getTable.DataTable("Выборка");

            int iCount = 1;

            // Заполним данными список.
            foreach (DataRow row in tab.Rows)
            {
                ПросроченныеДокументы item = new ПросроченныеДокументы();
                item.НомерПП = iCount.ToString().Trim();
                item.ОтветственныйИсполнитель = row["Резолюция"].ToString().Trim();
                item.ДатаПоступления = Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString();
                item.НомерВходящий = row["НомерВход"].ToString().Trim();
                item.СрокВыполнения = Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString();

                list.Add(item);

                iCount++;
            }



            // Добавим записи из представления ВыборкаПовтор.
            string querySelectP = "SELECT * FROM [ВыборкаПовтор] " +
                                "where СрокВыполнения <'" + ДатаSQL.Дата(DateTime.Today.ToShortDateString()) + "' AND ДатаПоступ >= '" + выбранныйГод + "0112' AND НаКонтроле='True' AND ВДело='False'";

            // Получим данные из представления ВыборкаПовтор.
            GetDataTable getTableP = new GetDataTable(querySelectP);
            DataTable tabP = getTableP.DataTable("ВыборкаПовтор");

            // Заполним данными список.
            foreach (DataRow row in tabP.Rows)
            {
                ПросроченныеДокументы item = new ПросроченныеДокументы();
                item.НомерПП = iCount.ToString().Trim();
                item.ОтветственныйИсполнитель = row["Резолюция"].ToString().Trim();
                item.ДатаПоступления = Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString();
                item.НомерВходящий = row["НомерВход"].ToString().Trim();
                //item.СрокВыполнения = Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString();
                item.СрокВыполнения = Convert.ToDateTime(row["СрокВыполнения"]).AddDays(2).ToShortDateString();

                list.Add(item);

                iCount++;
            }

            int ii = list.Count;


            //if (tab.Rows.Count > 0)
            if (list.Count > 0)
            {
                if (flagFirstLoad == false)
                {
                    //ПечатьПросроченныхДокументов(list); -- Оставим пока вдруг понадобится.
                    flagFirstLoad = true;
                }
            }

            // Стараый рабочий код.
            //DataRow[] rows = ds11.Выборка.Select("СрокВыполнения<'" + DateTime.Now.ToShortDateString() + "' AND ДатаПоступ >= '01.12." + выбранныйГод + "' AND НаКонтроле=True AND ВДело=False");
            //if (rows.Length > 0)
            //{
            //    ПечатьПросроченныхДокументов();
            //}

            //Фильтр конец года
            comboBoxФильтрИДПоДате.SelectedItem = "Весь год";
            string фильтр2 = "ВДело=False AND ДатаПоступ >='01.12." + выбранныйГод + "'";
            dataViewВыборкаРабДокументы.RowFilter = фильтр2;
            dataGridРабочиеДокументы.DataSource = dataViewВыборкаРабДокументы;

            string фильтр3 = "ВДело=True AND ДатаПоступ >='01.12." + выбранныйГод + "'";
            dataViewВыборкаДокументыВДеле.RowFilter = фильтр3;
            dataGridДокументыВДеле.DataSource = dataViewВыборкаДокументыВДеле;

            // Получим путь для копирования файла.
            string queryPatchServer = "select top 1 PatchServer from СерверПуть";

            // Получим данные из о пути копирования файла.
            GetDataTable getTablePatch = new GetDataTable(queryPatchServer);
            DataTable tabPatch = getTablePatch.DataTable("СерврПуть");

            // Запишем в переменную формы путь коприрования файлов.
            patchServerFile = tabPatch.Rows[0]["PatchServer"].ToString().Trim();

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormГлавная));
            this.dataViewВыборкаРабДокументы = new System.Data.DataView();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.panel2 = new System.Windows.Forms.Panel();
            this.labelОтобраноДокументовПоискомTab1 = new System.Windows.Forms.Label();
            this.buttonОчиститьСтрокуПоискаTab1 = new System.Windows.Forms.Button();
            this.checkBoxKontrolFilter = new System.Windows.Forms.CheckBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.textBoxСтрокаПоискаTab1 = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.textBoxСтрокаПоискаTab2 = new System.Windows.Forms.TextBox();
            this.buttonОчиститьСтрокуПоискаTab2 = new System.Windows.Forms.Button();
            this.textBoxСтрокаПоискаИсходящихДокументов = new System.Windows.Forms.TextBox();
            this.buttonОчиститьСтрокуПоискаИсходящихДокументов = new System.Windows.Forms.Button();
            this.panel4Tab1 = new System.Windows.Forms.Panel();
            this.dataGridРабочиеДокументы = new System.Windows.Forms.DataGrid();
            this.dataGridTableStyleРабочиеДокументы = new System.Windows.Forms.DataGridTableStyle();
            this.dataGridTextBoxColumn1 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn2 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn3 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn4 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn5 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn6 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn7 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn8 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.mainMenu1 = new System.Windows.Forms.MainMenu(this.components);
            this.menuItem1 = new System.Windows.Forms.MenuItem();
            this.menuItem8 = new System.Windows.Forms.MenuItem();
            this.menuItem9 = new System.Windows.Forms.MenuItem();
            this.menuItemСохранитьВФайл = new System.Windows.Forms.MenuItem();
            this.menuItem2 = new System.Windows.Forms.MenuItem();
            this.menuItemСправочникиКорреспонденты = new System.Windows.Forms.MenuItem();
            this.menuItemСправочникиПодразделения = new System.Windows.Forms.MenuItem();
            this.menuItemСправочникиПолучатели = new System.Windows.Forms.MenuItem();
            this.menuItemСправочникиДокументы = new System.Windows.Forms.MenuItem();
            this.menuItem11 = new System.Windows.Forms.MenuItem();
            this.menuItem12 = new System.Windows.Forms.MenuItem();
            this.menuItem31 = new System.Windows.Forms.MenuItem();
            this.menuItem20 = new System.Windows.Forms.MenuItem();
            this.menuItem21 = new System.Windows.Forms.MenuItem();
            this.menuItem24 = new System.Windows.Forms.MenuItem();
            this.menuItem25 = new System.Windows.Forms.MenuItem();
            this.menuItem22 = new System.Windows.Forms.MenuItem();
            this.menuItem26 = new System.Windows.Forms.MenuItem();
            this.menuItem27 = new System.Windows.Forms.MenuItem();
            this.menuItem28 = new System.Windows.Forms.MenuItem();
            this.menuItem23 = new System.Windows.Forms.MenuItem();
            this.menuItem29 = new System.Windows.Forms.MenuItem();
            this.menuItem30 = new System.Windows.Forms.MenuItem();
            this.menuItem32 = new System.Windows.Forms.MenuItem();
            this.menuItem17 = new System.Windows.Forms.MenuItem();
            this.menuItem18 = new System.Windows.Forms.MenuItem();
            this.menuItem19 = new System.Windows.Forms.MenuItem();
            this.menuItem3 = new System.Windows.Forms.MenuItem();
            this.menuItemПросрочДокументы = new System.Windows.Forms.MenuItem();
            this.menuItemКонтрольныеУведомления = new System.Windows.Forms.MenuItem();
            this.menuItemContextПечатьКарточки = new System.Windows.Forms.MenuItem();
            this.menuItem5 = new System.Windows.Forms.MenuItem();
            this.menuItem4 = new System.Windows.Forms.MenuItem();
            this.menuItem7 = new System.Windows.Forms.MenuItem();
            this.menuItem6 = new System.Windows.Forms.MenuItem();
            this.menuItem10 = new System.Windows.Forms.MenuItem();
            this.menuItem13 = new System.Windows.Forms.MenuItem();
            this.menuItem14 = new System.Windows.Forms.MenuItem();
            this.menuItem15 = new System.Windows.Forms.MenuItem();
            this.menuItem16 = new System.Windows.Forms.MenuItem();
            this.tabControlВходящиеДокументы = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel1Tab1 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.labelИнфоTab1 = new System.Windows.Forms.RichTextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel1Tab2 = new System.Windows.Forms.Panel();
            this.dataGridДокументыВДеле = new System.Windows.Forms.DataGrid();
            this.dataGridTableStyleДокументыВДеле = new System.Windows.Forms.DataGridTableStyle();
            this.dataGridTextBoxColumn9 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn10 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn11 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn12 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn13 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn14 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn15 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn16 = new System.Windows.Forms.DataGridTextBoxColumn();
            this.panel5 = new System.Windows.Forms.Panel();
            this.panel7 = new System.Windows.Forms.Panel();
            this.labelИнфоTab2 = new System.Windows.Forms.RichTextBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.checkBoxКорреспонденты = new System.Windows.Forms.CheckBox();
            this.comboBoxКорреспонденты = new System.Windows.Forms.ComboBox();
            this.labelОтобраноДокументовПоискомTab2 = new System.Windows.Forms.Label();
            this.panel6 = new System.Windows.Forms.Panel();
            this.dataGridTableStyle2 = new System.Windows.Forms.DataGridTableStyle();
            this.dataGridTextBoxColumnДокумент = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnКорреспондент = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnДатаОтправ = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnДатаПоступ = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnНомерИсход = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnНомерВход = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnСодержание = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnКонтроль = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridBoolColumnВДеле = new System.Windows.Forms.DataGridBoolColumn();
            this.dataViewВыборкаДокументыВДеле = new System.Data.DataView();
            this.tabControlТипыДокументов = new System.Windows.Forms.TabControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.splitContainer3 = new System.Windows.Forms.SplitContainer();
            this.dataGridИсходящиеДокументы = new System.Windows.Forms.DataGrid();
            this.dataGridTableStyleИсходящиеДокументы = new System.Windows.Forms.DataGridTableStyle();
            this.dataGridTextBoxColumnИсхДокДатаИсхода = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnИсхДокНомер = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnИсхДокОписаниеАдресата = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnИсхДокСодержание = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumnИсхДокНомерВходДокта = new System.Windows.Forms.DataGridTextBoxColumn();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.comboBoxФильтрИДПоДате = new System.Windows.Forms.ComboBox();
            this.labelОтобраноДокументовПоискомИсходящихДокументов = new System.Windows.Forms.Label();
            this.labelИнфоTab3 = new System.Windows.Forms.RichTextBox();
            this.dataViewИсходящиеДокументы = new System.Data.DataView();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.ds11 = new RegKor.DS1();
            this.menuItem33 = new System.Windows.Forms.MenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dataViewВыборкаРабДокументы)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4Tab1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridРабочиеДокументы)).BeginInit();
            this.tabControlВходящиеДокументы.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel1Tab1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel1Tab2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridДокументыВДеле)).BeginInit();
            this.panel5.SuspendLayout();
            this.panel7.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataViewВыборкаДокументыВДеле)).BeginInit();
            this.tabControlТипыДокументов.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.splitContainer3.Panel1.SuspendLayout();
            this.splitContainer3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridИсходящиеДокументы)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataViewИсходящиеДокументы)).BeginInit();
            this.tableLayoutPanel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).BeginInit();
            this.SuspendLayout();
            // 
            // contextMenu1
            // 
            this.contextMenu1.Popup += new System.EventHandler(this.contextMenu1_Popup);
            // 
            // panel2
            // 
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel2.Controls.Add(this.labelОтобраноДокументовПоискомTab1);
            this.panel2.Controls.Add(this.buttonОчиститьСтрокуПоискаTab1);
            this.panel2.Controls.Add(this.checkBoxKontrolFilter);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(280, 110);
            this.panel2.TabIndex = 0;
            // 
            // labelОтобраноДокументовПоискомTab1
            // 
            this.labelОтобраноДокументовПоискомTab1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.labelОтобраноДокументовПоискомTab1.Location = new System.Drawing.Point(0, 86);
            this.labelОтобраноДокументовПоискомTab1.Name = "labelОтобраноДокументовПоискомTab1";
            this.labelОтобраноДокументовПоискомTab1.Size = new System.Drawing.Size(276, 20);
            this.labelОтобраноДокументовПоискомTab1.TabIndex = 5;
            // 
            // buttonОчиститьСтрокуПоискаTab1
            // 
            this.buttonОчиститьСтрокуПоискаTab1.Location = new System.Drawing.Point(208, 20);
            this.buttonОчиститьСтрокуПоискаTab1.Name = "buttonОчиститьСтрокуПоискаTab1";
            this.buttonОчиститьСтрокуПоискаTab1.Size = new System.Drawing.Size(64, 22);
            this.buttonОчиститьСтрокуПоискаTab1.TabIndex = 4;
            this.buttonОчиститьСтрокуПоискаTab1.Text = "Очистить";
            this.toolTip1.SetToolTip(this.buttonОчиститьСтрокуПоискаTab1, "Очистить условия поиска.");
            this.buttonОчиститьСтрокуПоискаTab1.Click += new System.EventHandler(this.buttonОчиститьСтрокуПоискаTab1_Click);
            // 
            // checkBoxKontrolFilter
            // 
            this.checkBoxKontrolFilter.Location = new System.Drawing.Point(8, 28);
            this.checkBoxKontrolFilter.Name = "checkBoxKontrolFilter";
            this.checkBoxKontrolFilter.Size = new System.Drawing.Size(160, 27);
            this.checkBoxKontrolFilter.TabIndex = 3;
            this.checkBoxKontrolFilter.Text = "Только на контроле";
            this.toolTip1.SetToolTip(this.checkBoxKontrolFilter, "Фильтр - только подконтрольные документы.");
            this.checkBoxKontrolFilter.CheckedChanged += new System.EventHandler(this.checkBoxKontrolFilter_CheckedChanged);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.textBoxСтрокаПоискаTab1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(276, 20);
            this.panel3.TabIndex = 2;
            // 
            // textBoxСтрокаПоискаTab1
            // 
            this.textBoxСтрокаПоискаTab1.Dock = System.Windows.Forms.DockStyle.Top;
            this.textBoxСтрокаПоискаTab1.Location = new System.Drawing.Point(0, 0);
            this.textBoxСтрокаПоискаTab1.Name = "textBoxСтрокаПоискаTab1";
            this.textBoxСтрокаПоискаTab1.Size = new System.Drawing.Size(276, 21);
            this.textBoxСтрокаПоискаTab1.TabIndex = 0;
            this.toolTip1.SetToolTip(this.textBoxСтрокаПоискаTab1, "Введите текст для поиска.");
            this.textBoxСтрокаПоискаTab1.TextChanged += new System.EventHandler(this.textBoxСтрокаПоиска_TextChanged);
            // 
            // textBoxСтрокаПоискаTab2
            // 
            this.textBoxСтрокаПоискаTab2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBoxСтрокаПоискаTab2.Location = new System.Drawing.Point(0, 0);
            this.textBoxСтрокаПоискаTab2.Name = "textBoxСтрокаПоискаTab2";
            this.textBoxСтрокаПоискаTab2.Size = new System.Drawing.Size(276, 21);
            this.textBoxСтрокаПоискаTab2.TabIndex = 0;
            this.toolTip1.SetToolTip(this.textBoxСтрокаПоискаTab2, "Введите текст для поиска.");
            this.textBoxСтрокаПоискаTab2.TextChanged += new System.EventHandler(this.textBoxСтрокаПоискаTab2_TextChanged);
            // 
            // buttonОчиститьСтрокуПоискаTab2
            // 
            this.buttonОчиститьСтрокуПоискаTab2.Location = new System.Drawing.Point(208, 51);
            this.buttonОчиститьСтрокуПоискаTab2.Name = "buttonОчиститьСтрокуПоискаTab2";
            this.buttonОчиститьСтрокуПоискаTab2.Size = new System.Drawing.Size(64, 22);
            this.buttonОчиститьСтрокуПоискаTab2.TabIndex = 3;
            this.buttonОчиститьСтрокуПоискаTab2.Text = "Очистить";
            this.toolTip1.SetToolTip(this.buttonОчиститьСтрокуПоискаTab2, "Очистить условия поиска.");
            this.buttonОчиститьСтрокуПоискаTab2.Click += new System.EventHandler(this.buttonОчиститьСтрокуПоискаTab2_Click_1);
            // 
            // textBoxСтрокаПоискаИсходящихДокументов
            // 
            this.textBoxСтрокаПоискаИсходящихДокументов.Dock = System.Windows.Forms.DockStyle.Top;
            this.textBoxСтрокаПоискаИсходящихДокументов.Location = new System.Drawing.Point(0, 0);
            this.textBoxСтрокаПоискаИсходящихДокументов.Name = "textBoxСтрокаПоискаИсходящихДокументов";
            this.textBoxСтрокаПоискаИсходящихДокументов.Size = new System.Drawing.Size(267, 21);
            this.textBoxСтрокаПоискаИсходящихДокументов.TabIndex = 1;
            this.toolTip1.SetToolTip(this.textBoxСтрокаПоискаИсходящихДокументов, "Введите текст для поиска.");
            this.textBoxСтрокаПоискаИсходящихДокументов.TextChanged += new System.EventHandler(this.textBoxСтрокаПоискаИсходящихДокументов_TextChanged);
            // 
            // buttonОчиститьСтрокуПоискаИсходящихДокументов
            // 
            this.buttonОчиститьСтрокуПоискаИсходящихДокументов.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonОчиститьСтрокуПоискаИсходящихДокументов.Location = new System.Drawing.Point(200, 27);
            this.buttonОчиститьСтрокуПоискаИсходящихДокументов.Name = "buttonОчиститьСтрокуПоискаИсходящихДокументов";
            this.buttonОчиститьСтрокуПоискаИсходящихДокументов.Size = new System.Drawing.Size(64, 22);
            this.buttonОчиститьСтрокуПоискаИсходящихДокументов.TabIndex = 7;
            this.buttonОчиститьСтрокуПоискаИсходящихДокументов.Text = "Очистить";
            this.toolTip1.SetToolTip(this.buttonОчиститьСтрокуПоискаИсходящихДокументов, "Очистить условия поиска.");
            this.buttonОчиститьСтрокуПоискаИсходящихДокументов.Click += new System.EventHandler(this.buttonОчиститьСтрокуПоискаИсходящихДокументов_Click);
            // 
            // panel4Tab1
            // 
            this.panel4Tab1.Controls.Add(this.dataGridРабочиеДокументы);
            this.panel4Tab1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4Tab1.Location = new System.Drawing.Point(0, 0);
            this.panel4Tab1.Name = "panel4Tab1";
            this.panel4Tab1.Size = new System.Drawing.Size(740, 150);
            this.panel4Tab1.TabIndex = 2;
            // 
            // dataGridРабочиеДокументы
            // 
            this.dataGridРабочиеДокументы.CaptionFont = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridРабочиеДокументы.CaptionText = "Входящие документы ожидающие рассмотрения";
            this.dataGridРабочиеДокументы.DataMember = "";
            this.dataGridРабочиеДокументы.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridРабочиеДокументы.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridРабочиеДокументы.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridРабочиеДокументы.Location = new System.Drawing.Point(0, 0);
            this.dataGridРабочиеДокументы.Name = "dataGridРабочиеДокументы";
            this.dataGridРабочиеДокументы.ReadOnly = true;
            this.dataGridРабочиеДокументы.Size = new System.Drawing.Size(740, 150);
            this.dataGridРабочиеДокументы.TabIndex = 0;
            this.dataGridРабочиеДокументы.TableStyles.AddRange(new System.Windows.Forms.DataGridTableStyle[] {
            this.dataGridTableStyleРабочиеДокументы});
            this.dataGridРабочиеДокументы.Resize += new System.EventHandler(this.dataGridРабочиеДокументы_Resize);
            this.dataGridРабочиеДокументы.DoubleClick += new System.EventHandler(this.dataGridРабочиеДокументы_DoubleClick);
            this.dataGridРабочиеДокументы.CurrentCellChanged += new System.EventHandler(this.dataGridРабочиеДокументы_CurrentCellChanged);
            this.dataGridРабочиеДокументы.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dataGridРабочиеДокументы_MouseUp);
            this.dataGridРабочиеДокументы.Leave += new System.EventHandler(this.dataGridРабочиеДокументы_Leave);
            // 
            // dataGridTableStyleРабочиеДокументы
            // 
            this.dataGridTableStyleРабочиеДокументы.AlternatingBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.dataGridTableStyleРабочиеДокументы.DataGrid = this.dataGridРабочиеДокументы;
            this.dataGridTableStyleРабочиеДокументы.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
            this.dataGridTextBoxColumn1,
            this.dataGridTextBoxColumn2,
            this.dataGridTextBoxColumn3,
            this.dataGridTextBoxColumn4,
            this.dataGridTextBoxColumn5,
            this.dataGridTextBoxColumn6,
            this.dataGridTextBoxColumn7,
            this.dataGridTextBoxColumn8});
            this.dataGridTableStyleРабочиеДокументы.HeaderFont = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridTableStyleРабочиеДокументы.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridTableStyleРабочиеДокументы.MappingName = "Выборка";
            this.dataGridTableStyleРабочиеДокументы.ReadOnly = true;
            this.dataGridTableStyleРабочиеДокументы.RowHeadersVisible = false;
            // 
            // dataGridTextBoxColumn1
            // 
            this.dataGridTextBoxColumn1.Format = "";
            this.dataGridTextBoxColumn1.FormatInfo = null;
            this.dataGridTextBoxColumn1.HeaderText = "Документ";
            this.dataGridTextBoxColumn1.MappingName = "ОписаниеДокумента";
            this.dataGridTextBoxColumn1.NullText = "";
            this.dataGridTextBoxColumn1.Width = 75;
            // 
            // dataGridTextBoxColumn2
            // 
            this.dataGridTextBoxColumn2.Format = "";
            this.dataGridTextBoxColumn2.FormatInfo = null;
            this.dataGridTextBoxColumn2.HeaderText = "Корр-т";
            this.dataGridTextBoxColumn2.MappingName = "ОписаниеКорреспондента";
            this.dataGridTextBoxColumn2.NullText = "";
            this.dataGridTextBoxColumn2.Width = 75;
            // 
            // dataGridTextBoxColumn3
            // 
            this.dataGridTextBoxColumn3.Format = "";
            this.dataGridTextBoxColumn3.FormatInfo = null;
            this.dataGridTextBoxColumn3.HeaderText = "Дата отпр.";
            this.dataGridTextBoxColumn3.MappingName = "ДатаИсхода";
            this.dataGridTextBoxColumn3.NullText = "";
            this.dataGridTextBoxColumn3.Width = 65;
            // 
            // dataGridTextBoxColumn4
            // 
            this.dataGridTextBoxColumn4.Format = "";
            this.dataGridTextBoxColumn4.FormatInfo = null;
            this.dataGridTextBoxColumn4.HeaderText = "№ исход.";
            this.dataGridTextBoxColumn4.MappingName = "НомерИсход";
            this.dataGridTextBoxColumn4.NullText = "";
            this.dataGridTextBoxColumn4.Width = 65;
            // 
            // dataGridTextBoxColumn5
            // 
            this.dataGridTextBoxColumn5.Format = "";
            this.dataGridTextBoxColumn5.FormatInfo = null;
            this.dataGridTextBoxColumn5.HeaderText = "Дата пост.";
            this.dataGridTextBoxColumn5.MappingName = "ДатаПоступ";
            this.dataGridTextBoxColumn5.NullText = "";
            this.dataGridTextBoxColumn5.Width = 65;
            // 
            // dataGridTextBoxColumn6
            // 
            this.dataGridTextBoxColumn6.Format = "";
            this.dataGridTextBoxColumn6.FormatInfo = null;
            this.dataGridTextBoxColumn6.HeaderText = "№ вход.";
            this.dataGridTextBoxColumn6.MappingName = "НомерВход";
            this.dataGridTextBoxColumn6.NullText = "";
            this.dataGridTextBoxColumn6.Width = 75;
            // 
            // dataGridTextBoxColumn7
            // 
            this.dataGridTextBoxColumn7.Format = "";
            this.dataGridTextBoxColumn7.FormatInfo = null;
            this.dataGridTextBoxColumn7.HeaderText = "Содержание";
            this.dataGridTextBoxColumn7.MappingName = "КраткоеСодержание";
            this.dataGridTextBoxColumn7.NullText = "";
            this.dataGridTextBoxColumn7.Width = 240;
            // 
            // dataGridTextBoxColumn8
            // 
            this.dataGridTextBoxColumn8.Format = "";
            this.dataGridTextBoxColumn8.FormatInfo = null;
            this.dataGridTextBoxColumn8.HeaderText = "Контроль";
            this.dataGridTextBoxColumn8.MappingName = "СрокВыполнения";
            this.dataGridTextBoxColumn8.NullText = "";
            this.dataGridTextBoxColumn8.Width = 65;
            // 
            // mainMenu1
            // 
            this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem1,
            this.menuItem2,
            this.menuItem20,
            this.menuItem17,
            this.menuItem3});
            // 
            // menuItem1
            // 
            this.menuItem1.Index = 0;
            this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem8,
            this.menuItem9,
            this.menuItemСохранитьВФайл});
            this.menuItem1.Text = "Файл";
            // 
            // menuItem8
            // 
            this.menuItem8.Index = 0;
            this.menuItem8.Text = "Год";
            this.menuItem8.Click += new System.EventHandler(this.menuItem8_Click);
            // 
            // menuItem9
            // 
            this.menuItem9.Index = 1;
            this.menuItem9.Text = "-";
            // 
            // menuItemСохранитьВФайл
            // 
            this.menuItemСохранитьВФайл.Index = 2;
            this.menuItemСохранитьВФайл.Text = "Выход";
            this.menuItemСохранитьВФайл.Click += new System.EventHandler(this.menuItemЗакрыть_Click);
            // 
            // menuItem2
            // 
            this.menuItem2.Index = 1;
            this.menuItem2.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItemСправочникиКорреспонденты,
            this.menuItemСправочникиПодразделения,
            this.menuItemСправочникиПолучатели,
            this.menuItemСправочникиДокументы,
            this.menuItem11,
            this.menuItem12,
            this.menuItem31});
            this.menuItem2.Text = "Справочники";
            // 
            // menuItemСправочникиКорреспонденты
            // 
            this.menuItemСправочникиКорреспонденты.Index = 0;
            this.menuItemСправочникиКорреспонденты.Text = "Адресаты";
            this.menuItemСправочникиКорреспонденты.Click += new System.EventHandler(this.menuItemСправочникиКорреспонденты_Click);
            // 
            // menuItemСправочникиПодразделения
            // 
            this.menuItemСправочникиПодразделения.Index = 1;
            this.menuItemСправочникиПодразделения.Text = "Подразделения";
            this.menuItemСправочникиПодразделения.Click += new System.EventHandler(this.menuItemСправочникиПодразделения_Click);
            // 
            // menuItemСправочникиПолучатели
            // 
            this.menuItemСправочникиПолучатели.Index = 2;
            this.menuItemСправочникиПолучатели.Text = "Сотрудники";
            this.menuItemСправочникиПолучатели.Click += new System.EventHandler(this.menuItemСправочникиПолучатели_Click);
            // 
            // menuItemСправочникиДокументы
            // 
            this.menuItemСправочникиДокументы.Index = 3;
            this.menuItemСправочникиДокументы.Text = "Типы документов";
            this.menuItemСправочникиДокументы.Click += new System.EventHandler(this.menuItemСправочникиДокументы_Click);
            // 
            // menuItem11
            // 
            this.menuItem11.Index = 4;
            this.menuItem11.Text = "Персональные данные";
            this.menuItem11.Click += new System.EventHandler(this.menuItem11_Click);
            // 
            // menuItem12
            // 
            this.menuItem12.Index = 5;
            this.menuItem12.Text = "Цель получения персональных данных";
            this.menuItem12.Click += new System.EventHandler(this.menuItem12_Click);
            // 
            // menuItem31
            // 
            this.menuItem31.Index = 6;
            this.menuItem31.Text = "Справочник руководитель отдел";
            this.menuItem31.Click += new System.EventHandler(this.menuItem31_Click);
            // 
            // menuItem20
            // 
            this.menuItem20.Index = 2;
            this.menuItem20.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem21,
            this.menuItem22,
            this.menuItem23});
            this.menuItem20.Text = "Отчеты";
            this.menuItem20.Click += new System.EventHandler(this.menuItem20_Click);
            // 
            // menuItem21
            // 
            this.menuItem21.Index = 0;
            this.menuItem21.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem24,
            this.menuItem25});
            this.menuItem21.Text = "Общие отчёты";
            this.menuItem21.Click += new System.EventHandler(this.menuItem21_Click);
            // 
            // menuItem24
            // 
            this.menuItem24.Index = 0;
            this.menuItem24.Text = "Контрольное уведомление";
            this.menuItem24.Click += new System.EventHandler(this.menuItem24_Click);
            // 
            // menuItem25
            // 
            this.menuItem25.Index = 1;
            this.menuItem25.Text = "Документы с истёкшими сроками исполнения";
            this.menuItem25.Click += new System.EventHandler(this.menuItem25_Click);
            // 
            // menuItem22
            // 
            this.menuItem22.Index = 1;
            this.menuItem22.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem26,
            this.menuItem27,
            this.menuItem28,
            this.menuItem33});
            this.menuItem22.Text = "Отчёты по входящей корреспонденции";
            this.menuItem22.Click += new System.EventHandler(this.menuItem22_Click);
            // 
            // menuItem26
            // 
            this.menuItem26.Index = 0;
            this.menuItem26.Text = "Статистика по входящей корреспонденции";
            this.menuItem26.Click += new System.EventHandler(this.menuItem26_Click);
            // 
            // menuItem27
            // 
            this.menuItem27.Index = 1;
            this.menuItem27.Text = "Отчет о входящих документах";
            this.menuItem27.Click += new System.EventHandler(this.menuItem27_Click);
            // 
            // menuItem28
            // 
            this.menuItem28.Index = 2;
            this.menuItem28.Text = "Печать карточки";
            this.menuItem28.Click += new System.EventHandler(this.menuItem28_Click);
            // 
            // menuItem23
            // 
            this.menuItem23.Index = 2;
            this.menuItem23.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem29,
            this.menuItem30,
            this.menuItem32});
            this.menuItem23.Text = "Отчёты по исходящей корреспонденции";
            this.menuItem23.Click += new System.EventHandler(this.menuItem23_Click);
            // 
            // menuItem29
            // 
            this.menuItem29.Index = 0;
            this.menuItem29.Text = "Статистика по исходящей корреспонденции";
            this.menuItem29.Click += new System.EventHandler(this.menuItem29_Click);
            // 
            // menuItem30
            // 
            this.menuItem30.Index = 1;
            this.menuItem30.Text = "Отчет об исходящих документах";
            this.menuItem30.Click += new System.EventHandler(this.menuItem30_Click);
            // 
            // menuItem32
            // 
            this.menuItem32.Index = 2;
            this.menuItem32.Text = "Журнал учета передачи персональных данных";
            this.menuItem32.Click += new System.EventHandler(this.menuItem32_Click);
            // 
            // menuItem17
            // 
            this.menuItem17.Index = 3;
            this.menuItem17.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem18,
            this.menuItem19});
            this.menuItem17.Text = "Карточка";
            this.menuItem17.Visible = false;
            // 
            // menuItem18
            // 
            this.menuItem18.Index = 0;
            this.menuItem18.Text = "Входящая";
            this.menuItem18.Click += new System.EventHandler(this.menuItem18_Click);
            // 
            // menuItem19
            // 
            this.menuItem19.Index = 1;
            this.menuItem19.Text = "Исходящая";
            this.menuItem19.Click += new System.EventHandler(this.menuItem19_Click);
            // 
            // menuItem3
            // 
            this.menuItem3.Index = 4;
            this.menuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItemПросрочДокументы,
            this.menuItemКонтрольныеУведомления,
            this.menuItemContextПечатьКарточки,
            this.menuItem5,
            this.menuItem4,
            this.menuItem7,
            this.menuItem6,
            this.menuItem10,
            this.menuItem13,
            this.menuItem14,
            this.menuItem15,
            this.menuItem16});
            this.menuItem3.Text = "Отчеты";
            this.menuItem3.Visible = false;
            // 
            // menuItemПросрочДокументы
            // 
            this.menuItemПросрочДокументы.Index = 0;
            this.menuItemПросрочДокументы.Text = "Документы с истекшими сроками исполнения";
            this.menuItemПросрочДокументы.Click += new System.EventHandler(this.menuItemПросрочДокументы_Click);
            // 
            // menuItemКонтрольныеУведомления
            // 
            this.menuItemКонтрольныеУведомления.Index = 1;
            this.menuItemКонтрольныеУведомления.Text = "Контрольные уведомления";
            this.menuItemКонтрольныеУведомления.Click += new System.EventHandler(this.menuItemКонтрольныеУведомления_Click);
            // 
            // menuItemContextПечатьКарточки
            // 
            this.menuItemContextПечатьКарточки.Index = 2;
            this.menuItemContextПечатьКарточки.Text = "Печать карточки";
            this.menuItemContextПечатьКарточки.Click += new System.EventHandler(this.menuItemContextПечатьКарточки_Click);
            // 
            // menuItem5
            // 
            this.menuItem5.Index = 3;
            this.menuItem5.Text = "Статистика по исполнителям";
            this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
            // 
            // menuItem4
            // 
            this.menuItem4.Index = 4;
            this.menuItem4.Text = "Статистика по корреспондентам";
            this.menuItem4.Click += new System.EventHandler(this.menuItem4_Click);
            // 
            // menuItem7
            // 
            this.menuItem7.Index = 5;
            this.menuItem7.Text = "-";
            // 
            // menuItem6
            // 
            this.menuItem6.Index = 6;
            this.menuItem6.Text = "Статистика по исходящей корреспонденции";
            this.menuItem6.Click += new System.EventHandler(this.menuItem6_Click);
            // 
            // menuItem10
            // 
            this.menuItem10.Index = 7;
            this.menuItem10.Text = "Журнал учёта передачи персональных данных";
            this.menuItem10.Click += new System.EventHandler(this.menuItem10_Click);
            // 
            // menuItem13
            // 
            this.menuItem13.Index = 8;
            this.menuItem13.Text = "Печать карточки";
            this.menuItem13.Click += new System.EventHandler(this.menuItem13_Click);
            // 
            // menuItem14
            // 
            this.menuItem14.Index = 9;
            this.menuItem14.Text = "-";
            // 
            // menuItem15
            // 
            this.menuItem15.Index = 10;
            this.menuItem15.Text = "Отчёт о документах";
            this.menuItem15.Click += new System.EventHandler(this.menuItem15_Click);
            // 
            // menuItem16
            // 
            this.menuItem16.Index = 11;
            this.menuItem16.Text = "Отчет документов по исполнителям";
            this.menuItem16.Click += new System.EventHandler(this.menuItem16_Click);
            // 
            // tabControlВходящиеДокументы
            // 
            this.tabControlВходящиеДокументы.Controls.Add(this.tabPage1);
            this.tabControlВходящиеДокументы.Controls.Add(this.tabPage2);
            this.tabControlВходящиеДокументы.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlВходящиеДокументы.Location = new System.Drawing.Point(3, 3);
            this.tabControlВходящиеДокументы.Name = "tabControlВходящиеДокументы";
            this.tabControlВходящиеДокументы.SelectedIndex = 0;
            this.tabControlВходящиеДокументы.Size = new System.Drawing.Size(748, 286);
            this.tabControlВходящиеДокументы.TabIndex = 3;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.panel4Tab1);
            this.tabPage1.Controls.Add(this.panel1Tab1);
            this.tabPage1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Size = new System.Drawing.Size(740, 260);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Рабочие документы";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // panel1Tab1
            // 
            this.panel1Tab1.Controls.Add(this.panel1);
            this.panel1Tab1.Controls.Add(this.panel2);
            this.panel1Tab1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1Tab1.Location = new System.Drawing.Point(0, 150);
            this.panel1Tab1.Name = "panel1Tab1";
            this.panel1Tab1.Size = new System.Drawing.Size(740, 110);
            this.panel1Tab1.TabIndex = 3;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.labelИнфоTab1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(280, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(460, 110);
            this.panel1.TabIndex = 1;
            // 
            // labelИнфоTab1
            // 
            this.labelИнфоTab1.BackColor = System.Drawing.SystemColors.Window;
            this.labelИнфоTab1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelИнфоTab1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelИнфоTab1.Location = new System.Drawing.Point(0, 0);
            this.labelИнфоTab1.Name = "labelИнфоTab1";
            this.labelИнфоTab1.ReadOnly = true;
            this.labelИнфоTab1.Size = new System.Drawing.Size(456, 106);
            this.labelИнфоTab1.TabIndex = 0;
            this.labelИнфоTab1.TabStop = false;
            this.labelИнфоTab1.Text = "";
            this.labelИнфоTab1.Leave += new System.EventHandler(this.labelИнфоTab1_Leave);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.panel1Tab2);
            this.tabPage2.Controls.Add(this.panel5);
            this.tabPage2.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Size = new System.Drawing.Size(740, 260);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Документы \"В деле\"";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // panel1Tab2
            // 
            this.panel1Tab2.Controls.Add(this.dataGridДокументыВДеле);
            this.panel1Tab2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1Tab2.Location = new System.Drawing.Point(0, 0);
            this.panel1Tab2.Name = "panel1Tab2";
            this.panel1Tab2.Size = new System.Drawing.Size(740, 150);
            this.panel1Tab2.TabIndex = 3;
            // 
            // dataGridДокументыВДеле
            // 
            this.dataGridДокументыВДеле.CaptionFont = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridДокументыВДеле.CaptionText = "Входящие документы списанные \"В дело\"";
            this.dataGridДокументыВДеле.DataMember = "";
            this.dataGridДокументыВДеле.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridДокументыВДеле.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridДокументыВДеле.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridДокументыВДеле.Location = new System.Drawing.Point(0, 0);
            this.dataGridДокументыВДеле.Name = "dataGridДокументыВДеле";
            this.dataGridДокументыВДеле.ReadOnly = true;
            this.dataGridДокументыВДеле.Size = new System.Drawing.Size(740, 150);
            this.dataGridДокументыВДеле.TabIndex = 0;
            this.dataGridДокументыВДеле.TableStyles.AddRange(new System.Windows.Forms.DataGridTableStyle[] {
            this.dataGridTableStyleДокументыВДеле});
            this.dataGridДокументыВДеле.Resize += new System.EventHandler(this.dataGridДокументыВДеле_Resize);
            this.dataGridДокументыВДеле.DoubleClick += new System.EventHandler(this.dataGridДокументыВДеле_DoubleClick);
            this.dataGridДокументыВДеле.CurrentCellChanged += new System.EventHandler(this.dataGridДокументыВДеле_CurrentCellChanged);
            this.dataGridДокументыВДеле.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dataGridДокументыВДеле_MouseUp);
            this.dataGridДокументыВДеле.Leave += new System.EventHandler(this.dataGridДокументыВДеле_Leave);
            // 
            // dataGridTableStyleДокументыВДеле
            // 
            this.dataGridTableStyleДокументыВДеле.AlternatingBackColor = System.Drawing.Color.LavenderBlush;
            this.dataGridTableStyleДокументыВДеле.DataGrid = this.dataGridДокументыВДеле;
            this.dataGridTableStyleДокументыВДеле.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
            this.dataGridTextBoxColumn9,
            this.dataGridTextBoxColumn10,
            this.dataGridTextBoxColumn11,
            this.dataGridTextBoxColumn12,
            this.dataGridTextBoxColumn13,
            this.dataGridTextBoxColumn14,
            this.dataGridTextBoxColumn15,
            this.dataGridTextBoxColumn16});
            this.dataGridTableStyleДокументыВДеле.HeaderFont = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridTableStyleДокументыВДеле.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridTableStyleДокументыВДеле.MappingName = "Выборка";
            this.dataGridTableStyleДокументыВДеле.RowHeadersVisible = false;
            // 
            // dataGridTextBoxColumn9
            // 
            this.dataGridTextBoxColumn9.Format = "";
            this.dataGridTextBoxColumn9.FormatInfo = null;
            this.dataGridTextBoxColumn9.HeaderText = "Документ";
            this.dataGridTextBoxColumn9.MappingName = "ОписаниеДокумента";
            this.dataGridTextBoxColumn9.NullText = "";
            this.dataGridTextBoxColumn9.Width = 75;
            // 
            // dataGridTextBoxColumn10
            // 
            this.dataGridTextBoxColumn10.Format = "";
            this.dataGridTextBoxColumn10.FormatInfo = null;
            this.dataGridTextBoxColumn10.HeaderText = "Корр-т";
            this.dataGridTextBoxColumn10.MappingName = "ОписаниеКорреспондента";
            this.dataGridTextBoxColumn10.NullText = "";
            this.dataGridTextBoxColumn10.Width = 75;
            // 
            // dataGridTextBoxColumn11
            // 
            this.dataGridTextBoxColumn11.Format = "";
            this.dataGridTextBoxColumn11.FormatInfo = null;
            this.dataGridTextBoxColumn11.HeaderText = "Дата отпр.";
            this.dataGridTextBoxColumn11.MappingName = "ДатаИсхода";
            this.dataGridTextBoxColumn11.NullText = "";
            this.dataGridTextBoxColumn11.Width = 65;
            // 
            // dataGridTextBoxColumn12
            // 
            this.dataGridTextBoxColumn12.Format = "";
            this.dataGridTextBoxColumn12.FormatInfo = null;
            this.dataGridTextBoxColumn12.HeaderText = "№ исход.";
            this.dataGridTextBoxColumn12.MappingName = "НомерИсход";
            this.dataGridTextBoxColumn12.NullText = "";
            this.dataGridTextBoxColumn12.Width = 65;
            // 
            // dataGridTextBoxColumn13
            // 
            this.dataGridTextBoxColumn13.Format = "";
            this.dataGridTextBoxColumn13.FormatInfo = null;
            this.dataGridTextBoxColumn13.HeaderText = "Дата поступ.";
            this.dataGridTextBoxColumn13.MappingName = "ДатаПоступ";
            this.dataGridTextBoxColumn13.NullText = "";
            this.dataGridTextBoxColumn13.Width = 65;
            // 
            // dataGridTextBoxColumn14
            // 
            this.dataGridTextBoxColumn14.Format = "";
            this.dataGridTextBoxColumn14.FormatInfo = null;
            this.dataGridTextBoxColumn14.HeaderText = "№ вход.";
            this.dataGridTextBoxColumn14.MappingName = "НомерВход";
            this.dataGridTextBoxColumn14.NullText = "";
            this.dataGridTextBoxColumn14.Width = 75;
            // 
            // dataGridTextBoxColumn15
            // 
            this.dataGridTextBoxColumn15.Format = "";
            this.dataGridTextBoxColumn15.FormatInfo = null;
            this.dataGridTextBoxColumn15.HeaderText = "Содержание";
            this.dataGridTextBoxColumn15.MappingName = "КраткоеСодержание";
            this.dataGridTextBoxColumn15.NullText = "";
            this.dataGridTextBoxColumn15.Width = 180;
            // 
            // dataGridTextBoxColumn16
            // 
            this.dataGridTextBoxColumn16.Format = "";
            this.dataGridTextBoxColumn16.FormatInfo = null;
            this.dataGridTextBoxColumn16.HeaderText = "Резолюция";
            this.dataGridTextBoxColumn16.MappingName = "Резолюция";
            this.dataGridTextBoxColumn16.NullText = "";
            this.dataGridTextBoxColumn16.Width = 125;
            // 
            // panel5
            // 
            this.panel5.Controls.Add(this.panel7);
            this.panel5.Controls.Add(this.panel4);
            this.panel5.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel5.Location = new System.Drawing.Point(0, 150);
            this.panel5.Name = "panel5";
            this.panel5.Size = new System.Drawing.Size(740, 110);
            this.panel5.TabIndex = 0;
            // 
            // panel7
            // 
            this.panel7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel7.Controls.Add(this.labelИнфоTab2);
            this.panel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel7.Location = new System.Drawing.Point(280, 0);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(460, 110);
            this.panel7.TabIndex = 4;
            // 
            // labelИнфоTab2
            // 
            this.labelИнфоTab2.BackColor = System.Drawing.SystemColors.Window;
            this.labelИнфоTab2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelИнфоTab2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelИнфоTab2.Location = new System.Drawing.Point(0, 0);
            this.labelИнфоTab2.Name = "labelИнфоTab2";
            this.labelИнфоTab2.ReadOnly = true;
            this.labelИнфоTab2.Size = new System.Drawing.Size(456, 106);
            this.labelИнфоTab2.TabIndex = 0;
            this.labelИнфоTab2.TabStop = false;
            this.labelИнфоTab2.Text = "";
            this.labelИнфоTab2.Leave += new System.EventHandler(this.labelИнфоTab2_Leave);
            // 
            // panel4
            // 
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel4.Controls.Add(this.checkBoxКорреспонденты);
            this.panel4.Controls.Add(this.comboBoxКорреспонденты);
            this.panel4.Controls.Add(this.labelОтобраноДокументовПоискомTab2);
            this.panel4.Controls.Add(this.buttonОчиститьСтрокуПоискаTab2);
            this.panel4.Controls.Add(this.panel6);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(280, 110);
            this.panel4.TabIndex = 3;
            // 
            // checkBoxКорреспонденты
            // 
            this.checkBoxКорреспонденты.AutoSize = true;
            this.checkBoxКорреспонденты.Location = new System.Drawing.Point(4, 55);
            this.checkBoxКорреспонденты.Name = "checkBoxКорреспонденты";
            this.checkBoxКорреспонденты.Size = new System.Drawing.Size(195, 17);
            this.checkBoxКорреспонденты.TabIndex = 6;
            this.checkBoxКорреспонденты.Text = "Скрыть список корреспондентов";
            this.checkBoxКорреспонденты.UseVisualStyleBackColor = true;
            this.checkBoxКорреспонденты.CheckedChanged += new System.EventHandler(this.checkBoxКорреспонденты_CheckedChanged);
            // 
            // comboBoxКорреспонденты
            // 
            this.comboBoxКорреспонденты.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBoxКорреспонденты.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBoxКорреспонденты.DisplayMember = "ОписаниеКорреспондента";
            this.comboBoxКорреспонденты.DropDownHeight = 400;
            this.comboBoxКорреспонденты.DropDownWidth = 400;
            this.comboBoxКорреспонденты.FormattingEnabled = true;
            this.comboBoxКорреспонденты.IntegralHeight = false;
            this.comboBoxКорреспонденты.Location = new System.Drawing.Point(0, 27);
            this.comboBoxКорреспонденты.Name = "comboBoxКорреспонденты";
            this.comboBoxКорреспонденты.Size = new System.Drawing.Size(269, 21);
            this.comboBoxКорреспонденты.TabIndex = 5;
            this.comboBoxКорреспонденты.ValueMember = "ОписаниеКорреспондента";
            // 
            // labelОтобраноДокументовПоискомTab2
            // 
            this.labelОтобраноДокументовПоискомTab2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.labelОтобраноДокументовПоискомTab2.Location = new System.Drawing.Point(0, 89);
            this.labelОтобраноДокументовПоискомTab2.Name = "labelОтобраноДокументовПоискомTab2";
            this.labelОтобраноДокументовПоискомTab2.Size = new System.Drawing.Size(276, 17);
            this.labelОтобраноДокументовПоискомTab2.TabIndex = 4;
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.textBoxСтрокаПоискаTab2);
            this.panel6.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel6.Location = new System.Drawing.Point(0, 0);
            this.panel6.Name = "panel6";
            this.panel6.Size = new System.Drawing.Size(276, 20);
            this.panel6.TabIndex = 2;
            // 
            // dataGridTableStyle2
            // 
            this.dataGridTableStyle2.AlternatingBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.dataGridTableStyle2.DataGrid = null;
            this.dataGridTableStyle2.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
            this.dataGridTextBoxColumnДокумент,
            this.dataGridTextBoxColumnКорреспондент,
            this.dataGridTextBoxColumnДатаОтправ,
            this.dataGridTextBoxColumnДатаПоступ,
            this.dataGridTextBoxColumnНомерИсход,
            this.dataGridTextBoxColumnНомерВход,
            this.dataGridTextBoxColumnСодержание,
            this.dataGridTextBoxColumnКонтроль,
            this.dataGridBoolColumnВДеле});
            this.dataGridTableStyle2.HeaderFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridTableStyle2.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridTableStyle2.MappingName = "Выборка";
            this.dataGridTableStyle2.ReadOnly = true;
            this.dataGridTableStyle2.RowHeadersVisible = false;
            // 
            // dataGridTextBoxColumnДокумент
            // 
            this.dataGridTextBoxColumnДокумент.Format = "";
            this.dataGridTextBoxColumnДокумент.FormatInfo = null;
            this.dataGridTextBoxColumnДокумент.HeaderText = "Документ";
            this.dataGridTextBoxColumnДокумент.MappingName = "ОписаниеДокумента";
            this.dataGridTextBoxColumnДокумент.NullText = "";
            this.dataGridTextBoxColumnДокумент.ReadOnly = true;
            this.dataGridTextBoxColumnДокумент.Width = 75;
            // 
            // dataGridTextBoxColumnКорреспондент
            // 
            this.dataGridTextBoxColumnКорреспондент.Format = "";
            this.dataGridTextBoxColumnКорреспондент.FormatInfo = null;
            this.dataGridTextBoxColumnКорреспондент.HeaderText = "Корр-т";
            this.dataGridTextBoxColumnКорреспондент.MappingName = "ОписаниеКорреспондента";
            this.dataGridTextBoxColumnКорреспондент.NullText = "";
            this.dataGridTextBoxColumnКорреспондент.ReadOnly = true;
            this.dataGridTextBoxColumnКорреспондент.Width = 75;
            // 
            // dataGridTextBoxColumnДатаОтправ
            // 
            this.dataGridTextBoxColumnДатаОтправ.Format = "";
            this.dataGridTextBoxColumnДатаОтправ.FormatInfo = null;
            this.dataGridTextBoxColumnДатаОтправ.HeaderText = "Отправлено";
            this.dataGridTextBoxColumnДатаОтправ.MappingName = "ДатаИсхода";
            this.dataGridTextBoxColumnДатаОтправ.NullText = "";
            this.dataGridTextBoxColumnДатаОтправ.ReadOnly = true;
            this.dataGridTextBoxColumnДатаОтправ.Width = 67;
            // 
            // dataGridTextBoxColumnДатаПоступ
            // 
            this.dataGridTextBoxColumnДатаПоступ.Format = "";
            this.dataGridTextBoxColumnДатаПоступ.FormatInfo = null;
            this.dataGridTextBoxColumnДатаПоступ.HeaderText = "Поступило";
            this.dataGridTextBoxColumnДатаПоступ.MappingName = "ДатаПоступ";
            this.dataGridTextBoxColumnДатаПоступ.NullText = "";
            this.dataGridTextBoxColumnДатаПоступ.ReadOnly = true;
            this.dataGridTextBoxColumnДатаПоступ.Width = 67;
            // 
            // dataGridTextBoxColumnНомерИсход
            // 
            this.dataGridTextBoxColumnНомерИсход.Format = "";
            this.dataGridTextBoxColumnНомерИсход.FormatInfo = null;
            this.dataGridTextBoxColumnНомерИсход.HeaderText = "№исход.";
            this.dataGridTextBoxColumnНомерИсход.MappingName = "НомерИсход";
            this.dataGridTextBoxColumnНомерИсход.NullText = "";
            this.dataGridTextBoxColumnНомерИсход.ReadOnly = true;
            this.dataGridTextBoxColumnНомерИсход.Width = 65;
            // 
            // dataGridTextBoxColumnНомерВход
            // 
            this.dataGridTextBoxColumnНомерВход.Format = "";
            this.dataGridTextBoxColumnНомерВход.FormatInfo = null;
            this.dataGridTextBoxColumnНомерВход.HeaderText = "№вход.";
            this.dataGridTextBoxColumnНомерВход.MappingName = "НомерВход";
            this.dataGridTextBoxColumnНомерВход.NullText = "";
            this.dataGridTextBoxColumnНомерВход.ReadOnly = true;
            this.dataGridTextBoxColumnНомерВход.Width = 65;
            // 
            // dataGridTextBoxColumnСодержание
            // 
            this.dataGridTextBoxColumnСодержание.Format = "";
            this.dataGridTextBoxColumnСодержание.FormatInfo = null;
            this.dataGridTextBoxColumnСодержание.HeaderText = "Содержание";
            this.dataGridTextBoxColumnСодержание.MappingName = "КраткоеСодержание";
            this.dataGridTextBoxColumnСодержание.NullText = "";
            this.dataGridTextBoxColumnСодержание.ReadOnly = true;
            this.dataGridTextBoxColumnСодержание.Width = 250;
            // 
            // dataGridTextBoxColumnКонтроль
            // 
            this.dataGridTextBoxColumnКонтроль.Format = "";
            this.dataGridTextBoxColumnКонтроль.FormatInfo = null;
            this.dataGridTextBoxColumnКонтроль.HeaderText = "Контроль";
            this.dataGridTextBoxColumnКонтроль.MappingName = "СрокВыполнения";
            this.dataGridTextBoxColumnКонтроль.NullText = "";
            this.dataGridTextBoxColumnКонтроль.ReadOnly = true;
            this.dataGridTextBoxColumnКонтроль.Width = 65;
            // 
            // dataGridBoolColumnВДеле
            // 
            this.dataGridBoolColumnВДеле.HeaderText = "В деле";
            this.dataGridBoolColumnВДеле.MappingName = "ВДело";
            this.dataGridBoolColumnВДеле.NullText = "";
            this.dataGridBoolColumnВДеле.Width = 45;
            // 
            // tabControlТипыДокументов
            // 
            this.tabControlТипыДокументов.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabControlТипыДокументов.Controls.Add(this.tabPage3);
            this.tabControlТипыДокументов.Controls.Add(this.tabPage4);
            this.tabControlТипыДокументов.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControlТипыДокументов.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabControlТипыДокументов.ItemSize = new System.Drawing.Size(40, 30);
            this.tabControlТипыДокументов.Location = new System.Drawing.Point(0, 0);
            this.tabControlТипыДокументов.Multiline = true;
            this.tabControlТипыДокументов.Name = "tabControlТипыДокументов";
            this.tabControlТипыДокументов.SelectedIndex = 0;
            this.tabControlТипыДокументов.Size = new System.Drawing.Size(792, 300);
            this.tabControlТипыДокументов.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
            this.tabControlТипыДокументов.TabIndex = 4;
            this.tabControlТипыДокументов.SelectedIndexChanged += new System.EventHandler(this.tabControlТипыДокументов_SelectedIndexChanged);
            // 
            // tabPage3
            // 
            this.tabPage3.BackColor = System.Drawing.Color.Transparent;
            this.tabPage3.Controls.Add(this.tabControlВходящиеДокументы);
            this.tabPage3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabPage3.Location = new System.Drawing.Point(34, 4);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(754, 292);
            this.tabPage3.TabIndex = 0;
            this.tabPage3.Text = "Входящие";
            this.tabPage3.ToolTipText = "Документы входящие";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.BackColor = System.Drawing.Color.Transparent;
            this.tabPage4.Controls.Add(this.splitContainer1);
            this.tabPage4.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabPage4.Location = new System.Drawing.Point(34, 4);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(754, 292);
            this.tabPage4.TabIndex = 1;
            this.tabPage4.Text = "Исходящие";
            this.tabPage4.ToolTipText = "Документы исходящие";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // splitContainer1
            // 
            this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(3, 3);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.splitContainer3);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(748, 286);
            this.splitContainer1.SplitterDistance = 202;
            this.splitContainer1.TabIndex = 1;
            // 
            // splitContainer3
            // 
            this.splitContainer3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer3.Location = new System.Drawing.Point(0, 0);
            this.splitContainer3.Name = "splitContainer3";
            // 
            // splitContainer3.Panel1
            // 
            this.splitContainer3.Panel1.Controls.Add(this.dataGridИсходящиеДокументы);
            this.splitContainer3.Panel2MinSize = 0;
            this.splitContainer3.Size = new System.Drawing.Size(744, 198);
            this.splitContainer3.SplitterDistance = 737;
            this.splitContainer3.TabIndex = 1;
            // 
            // dataGridИсходящиеДокументы
            // 
            this.dataGridИсходящиеДокументы.CaptionFont = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridИсходящиеДокументы.CaptionText = "Исходящие документы";
            this.dataGridИсходящиеДокументы.DataMember = "";
            this.dataGridИсходящиеДокументы.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridИсходящиеДокументы.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridИсходящиеДокументы.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridИсходящиеДокументы.Location = new System.Drawing.Point(0, 0);
            this.dataGridИсходящиеДокументы.Name = "dataGridИсходящиеДокументы";
            this.dataGridИсходящиеДокументы.ReadOnly = true;
            this.dataGridИсходящиеДокументы.Size = new System.Drawing.Size(737, 198);
            this.dataGridИсходящиеДокументы.TabIndex = 0;
            this.dataGridИсходящиеДокументы.TableStyles.AddRange(new System.Windows.Forms.DataGridTableStyle[] {
            this.dataGridTableStyleИсходящиеДокументы});
            this.dataGridИсходящиеДокументы.Resize += new System.EventHandler(this.dataGridИсходящиеДокументы_Resize);
            this.dataGridИсходящиеДокументы.DoubleClick += new System.EventHandler(this.dataGridИсходящиеДокументы_DoubleClick);
            this.dataGridИсходящиеДокументы.CurrentCellChanged += new System.EventHandler(this.dataGridИсходящиеДокументы_CurrentCellChanged);
            this.dataGridИсходящиеДокументы.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dataGridИсходящиеДокументы_MouseUp);
            this.dataGridИсходящиеДокументы.Leave += new System.EventHandler(this.dataGridИсходящиеДокументы_Leave);
            // 
            // dataGridTableStyleИсходящиеДокументы
            // 
            this.dataGridTableStyleИсходящиеДокументы.AlternatingBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.dataGridTableStyleИсходящиеДокументы.DataGrid = this.dataGridИсходящиеДокументы;
            this.dataGridTableStyleИсходящиеДокументы.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
            this.dataGridTextBoxColumnИсхДокДатаИсхода,
            this.dataGridTextBoxColumnИсхДокНомер,
            this.dataGridTextBoxColumnИсхДокОписаниеАдресата,
            this.dataGridTextBoxColumnИсхДокСодержание,
            this.dataGridTextBoxColumnИсхДокНомерВходДокта});
            this.dataGridTableStyleИсходящиеДокументы.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridTableStyleИсходящиеДокументы.MappingName = "ВыборкаИсходящихДокументов";
            this.dataGridTableStyleИсходящиеДокументы.RowHeadersVisible = false;
            // 
            // dataGridTextBoxColumnИсхДокДатаИсхода
            // 
            this.dataGridTextBoxColumnИсхДокДатаИсхода.Format = "";
            this.dataGridTextBoxColumnИсхДокДатаИсхода.FormatInfo = null;
            this.dataGridTextBoxColumnИсхДокДатаИсхода.HeaderText = "Дата отпр.";
            this.dataGridTextBoxColumnИсхДокДатаИсхода.MappingName = "Дата";
            this.dataGridTextBoxColumnИсхДокДатаИсхода.NullText = "";
            this.dataGridTextBoxColumnИсхДокДатаИсхода.ReadOnly = true;
            this.dataGridTextBoxColumnИсхДокДатаИсхода.Width = 75;
            // 
            // dataGridTextBoxColumnИсхДокНомер
            // 
            this.dataGridTextBoxColumnИсхДокНомер.Format = "";
            this.dataGridTextBoxColumnИсхДокНомер.FormatInfo = null;
            this.dataGridTextBoxColumnИсхДокНомер.HeaderText = "Номер документа";
            this.dataGridTextBoxColumnИсхДокНомер.MappingName = "ТекстовыйНомер";
            this.dataGridTextBoxColumnИсхДокНомер.NullText = "";
            this.dataGridTextBoxColumnИсхДокНомер.ReadOnly = true;
            this.dataGridTextBoxColumnИсхДокНомер.Width = 110;
            // 
            // dataGridTextBoxColumnИсхДокОписаниеАдресата
            // 
            this.dataGridTextBoxColumnИсхДокОписаниеАдресата.Format = "";
            this.dataGridTextBoxColumnИсхДокОписаниеАдресата.FormatInfo = null;
            this.dataGridTextBoxColumnИсхДокОписаниеАдресата.HeaderText = "Адресат";
            this.dataGridTextBoxColumnИсхДокОписаниеАдресата.MappingName = "ОписаниеАдресата";
            this.dataGridTextBoxColumnИсхДокОписаниеАдресата.NullText = "";
            this.dataGridTextBoxColumnИсхДокОписаниеАдресата.ReadOnly = true;
            this.dataGridTextBoxColumnИсхДокОписаниеАдресата.Width = 150;
            // 
            // dataGridTextBoxColumnИсхДокСодержание
            // 
            this.dataGridTextBoxColumnИсхДокСодержание.Format = "";
            this.dataGridTextBoxColumnИсхДокСодержание.FormatInfo = null;
            this.dataGridTextBoxColumnИсхДокСодержание.HeaderText = "Содержание";
            this.dataGridTextBoxColumnИсхДокСодержание.MappingName = "Содержание";
            this.dataGridTextBoxColumnИсхДокСодержание.NullText = "";
            this.dataGridTextBoxColumnИсхДокСодержание.ReadOnly = true;
            this.dataGridTextBoxColumnИсхДокСодержание.Width = 300;
            // 
            // dataGridTextBoxColumnИсхДокНомерВходДокта
            // 
            this.dataGridTextBoxColumnИсхДокНомерВходДокта.Format = "";
            this.dataGridTextBoxColumnИсхДокНомерВходДокта.FormatInfo = null;
            this.dataGridTextBoxColumnИсхДокНомерВходДокта.HeaderText = "Входящий документ";
            this.dataGridTextBoxColumnИсхДокНомерВходДокта.MappingName = "НомерВходВходящегоДокумента";
            this.dataGridTextBoxColumnИсхДокНомерВходДокта.NullText = "";
            this.dataGridTextBoxColumnИсхДокНомерВходДокта.ReadOnly = true;
            this.dataGridTextBoxColumnИсхДокНомерВходДокта.Width = 110;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.comboBoxФильтрИДПоДате);
            this.splitContainer2.Panel1.Controls.Add(this.buttonОчиститьСтрокуПоискаИсходящихДокументов);
            this.splitContainer2.Panel1.Controls.Add(this.labelОтобраноДокументовПоискомИсходящихДокументов);
            this.splitContainer2.Panel1.Controls.Add(this.textBoxСтрокаПоискаИсходящихДокументов);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.labelИнфоTab3);
            this.splitContainer2.Size = new System.Drawing.Size(744, 76);
            this.splitContainer2.SplitterDistance = 267;
            this.splitContainer2.TabIndex = 0;
            // 
            // comboBoxФильтрИДПоДате
            // 
            this.comboBoxФильтрИДПоДате.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxФильтрИДПоДате.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBoxФильтрИДПоДате.FormattingEnabled = true;
            this.comboBoxФильтрИДПоДате.Items.AddRange(new object[] {
            "Весь год",
            "Январь",
            "Февраль",
            "Март",
            "Апрель",
            "Май",
            "Июнь",
            "Июль",
            "Август",
            "Сентябрь",
            "Октябрь",
            "Ноябрь",
            "Декабрь"});
            this.comboBoxФильтрИДПоДате.Location = new System.Drawing.Point(1, 27);
            this.comboBoxФильтрИДПоДате.Name = "comboBoxФильтрИДПоДате";
            this.comboBoxФильтрИДПоДате.Size = new System.Drawing.Size(189, 24);
            this.comboBoxФильтрИДПоДате.TabIndex = 8;
            this.comboBoxФильтрИДПоДате.SelectedIndexChanged += new System.EventHandler(this.comboBoxФильтрИДПоДате_SelectedIndexChanged);
            // 
            // labelОтобраноДокументовПоискомИсходящихДокументов
            // 
            this.labelОтобраноДокументовПоискомИсходящихДокументов.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.labelОтобраноДокументовПоискомИсходящихДокументов.Location = new System.Drawing.Point(0, 56);
            this.labelОтобраноДокументовПоискомИсходящихДокументов.Name = "labelОтобраноДокументовПоискомИсходящихДокументов";
            this.labelОтобраноДокументовПоискомИсходящихДокументов.Size = new System.Drawing.Size(267, 20);
            this.labelОтобраноДокументовПоискомИсходящихДокументов.TabIndex = 6;
            // 
            // labelИнфоTab3
            // 
            this.labelИнфоTab3.BackColor = System.Drawing.SystemColors.Window;
            this.labelИнфоTab3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.labelИнфоTab3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelИнфоTab3.Location = new System.Drawing.Point(0, 0);
            this.labelИнфоTab3.Name = "labelИнфоTab3";
            this.labelИнфоTab3.ReadOnly = true;
            this.labelИнфоTab3.Size = new System.Drawing.Size(473, 76);
            this.labelИнфоTab3.TabIndex = 1;
            this.labelИнфоTab3.TabStop = false;
            this.labelИнфоTab3.Text = "";
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.checkBox1, 0, 3);
            this.tableLayoutPanel2.Controls.Add(this.checkBox2, 0, 1);
            this.tableLayoutPanel2.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 4;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(200, 100);
            this.tableLayoutPanel2.TabIndex = 0;
            // 
            // checkBox1
            // 
            this.checkBox1.Appearance = System.Windows.Forms.Appearance.Button;
            this.checkBox1.AutoSize = true;
            this.checkBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checkBox1.Location = new System.Drawing.Point(3, 63);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(194, 34);
            this.checkBox1.TabIndex = 4;
            this.checkBox1.Text = "Февраль";
            this.checkBox1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.Appearance = System.Windows.Forms.Appearance.Button;
            this.checkBox2.AutoSize = true;
            this.checkBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checkBox2.Location = new System.Drawing.Point(3, 23);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(194, 14);
            this.checkBox2.TabIndex = 3;
            this.checkBox2.Text = "Весь год";
            this.checkBox2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(194, 13);
            this.label2.TabIndex = 0;
            this.label2.Text = "Фильтр по дате";
            // 
            // ds11
            // 
            this.ds11.DataSetName = "DS1";
            this.ds11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // menuItem33
            // 
            this.menuItem33.Index = 3;
            this.menuItem33.Text = "Журнал учета входящих персональных данных";
            this.menuItem33.Click += new System.EventHandler(this.menuItem33_Click_1);
            // 
            // FormГлавная
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(792, 300);
            this.Controls.Add(this.tabControlТипыДокументов);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Menu = this.mainMenu1;
            this.Name = "FormГлавная";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Регистрация корреспонденции";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormГлавная_FormClosing);
            this.Load += new System.EventHandler(this.FormГлавная_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataViewВыборкаРабДокументы)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4Tab1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridРабочиеДокументы)).EndInit();
            this.tabControlВходящиеДокументы.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panel1Tab1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.panel1Tab2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridДокументыВДеле)).EndInit();
            this.panel5.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataViewВыборкаДокументыВДеле)).EndInit();
            this.tabControlТипыДокументов.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer3.Panel1.ResumeLayout(false);
            this.splitContainer3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridИсходящиеДокументы)).EndInit();
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel1.PerformLayout();
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataViewИсходящиеДокументы)).EndInit();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion


        /// <summary>
        /// Точка входа в приложение
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (ПрограммаУжеЗапущена())
            {
                MessageBox.Show(null, "Программа уже запущена", "Регистрация корреспонденции");
                return;
            }

            Application.Run(new FormГлавная());
        }

        #region Методы

        /// <summary>
        /// Подключается к БД. 
        /// Очищает таблицы в ДатаСете. 
        /// Заполняет ДатаСет данными из БД.
        /// Создает ДатаВиев и подключает его в качестве источника данных к ДатаГрид
        /// </summary>
        private void ПодключитьсяПолучитьДанные()
        {
            потокОжидания = new System.Threading.Thread(new System.Threading.ThreadStart(ЗапуститьФормуОжидания));
            потокОжидания.Start();
            try
            {
                //=============Очистим ds11
                //входящие
                this.ds11.Карточка.Clear();
                this.ds11.Выборка.Clear();
                this.ds11.Документы.Clear();

                this.ds11.Корреспонденты.Clear();
                this.ds11.Получатели.Clear();


                //исходящие
                this.ds11.ВыборкаИсходящихДокументов.Clear();
                this.ds11.КарточкаИсходящая.Clear();

                this.ds11.ПодразделенияКомитета.Clear();
                this.ds11.ВыборкаКоличествоИсходящихДокументов.Clear();

                //Заплоним ds11 новыми данными =====================
                //входящие
                DS1TableAdapters.ДокументыTableAdapter документыTableAdapter = new RegKor.DS1TableAdapters.ДокументыTableAdapter();
                документыTableAdapter.Fill(ds11.Документы);

                //Определем ключ в конфигурационном файле если ГодДокумента = "истина" тогда 2012 год или больше
                //если ГодДокумента = "ложно" тогда 2011 год или раньше
                Classess.ГодДокументооборота год = new RegKor.Classess.ГодДокументооборота();
                bool flag = год.ГодВКонфигурационномФайле();

                ////if (flag == false)
                ////{
                ////    //Если год 2011 или более ранний то оставляем всё без изменения
                DS1TableAdapters.КорреспондентыTableAdapter корреспондентыTableAdapter = new RegKor.DS1TableAdapters.КорреспондентыTableAdapter();
                корреспондентыTableAdapter.Fill(ds11.Корреспонденты);
                ////}

                //////Если флаг = true значит год в app.config 2012 или более поздний

                ////if (flag == true)
                ////{
                ////    Classess.Корреспонденты корреспонденты = new RegKor.Classess.Корреспонденты();
                ////    DataSet dsКорреспонденты = корреспонденты.ЗаполнитьКорреспонденты_DataSet();

                ////    foreach (DataRow rowКорреспонденты in dsКорреспонденты.Tables[0].Rows)
                ////    {
                ////        DataRow row1 = ds11.Корреспонденты.NewRow();
                ////        row1[0] = rowКорреспонденты[0];
                ////        row1[1] = rowКорреспонденты[1];
                ////        ds11.Корреспонденты.Rows.Add(row1);
                ////    }
                ////}
                ////==========================================

                DS1TableAdapters.ПолучателиTableAdapter получателиTableAdapter = new RegKor.DS1TableAdapters.ПолучателиTableAdapter();
                получателиTableAdapter.Fill(ds11.Получатели);

                //Выполним в единой транзакции
                ПодключитьБД бд = new ПодключитьБД();
                FillDataSet fillDataSet = new FillDataSet();

                using (SqlConnection con = new SqlConnection(бд.СтрокаПодключения()))
                {
                    con.Open();
                    SqlTransaction transaction = con.BeginTransaction("transactLoad");

                    //Загрузим данные из БД на 1 декабря предыдущего года (1.01.2012 если программа настроена на 2013 год)
                    //DS1TableAdapters.КарточкаTableAdapter карточкаTableAdapter = new RegKor.DS1TableAdapters.КарточкаTableAdapter();
                    //карточкаTableAdapter.Fill(ds11.Карточка);
                    string queryКарточка = "select * from dbo.Карточка where ДатаПоступ >= '" + выбраннаяДата + "' and ДатаПоступ <= '" + следующаяДата + "' ";
                    fillDataSet.FillTable(queryКарточка, ds11, "Карточка", con, transaction);

                    DataTable tabTest = ds11.Карточка;

                    //Заполним таблицу Выборка
                    //DS1TableAdapters.ВыборкаTableAdapter выборкаTableAdapter = new RegKor.DS1TableAdapters.ВыборкаTableAdapter();
                    //выборкаTableAdapter.Fill(ds11.Выборка);
                    string queryВыборка = "select * from Выборка where ДатаПоступ >= '" + выбраннаяДата + "'  and ДатаПоступ <= '" + следующаяДата + "' ";
                    fillDataSet.FillTable(queryВыборка, ds11, "Выборка", con, transaction);

                    //Заполним КарточкаИсходящая
                    string queryКарточкаИсходящая = "select * from КарточкаИсходящая where Дата >= '" + выбраннаяДата + "' ";
                    fillDataSet.FillTable(queryКарточкаИсходящая, ds11, "КарточкаИсходящая", con, transaction);

                    //Заполним карточку ВыборкаИсходящихДокументов
                    string queryВыборкаИсходящихДокументов = "select * from ВыборкаИсходящихДокументов where Дата >= '" + выбраннаяДата + "' ";
                    fillDataSet.FillTable(queryВыборкаИсходящихДокументов, ds11, "ВыборкаИсходящихДокументов", con, transaction);

                }
                

                //DS1TableAdapters.ВыборкаКоличествоИсходящихДокументовTableAdapter выборкаИсходящихTableAdapter = new RegKor.DS1TableAdapters.ВыборкаКоличествоИсходящихДокументовTableAdapter();
                //выборкаИсходящихTableAdapter.Fill(ds11.ВыборкаКоличествоИсходящихДокументов);
                


                // исходящие
                DS1TableAdapters.ПодразделенияКомитетаTableAdapter подразделенияКомитетаTableAdapter = new RegKor.DS1TableAdapters.ПодразделенияКомитетаTableAdapter();
                подразделенияКомитетаTableAdapter.Fill(ds11.ПодразделенияКомитета);

                //Данные таблицы заполнены выше
                //DS1TableAdapters.КарточкаИсходящаяTableAdapter карточкаИсходящаяTableAdapter = new RegKor.DS1TableAdapters.КарточкаИсходящаяTableAdapter();
                //карточкаИсходящаяTableAdapter.Fill(ds11.КарточкаИсходящая);

                //DS1TableAdapters.ВыборкаИсходящихДокументовTableAdapter выборкаИсходящихДокументовTableAdapter = new RegKor.DS1TableAdapters.ВыборкаИсходящихДокументовTableAdapter();
                //выборкаИсходящихДокументовTableAdapter.Fill(ds11.ВыборкаИсходящихДокументов);

                this.Refresh();

                //входящие
                dataViewВыборкаРабДокументы.Table = ds11.Выборка;
                dataViewВыборкаРабДокументы.RowFilter = "ВДело=False AND ДатаПоступ >='01.12." + выбранныйГод + "'";
                dataGridРабочиеДокументы.DataSource = dataViewВыборкаРабДокументы;

                dataViewВыборкаДокументыВДеле.Table = ds11.Выборка;
                dataViewВыборкаДокументыВДеле.RowFilter = "ВДело=True AND ДатаПоступ >='01.12." + выбранныйГод + "'";
                dataGridДокументыВДеле.DataSource = dataViewВыборкаДокументыВДеле;
                //исходящие
                dataViewИсходящиеДокументы.Table = ds11.ВыборкаИсходящихДокументов;
                dataGridИсходящиеДокументы.DataSource = dataViewИсходящиеДокументы;

                //Заполним ComboBoxКорреспонденты
                this.comboBoxКорреспонденты.DataSource = ds11.Корреспонденты.Select("","ОписаниеКорреспондента");
                

                this.Refresh();

                Статистика();
                this.Refresh();

            }
            //catch (Exception exc)
            //{
            //    потокОжидания.Abort();
            //    MessageBox.Show(this, exc.Message + "\n" + exc.Source, "Метод \"ПодключитьсяПолучитьДанные()\"");
            //    this.Enabled = false;
            //    this.menuItem2.Enabled = false;
            //    this.menuItem3.Enabled = false;
            //    string str = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Environment.CurrentDirectory + "\\RegKor.exe").FileVersion;
            //    this.Text = "Регистрация корреспонденции. Версия: " + str + ". SQL Server: подключение не установлено";
            //}
            finally
            {
                потокОжидания.Abort();
            }
        }

        /// <summary>
        /// Делает апдэйт таблиц в базе, очищает таблицы в датасете
        /// и заполняет их заново из базы
        /// </summary>
        private void ОбновитьДанные()
        {
            потокОжидания = new System.Threading.Thread(new System.Threading.ThreadStart(ЗапуститьФормуОжидания));
            потокОжидания.Start();
            this.Refresh();
            try
            {
                //входящие
                DS1TableAdapters.ДокументыTableAdapter документыTableAdapter = new RegKor.DS1TableAdapters.ДокументыTableAdapter();
                документыTableAdapter.Update(ds11.Документы);

                DS1TableAdapters.КорреспондентыTableAdapter корреспондентыTableAdapter = new RegKor.DS1TableAdapters.КорреспондентыTableAdapter();
                корреспондентыTableAdapter.Update(ds11.Корреспонденты);

                DS1TableAdapters.ПолучателиTableAdapter получателиTableAdapter = new RegKor.DS1TableAdapters.ПолучателиTableAdapter();
                получателиTableAdapter.Update(ds11.Получатели);

                //DS1TableAdapters.КарточкаTableAdapter карточкаTableAdapter = new RegKor.DS1TableAdapters.КарточкаTableAdapter();
                //карточкаTableAdapter.Update(ds11.Карточка);
                //исходящие
                DS1TableAdapters.ПодразделенияКомитетаTableAdapter подразделенияКомитетаTableAdapter = new RegKor.DS1TableAdapters.ПодразделенияКомитетаTableAdapter();
                подразделенияКомитетаTableAdapter.Update(ds11.ПодразделенияКомитета);

                //DS1TableAdapters.КарточкаИсходящаяTableAdapter карточкаИсходящаяTableAdapter = new RegKor.DS1TableAdapters.КарточкаИсходящаяTableAdapter();
                //карточкаИсходящаяTableAdapter.Update(ds11.КарточкаИсходящая);
                this.Refresh();

                //входящие
                this.ds11.Карточка.Clear();
                this.ds11.Выборка.Clear();
                this.ds11.Документы.Clear();
                this.ds11.Корреспонденты.Clear();
                this.ds11.Получатели.Clear();
                //исходящие
                this.ds11.ВыборкаИсходящихДокументов.Clear();
                this.ds11.КарточкаИсходящая.Clear();
                this.ds11.ПодразделенияКомитета.Clear();

                //входящие
                документыTableAdapter = new RegKor.DS1TableAdapters.ДокументыTableAdapter();
                документыTableAdapter.Fill(ds11.Документы);
                корреспондентыTableAdapter = new RegKor.DS1TableAdapters.КорреспондентыTableAdapter();
                корреспондентыTableAdapter.Fill(ds11.Корреспонденты);
                получателиTableAdapter = new RegKor.DS1TableAdapters.ПолучателиTableAdapter();
                получателиTableAdapter.Fill(ds11.Получатели);
                //исходящие
                подразделенияКомитетаTableAdapter = new RegKor.DS1TableAdapters.ПодразделенияКомитетаTableAdapter();
                подразделенияКомитетаTableAdapter.Fill(ds11.ПодразделенияКомитета);

                ПодключитьБД бд = new ПодключитьБД();
                FillDataSet fillDataSet = new FillDataSet();

                using (SqlConnection con = new SqlConnection(бд.СтрокаПодключения()))
                {
                    StringBuilder builder = new StringBuilder();

                    con.Open();
                    SqlTransaction transaction = con.BeginTransaction("updateTransaction");
                    //Заполним карточку
                    //карточкаTableAdapter = new RegKor.DS1TableAdapters.КарточкаTableAdapter();
                    //карточкаTableAdapter.Fill(ds11.Карточка);
                    string queryКарточка = "select * from dbo.Карточка where ДатаПоступ >= '" + выбраннаяДата + "' and ДатаПоступ <= '" + следующаяДата + "'  ";
                    fillDataSet.FillTable(queryКарточка, ds11, "Карточка", con, transaction);

                    //Заполним таблицу Выборка
                    //DS1TableAdapters.ВыборкаTableAdapter выборкаTableAdapter = new RegKor.DS1TableAdapters.ВыборкаTableAdapter();
                    //выборкаTableAdapter.Fill(ds11.Выборка);
                    string queryВыборка = "select * from Выборка where ДатаПоступ >= '" + выбраннаяДата + "'";
                    fillDataSet.FillTable(queryВыборка, ds11, "Выборка", con, transaction);

                    //исходящие

                    //Заполним карточка Исходящая
                    //карточкаИсходящаяTableAdapter = new RegKor.DS1TableAdapters.КарточкаИсходящаяTableAdapter();
                    //карточкаИсходящаяTableAdapter.Fill(ds11.КарточкаИсходящая);
                    string queryКарточкаИсходящая = "select * from КарточкаИсходящая where Дата >= '" + выбраннаяДата + "'  and Дата <= '" + следующаяДата + "' ";
                    fillDataSet.FillTable(queryКарточкаИсходящая, ds11, "КарточкаИсходящая", con, transaction);

                    //Заполним таблицу ВыборкаИсходящихДокументов
                    //DS1TableAdapters.ВыборкаИсходящихДокументовTableAdapter выборкаИсходящихДокументовTableAdapter = new RegKor.DS1TableAdapters.ВыборкаИсходящихДокументовTableAdapter();
                    //выборкаИсходящихДокументовTableAdapter.Fill(ds11.ВыборкаИсходящихДокументов);
                    string queryВыборкаИсходящихДокументов = "select * from ВыборкаИсходящихДокументов where Дата >= '" + выбраннаяДата + "'";
                    fillDataSet.FillTable(queryВыборкаИсходящихДокументов, ds11, "ВыборкаИсходящихДокументов", con, transaction);

                    // Ниже изложенный Бред не нужен, но пока оставим вдруг к ниму придётся вернуться, так же здесь не хватает метода на обновление данных (но он скорее всего не нужен).
                    //// Попробуем выполнить запрос по кдалению данных.
                    //builder.Append(queryКарточка);
                    //builder.Append(queryВыборка);
                    //builder.Append(queryКарточкаИсходящая);
                    //builder.Append(queryВыборкаИсходящихДокументов);

                    //SqlCommand comDel = new SqlCommand(builder.ToString().Trim(), con);
                    //comDel.Transaction = transaction;

                    //comDel.ExecuteNonQuery();

                }

                dataViewВыборкаРабДокументы = new DataView(ds11.Выборка);
                dataViewВыборкаДокументыВДеле = new DataView(ds11.Выборка);
                dataViewИсходящиеДокументы = new DataView(ds11.ВыборкаИсходящихДокументов);
                this.Refresh();

                dataGridРабочиеДокументы.DataSource = null;
                dataGridДокументыВДеле.DataSource = null;
                dataGridИсходящиеДокументы.DataSource = null;

                dataViewВыборкаРабДокументы.Table = ds11.Выборка;
                dataViewВыборкаРабДокументы.RowFilter = "ВДело=False AND ДатаПоступ >='01.12." + выбранныйГод + "'";
                dataGridРабочиеДокументы.DataSource = dataViewВыборкаРабДокументы;

                dataViewВыборкаДокументыВДеле.Table = ds11.Выборка;
                dataViewВыборкаДокументыВДеле.RowFilter = "ВДело=True AND ДатаПоступ >='01.12." + выбранныйГод + "'";
                dataGridДокументыВДеле.DataSource = dataViewВыборкаДокументыВДеле;

                dataViewИсходящиеДокументы.Table = ds11.ВыборкаИсходящихДокументов;
                dataViewИсходящиеДокументы.RowFilter = "Дата >='01.12." + выбранныйГод + "'";
                dataGridИсходящиеДокументы.DataSource = dataViewИсходящиеДокументы;
                this.Refresh();

                Статистика();
                this.Refresh();
            }
            catch (Exception exc)
            {
                MessageBox.Show("" + exc.InnerException + "\n" + exc.Message + "\n" + exc.Source);
                Dispose(true);
            }
            finally
            {
                потокОжидания.Abort();
                this.Refresh();
            }

        }

        /// <summary>
        /// Осуществляет статистический расчет по базе 
        /// и отображение результатов на информационном лэйбле
        /// </summary>
        private void Статистика()
        {
            // входящие
            //int общееКолво = 0;
            string общееКолво = "0";
            string всегоВходящихДокументов = "0";
            string всегоПоставленныхНаКонтроль = "0";
            string всегоИсполненоКотрольныхДокументов = "0";
            string всегоИсходящихДокументов = "0";

            int вДеле = 0;
            int ожидающихРассмотрения = 0;
            int наКонтроле = 0;

            ПодключитьБД sConnect = new ПодключитьБД();
            string sConn = sConnect.СтрокаПодключения();

            using (SqlConnection con = new SqlConnection(sConn))
            {
                con.Open();

                // Новая реализация.
                Statistic statistic = new Statistic(selectedYear);
                DataTable tab = statistic.ВсегоДокументов(con);

                общееКолво = tab.Rows[0][0].ToString();

                // Всего входящих документов.
                всегоВходящихДокументов = statistic.ВсегоВходящихДокументов(con).Rows[0][0].ToString();

                // Всего документов поставленных на контроль.
                всегоПоставленныхНаКонтроль = statistic.ВсегоДокументовПоставленныхНаКонтроль(con).Rows[0][0].ToString();

                // Всего исполненно контрольных документов.
                всегоИсполненоКотрольныхДокументов = statistic.ВсегоИсполненныхДокументовПоставленныхНаКонтроль(con).Rows[0][0].ToString();

                // Всего исходящих документов.
                всегоИсходящихДокументов = statistic.ВсегоИсходящихДокументов(con).Rows[0][0].ToString(); ;

            }
            

            //Отсавим пока старую реализацию вдруг понадобиться.


            //DataRow[] rows = ds11.Выборка.Select("ДатаПоступ >='01.12." + выбранныйГод + "'");
            //общееКолво = rows.Length;

            //rows = ds11.Выборка.Select("ВДело=True AND ДатаПоступ >='01.12." + выбранныйГод + "'");
            //вДеле = rows.Length;

            //rows = ds11.Выборка.Select("ВДело=False AND ДатаПоступ >='01.12." + выбранныйГод + "'");
            //ожидающихРассмотрения = rows.Length;

            //rows = ds11.Выборка.Select("НаКонтроле=True AND ДатаПоступ >='01.12." + выбранныйГод + "'");
            //наКонтроле = rows.Length;

            //string инфо = "Общее количество документов в базе: " + общееКолво + "\n" +
            //                     "Документов списанных в дело: " + вДеле + "\n" +
            //                     "Документов ожидающих рассмотрения : " + ожидающихРассмотрения + "\n" +
            //                     "Документов стоящих на контроле : " + наКонтроле;

            string инфо = "Всего документов в базе: " + общееКолво + "\n" +
                                 "Всего входящих документов: " + всегоВходящихДокументов + "\n" +
                                 "Всего документов поставленных на контроль : " + всегоПоставленныхНаКонтроль + "\n" +
                                 "Всего исполнено документов поставленных на контроль : " + всегоИсполненоКотрольныхДокументов;


            labelИнфоTab1.Text = инфо;
            labelИнфоTab2.Text = инфо;

            // исходящие
            //общееКолво = 0;
            //rows = ds11.ВыборкаИсходящихДокументов.Select("Дата >='01.12." + выбранныйГод + "'");
            //общееКолво = rows.Length;
            //labelИнфоTab3.Text = "Исходящих документов в базе: " + общееКолво + "\n";
            labelИнфоTab3.Text = "Исходящих документов в базе: " + всегоИсходящихДокументов + "\n";
        }

        /// <summary>
        /// Помещает пару имя-значение в контейнер ParameterFields для отчета Crystal Reports
        /// </summary>
        /// <param name="paramName">имя параметра</param>
        /// <param name="paramValue">string значение параметра</param>
        /// <param name="paramFields">string контейнер параметров</param>
        public static void ПараметрыДляОтчета(string paramName,
            string paramValue,
            ParameterFields paramFields)
        {
            ParameterField paramField = new ParameterField();// параметр
            ParameterDiscreteValue paramDiscreteValue = new ParameterDiscreteValue();
            ParameterValues paramValues = new ParameterValues();
            // Устанавливаем имя параметра
            paramField.ParameterFieldName = paramName;// имя параметра
            // Устанавливаем значение параметра
            paramDiscreteValue.Value = paramValue;
            paramValues.Add(paramDiscreteValue);
            paramField.CurrentValues = paramValues;
            // Добавляем параметр в переданный контейнер
            paramFields.Add(paramField);
        }

        /// <summary>
        /// Отображает форму с отчетом "Печать контрольной карточки"
        /// </summary>
        private void ПечатьКарточки()
        {
            потокОжидания = new System.Threading.Thread(new System.Threading.ThreadStart(ЗапуститьФормуОжидания));
            потокОжидания.Start();

            DataGrid datagrid = new DataGrid();

            if (dataGridРабочиеДокументы.CanSelect)
            {
                datagrid = dataGridРабочиеДокументы;
            }
            if (dataGridДокументыВДеле.CanSelect)
            {
                datagrid = dataGridДокументыВДеле;
            }

            if (datagrid.CurrentCell.RowNumber == -1)
            {
                return;
            }

            int idТекущейКарточки = this.IDТекущейКарточки;

            FormView формаОтчета = new FormView();
            // Главная форма не активна:
            this.Enabled = false;

            try
            {
                // ReportDocument содержит свойства и методы для загрузки отчета:
                ReportDocument rptDoc = new ReportDocument();
                // загружает файл отчета:
                string fileName = @"..\report\Card.rpt";
                // файл отчета:
                rptDoc.Load(fileName);
                // источник данных:
                rptDoc.SetDataSource(ds11);
                // просмотрщику передаёт источник отчета и параметры к нему:
                формаОтчета.reportViewer.ReportSource = rptDoc;
                // Передаем параметры в отчет:
                ПараметрыДляОтчета("id_card", Convert.ToString(idТекущейКарточки), формаОтчета.reportViewer.ParameterFieldInfo);
                // показываем форму:
                потокОжидания.Abort();
                формаОтчета.ShowDialog(this);
            }
            catch (System.IndexOutOfRangeException exc)
            {
                MessageBox.Show("Нет записей для печати. \n" + exc.StackTrace);
                return;
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, "Произошла ошибка при открытии файла отчета \"Печать карточки документа\".\n" + exc.Message + "\n" + exc.InnerException, "Ошибка открытия файла отчета");
                return;
            }
            finally
            {
                потокОжидания.Abort();
                this.Enabled = true;
            }
        }

        /// <summary>
        /// Отображает форму с отчетом "Документы с истекшим сроком исполнения"
        /// </summary>
        private void ПечатьПросроченныхДокументов()
        {
            FormView формаОтчета = new FormView();
            this.Enabled = false;
            try
            {
                // ReportDocument содержит свойства и методы для загрузки отчета:
                ReportDocument rptDoc = new ReportDocument();
                // загружает файл отчета:
                string fileName = @"..\report\ExpiredDoc.rpt";
                // файл отчета:
                rptDoc.Load(fileName);
                // источник данных:
                rptDoc.SetDataSource(ds11);
                // просмотрщику передаёт источник отчета и параметры к нему:
                формаОтчета.reportViewer.ReportSource = rptDoc;
                // окно появляеться развернутое во весь экран:
                формаОтчета.WindowState = FormWindowState.Maximized;
                // отключаем окно ожидания
                потокОжидания.Abort();
                // показываем форму:
                формаОтчета.ShowDialog(this);
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, "Произошла ошибка при открытии файла отчета \"Документы с истекшими сроками исполнения\".\n" + exc.Message, "Ошибка открытия файла отчета");
                return;
            }
            finally
            {
                потокОжидания.Abort();
                this.Enabled = true;
            }
        }

        /// <summary>
        /// Отпарвляет в Word перечень просроченных документов.
        /// </summary>
        /// <param name="tab">Выборка</param>
        /// <param name="tabP">ВыборкаПовтор</param>
        //private void ПечатьПросроченныхДокументов(DataTable tab)//
        private void ПечатьПросроченныхДокументов(List<ПросроченныеДокументы> list)
        {
            //Создаём новый Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //app.Documents.Add(("ИстекшиеСроки.doc");

            string filName = Environment.CurrentDirectory + @"\Шаблон\Документы с истекшими сроками исполнения на.doc";

            //Загружаем документ
                        Microsoft.Office.Interop.Word.Document doc = null;

                        object fileName = filName;
                        object falseValue = false;
                        object trueValue = true;
                        object missing = Type.Missing;
                        object writePasswordDocument = "12A86Asd";

                        doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref writePasswordDocument,
            ref missing, ref missing, ref missing, ref missing, ref trueValue,
            ref missing, ref missing, ref missing);

            ////Дата начало отчёта.
            object wdrepl = Word.WdReplace.wdReplaceAll;
            //object searchtxt = "GreetingLine";
            object searchtxt = "date";
            object newtxt = (object)DateTime.Today.ToShortDateString();
            //object frwd = true;
            object frwd = false;
            doc.Content.Find.Execute(ref searchtxt, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd, ref missing, ref missing, ref newtxt, ref wdrepl, ref missing, ref missing,
            ref missing, ref missing);

            //Вставить таблицу
            object bookNaziv = "таблица";
            Word.Range wrdRng = doc.Bookmarks.get_Item(ref  bookNaziv).Range;

            object behavior = Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord8TableBehavior;
            object autobehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow;
            
            Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(wrdRng, 1, 5, ref behavior, ref autobehavior);
            table.Range.ParagraphFormat.SpaceAfter = 11;

            table.Columns[1].Width = 40;
            table.Columns[2].Width = 150;
            table.Columns[3].Width = 80;
            table.Columns[4].Width = 120;
            table.Columns[5].Width = 80;
            //table.Columns[6].Width = 120;

            table.Borders.Enable = 1; // Рамка - сплошная линия
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 9;

            //Запишем шапку таблицы.
            table.Cell(1, 1).Range.Text = "№ п/п";
            table.Cell(1, 2).Range.Text = "Ответственный исполнитель";
            table.Cell(1, 3).Range.Text = "Дата поступления";
            table.Cell(1, 4).Range.Text = "Номер входящий";
            table.Cell(1, 5).Range.Text = "Срок выполнения";

            Object beforeRow1 = Type.Missing;
            table.Rows.Add(ref beforeRow1);

            int count = 1;

            //Заполним таблицу данными.
            //foreach (DataRow row in tab.Rows)
            foreach(ПросроченныеДокументы item in list)
            {
                table.Cell(count + 1, 1).Range.Text = item.НомерПП.Trim(); // count.ToString().Trim();
                table.Cell(count + 1, 2).Range.Text = item.ОтветственныйИсполнитель.Trim(); // row["Резолюция"].ToString().Trim();
                table.Cell(count + 1, 3).Range.Text = item.ДатаПоступления.Trim(); //Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString();
                table.Cell(count + 1, 4).Range.Text = item.НомерВходящий.Trim();  //row["НомерВход"].ToString().Trim();
                table.Cell(count + 1, 5).Range.Text = item.СрокВыполнения.Trim(); //Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString();

                Object beforeRow2 = Type.Missing;
                table.Rows.Add(ref beforeRow2);

                count++;
            }


            ////Заполним таблицу данными.
            //foreach (DataRow row in tabP.Rows)
            //{
            //    table.Cell(count + 1, 1).Range.Text = count.ToString().Trim();
            //    table.Cell(count + 1, 2).Range.Text = row["Резолюция"].ToString().Trim();
            //    table.Cell(count + 1, 3).Range.Text = Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString();
            //    table.Cell(count + 1, 4).Range.Text = row["НомерВход"].ToString().Trim();
            //    table.Cell(count + 1, 5).Range.Text = Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString();

            //    Object beforeRow2 = Type.Missing;
            //    table.Rows.Add(ref beforeRow2);

            //    count++;
            //}

            //удалим последную строку
            table.Rows[count + 1].Delete();

            // Дата отчтёта.
            object wdrepl2 = Word.WdReplace.wdReplaceAll;
            //object searchtxt = "GreetingLine";
            object searchtxt2 = "countdoc";
            object newtxt2 = (object)list.Count;
            //object frwd = true;
            object frwd2 = false;
            doc.Content.Find.Execute(ref searchtxt2, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd2, ref missing, ref missing, ref newtxt2, ref wdrepl2, ref missing, ref missing,
            ref missing, ref missing);

           // Отобрпазим документ и закроем окно.
           app.Visible = true;

           //doc = Application.Documents["ИстекшиеСроки.doc"] as Word._Document;
           //doc.Close(ref doNotSaveChanges, ref missing, ref missing);
        }

        /// <summary>
        /// Отображает форму с отчетом "Контрольные уведомления"
        /// </summary>
        private void ПечатьКонтрольныхУведомлений()
        {
            //потокОжидания = new System.Threading.Thread(new System.Threading.ThreadStart(ЗапуститьФормуОжидания));
            //потокОжидания.Start();

            // Датасет для контрольных уведомлений:
            DSКонтрольныеУведомления dsУведомления = new DSКонтрольныеУведомления();

            // Завтрашняя дата:
            DateTime завтра = DateTime.Now.AddDays(1);

            // Старая реализация.
            // Получить всех исполнителей имеющих документы с истекающим сроком исполнения завтра или раньше
            //DataRow[] строкиПолучателей = ds11.Выборка.Select("СрокВыполнения<='" + завтра.Date + "' AND ДатаПоступ >='01.12.2011' AND НаКонтроле=True AND ВДело=False");

            Выборка выборка = new Выборка();
            //DataRow[] строкиПолучателей = выборка.ВыборкаДокументовИстекающимСроком().Select("СрокВыполнения<='" + завтра.Date + "' AND ДатаПоступ >='01.12.2011' AND НаКонтроле=True AND ВДело=False");
            DataRow[] строкиПолучателей = выборка.ВыборкаВсегоПолучателей();//.Select("ДатаПоступ >='01.12.2017' AND НаКонтроле=True AND ВДело=False");
           

            // Массив под имена получателей
            System.Collections.ArrayList списокПолучателей = new ArrayList();

            // Заполняем массив именами
            foreach (DataRow row in строкиПолучателей)
            {
                //string строкаИмен = (string)row["Резолюция"];
                string строкаИмен = (string)row["ОписаниеПолучателя"];
                string[] массивИмен = строкаИмен.Split(',');
                foreach (string имя in массивИмен)
                {
                    if (!списокПолучателей.Contains(имя.Trim()))
                    {
                        списокПолучателей.Add(имя.Trim());
                    }
                }
            }

            int CountDocumentControl = 0;

            StatisticControlNotific statistic = new StatisticControlNotific();

            // Получим всего документов на контроле.
            statistic.ВсегоДокументыНаКонтроле = выборка.ВыборкаДокументовИстекающимСроком().Select("НаКонтроле=True AND ВДело=False").Length;

            // Получим всего документов на контроле с истекшим сроком.
            
            statistic.КоличествоПросроченныхДокументов = выборка.ВыборкаДокументовИстекающимСроком().Select("СрокВыполнения<'" + DateTime.Now.Date + "' AND НаКонтроле=True AND ВДело=False").Length;

           
            foreach (Object имя in списокПолучателей)
            {
                // Получим данныепокаждому исполнителюдокументов.
                PersonDocument pd = new PersonDocument();

                pd.FioPerson = имя.ToString().Trim();

                // Количество документов на контроле.
                pd.ВсегоДокументыНаКонтроле = выборка.ВыборкаДокументовНаКонтроле(имя.ToString().Trim()).Tables[0].Rows.Count;

                // Количество просроченных документов.
                pd.ПросроченныеДокументы = выборка.ВыборкаПросроченныеДокументы(имя.ToString().Trim()).Tables[0].Select();

                // Количество просроченных документов.
                pd.КоличествоПросроченныхДокументов = pd.ПросроченныеДокументы.Length;

                // Количество не просроченных документов.
                pd.НеПрсороченныеДокументы = выборка.ВыборкаНеПросроченныеДокументы(имя.ToString().Trim()).Tables[0].Select();

                pd.КоличествоНеПросроченныхДокументов = pd.НеПрсороченныеДокументы.Length;

                // Таблица с просроченными документами.
                //DataRow[] dtOverDoc = выборка.ВыборкаДокументовИстекающимСроком().Rows;//имя.ToString().Trim());//.Select("СрокВыполнения<'" + DateTime.Now.Date + "' AND ОписаниеПолучателя LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");

                //// Получим список просроченных документов.
                //pd.ПросроченныеДокументы = dtOverDoc;

                //// Всего документов на контроле.
                //pd.ВсегоДокументыНаКонтроле = выборка.ВыборкаДокументовИстекающимСроком().Select("ОписаниеПолучателя LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False").Length;

                //// Список документов на контроле.
                //pd.ДокументыНаКонтроле = выборка.ВыборкаДокументовИстекающимСроком().Select("ОписаниеПолучателя LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");

                //// Количество просроченных документов для текущего исполнитьеля.
                //pd.КоличествоПросроченныхДокументов = dtOverDoc.Length;

               

                //// Таблица с не просроченными документами.
                ////DataRow[] dtNotOverDoc = выборка.ВыборкаДокументовИстекающимСроком().Select("СрокВыполнения >='" + завтра.Date + "' and СрокВыполнения<'" + DateTime.Now.Date + "' AND ОписаниеПолучателя LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");

                //DataRow[] dtNotOverDoc = выборка.ВыборкаДокументовИстекающимСроком().Select("СрокВыполнения >'" + DateTime.Now.Date + "' AND ОписаниеПолучателя LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");

                //// Запишем не просроченные документы.
                //pd.НеПрсороченныеДокументы = dtNotOverDoc;

                //// Запишем количество не просроченных документов.
                //pd.КоличествоНеПросроченныхДокументов = dtNotOverDoc.Length;

                statistic.СписокИсполнителей.Add(pd);
            }

            string iTest = "";

            /*
            foreach (Object имя in списокПолучателей)
            {

                // Документы на контроле для текущего получателя:
                //DataRow[] общее = ds11.Выборка.Select("Резолюция LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");
                DataRow[] общее = выборка.ВыборкаДокументовИстекающимСроком().Select("ОписаниеПолучателя LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");

                CountDocumentControl += общее.Length;

                // Просроченные документы для текущего получателя:
                //DataRow[] просроченные = ds11.Выборка.Select("СрокВыполнения<'" + DateTime.Now.Date + "' AND Резолюция LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");
                DataRow[] просроченные = выборка.ВыборкаДокументовИстекающимСроком().Select("СрокВыполнения<'" + DateTime.Now.Date + "' AND ОписаниеПолучателя LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");
                
                
                // Документы с 1 днем для текущего получателя:
                //DataRow[] с1днём = ds11.Выборка.Select("СрокВыполнения='" + завтра.Date + "' AND Резолюция LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");
                DataRow[] с1днём = выборка.ВыборкаДокументовИстекающимСроком().Select("СрокВыполнения='" + завтра.Date + "' AND ОписаниеПолучателя LIKE '%" + имя + "%' AND НаКонтроле=True AND ВДело=False");

                // Добавляем строку в таблицу "получатели":
                dsУведомления.Получатели.AddПолучателиRow(имя.ToString(), общее.Length, с1днём.Length, просроченные.Length);

                // Узнаем ид текущего получателя:
                DataRow[] получатель = dsУведомления.Получатели.Select("ОписаниеПолучателя='" + имя.ToString() + "'");
                int idПолучателя = (int)получатель[0]["id_Получателя"];

                // Порядковый номер документа для текущего получателя:
                int НомерПП = 1;

                // Тип документа:
                int тип = 0;// с1днем = 0, просроченные = 1

                //Добавляем документы с одним днем в таблицу "Документы"
                foreach (DataRow документ in с1днём)
                {
                    DateTime датаПоступления = (DateTime)документ["ДатаПоступ"];
                    string номерВходящий = (string)документ["НомерВход"];
                    DateTime датаКонтроля = (DateTime)документ["СрокВыполнения"];
                    dsУведомления.Документы.AddДокументыRow(
                                                            idПолучателя,
                                                            датаПоступления,
                                                            номерВходящий,
                                                            датаКонтроля,
                                                            НомерПП,
                                                            тип
                                                            );
                    НомерПП++;
                }

                //Добавляем документы с одним днем в таблицу "Документы"
                НомерПП = 1;
                тип = 1;
                foreach (DataRow документ in просроченные)
                {
                    DateTime датаПоступления = (DateTime)документ["ДатаПоступ"];
                    string номерВходящий = (string)документ["НомерВход"];
                    DateTime датаКонтроля = (DateTime)документ["СрокВыполнения"];
                    dsУведомления.Документы.AddДокументыRow(idПолучателя,
                                                            датаПоступления,
                                                            номерВходящий,
                                                            датаКонтроля,
                                                            НомерПП,
                                                            тип);
                    НомерПП++;
                }
            }


            */

            FormPrintКонтрольноеУведомление formPrint = new FormPrintКонтрольноеУведомление();
            
            // Передадим данныедля отчета.
            //formPrint.DataSetForm = dsУведомления;
            formPrint.DataStatistic = statistic;

            // Передадим количество документов на контроле.
            formPrint.CountDocControl = CountDocumentControl;
            formPrint.Show();


            //string sTest = "";

            //FormКонтрольныеУведомления формаОтчета = new FormКонтрольныеУведомления(dsУведомления);
            //// Отключить гл. форму:
            //this.Enabled = false;

            //try
            //{
            //    // ReportDocument содержит свойства и методы для загрузки отчета:
            //    ReportDocument rptDoc = new ReportDocument();
            //    // загружает файл отчета:
            //    string fileName = @"..\report\KontrolMessage.rpt";
            //    // файл отчета:
            //    rptDoc.Load(fileName);
            //    // источник данных:
            //    rptDoc.SetDataSource(dsУведомления);
            //    // просмотрщику передаёт источник отчета и параметры к нему:
            //    формаОтчета.reportViewer.ReportSource = rptDoc;
            //    // Закрываем окно ожидания:
            //    потокОжидания.Abort();
            //    // Показываем форму:
            //    формаОтчета.ShowDialog(this);
            //}
            //catch (System.Exception exc)
            //{
            //    потокОжидания.Abort();
            //    MessageBox.Show("Ошибка отчета \"Контрольные уведомления\". \n" + exc.Message + "\n" + exc.StackTrace);
            //    return;
            //}
            //finally
            //{
            //    потокОжидания.Abort();
            //    this.Enabled = true;
            //}
        }

        private void ЗапуститьФормуОжидания()
        {
            FormОжидание form = new FormОжидание();
            form.Left = (this.Left) + this.Width / 2 - (form.Width / 2);
            form.Top = (this.Top) + this.Height / 2 - (form.Height / 2);
            form.TopMost = true;
            form.ShowDialog();
        }

        /// <summary>
        /// Проверяет, запущена копия программы или нет
        /// </summary>
        /// <returns>true если программа запущена, иначе false</returns>
        static bool ПрограммаУжеЗапущена()
        {
            bool createdNew;
            mutex = new System.Threading.Mutex(false, "RegKorMutex", out createdNew);
            return !createdNew;
        }


        #endregion

        #region Свойства
        private int IDТекущейКарточки
        {
            get
            {

                DataGrid datagrid = new DataGrid();

                if (dataGridРабочиеДокументы.CanSelect)
                {
                    datagrid = dataGridРабочиеДокументы;
                }
                if (dataGridДокументыВДеле.CanSelect)
                {
                    datagrid = dataGridДокументыВДеле;
                }
                if (dataGridИсходящиеДокументы.CanSelect)
                {
                    datagrid = dataGridИсходящиеДокументы;
                }
                // получаем данные отображаемые в выделенной строке:
                BindingManagerBase bmb = this.BindingContext[datagrid.DataSource, datagrid.DataMember];
                bmb.Position = datagrid.CurrentCell.RowNumber;
                datagrid.Select(datagrid.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                return (int)drv["id_карточки"];
            }
        }

        #endregion

        #region События

        /// <summary>
        /// Двойной щелчок в ячейке таблицы "РабочиеДокументы"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridРабочиеДокументы_DoubleClick(object sender, EventArgs e)
        {
            //// получаем область в которой щелкнули мышкой:
            //DataGrid.HitTestInfo myHitTest = dataGridРабочиеДокументы.HitTest(мышь.X, мышь.Y);

            //if (myHitTest.Type == DataGrid.HitTestType.Cell)// если щелкнули в ячейке, без разницы какой кнопкой
            //{
            //    FormКарточка form = new FormКарточка(ds11, IDТекущейКарточки, выбранныйГод);
            //    form.ShowDialog(this);
            //    if (form.DialogResult == DialogResult.OK)
            //    {
            //        DS1TableAdapters.КарточкаTableAdapter адаптер = new RegKor.DS1TableAdapters.КарточкаTableAdapter();
            //        адаптер.Update(form.строкаКарточки);
            //        ОбновитьДанные();
            //    }
            //}

        }

        /// <summary>
        /// Двойной щелчок в таблице "ДокументыВДеле"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridДокументыВДеле_DoubleClick(object sender, EventArgs e)
        {
            //// получаем область в которой щелкнули мышкой:
            //DataGrid.HitTestInfo myHitTest = dataGridДокументыВДеле.HitTest(мышь.X, мышь.Y);

            //if (myHitTest.Type == DataGrid.HitTestType.Cell)// если щелкнули в ячейке, без разницы какой кнопкой
            //{
            //    FormКарточка form = new FormКарточка(ds11, IDТекущейКарточки, выбранныйГод);
            //    form.ShowDialog(this);
            //    if (form.DialogResult == DialogResult.OK)
            //    {
            //        DS1TableAdapters.КарточкаTableAdapter адаптер = new RegKor.DS1TableAdapters.КарточкаTableAdapter();
            //        адаптер.Update(form.строкаКарточки);
            //        ОбновитьДанные();
            //    }
            //}

        }

        /// <summary>
        /// Двойной щелчок в таблице "ИсходящиеДокументы"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridИсходящиеДокументы_DoubleClick(object sender, EventArgs e)
        {
            /*
            // получаем область в которой щелкнули мышкой:
            DataGrid.HitTestInfo myHitTest = dataGridИсходящиеДокументы.HitTest(мышь.X, мышь.Y);

            if (myHitTest.Type == DataGrid.HitTestType.Cell)// если щелкнули в ячейке, без разницы какой кнопкой
            {
                потокОжидания = new System.Threading.Thread(new System.Threading.ThreadStart(ЗапуститьФормуОжидания));
                потокОжидания.Start();
                // получаем данные отображаемые в выделенной строке:
                BindingManagerBase bmb = this.BindingContext[dataGridИсходящиеДокументы.DataSource, dataGridИсходящиеДокументы.DataMember];
                bmb.Position = dataGridИсходящиеДокументы.CurrentCell.RowNumber;
                dataGridИсходящиеДокументы.Select(dataGridИсходящиеДокументы.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                DataRow[] row = ds11.КарточкаИсходящая.Select("id_карточки=" + (int)drv["id_карточки"]);
                DS1.КарточкаИсходящаяRow строкаДляИзменения = (DS1.КарточкаИсходящаяRow)row[0];
                FormКарточкаИсходящая form = new FormКарточкаИсходящая(ds11, строкаДляИзменения, выбранныйГод);
                потокОжидания.Abort();
                form.ShowDialog(this);
                if (form.DialogResult == DialogResult.OK)
                {
                    string sTest = "asd";

                    DS1.КарточкаИсходящаяRow rowTest = form.строкаИсходящейКарточки;


                    DS1TableAdapters.КарточкаИсходящаяTableAdapter адаптер = new RegKor.DS1TableAdapters.КарточкаИсходящаяTableAdapter();
                    адаптер.Update(form.строкаИсходящейКарточки);
                    ОбновитьДанные();
                }
            }*/

        }

        /// <summary>
        /// Щелчёк мыши в таблице РабочиеДокументы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridРабочиеДокументы_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            мышь.X = e.X;
            мышь.Y = e.Y;

            // получаем область в которой щелкнули мышкой:
            DataGrid.HitTestInfo myHitTest = dataGridРабочиеДокументы.HitTest(e.X, e.Y);

            if (e.Button == System.Windows.Forms.MouseButtons.Right)// если щелкнули правой кнопкой
            {
                // удаляем все пункты контекстного меню
                contextMenu1.MenuItems.Clear();

                // Создаем пункт меню Добавить посетителя:
                MenuItem menuItemContextДобавитьЗапись = new System.Windows.Forms.MenuItem("Добавить запись");
                menuItemContextДобавитьЗапись.Click += new EventHandler(menuItemContextДобавитьЗапись_Click);
                this.contextMenu1.MenuItems.Add(0, menuItemContextДобавитьЗапись);

                if (myHitTest.Type == DataGrid.HitTestType.Cell)// если щелкнули в ячейке
                {
                    // убираем текущее выделение:
                    dataGridРабочиеДокументы.UnSelect(dataGridРабочиеДокументы.CurrentRowIndex);
                    // делаем текущей ячейкой:
                    dataGridРабочиеДокументы.CurrentCell = new DataGridCell(myHitTest.Row, myHitTest.Column);
                    // выделяем эту строку:
                    dataGridРабочиеДокументы.Select(myHitTest.Row);

                    MenuItem menuItemContextИзменитьЗапись = new System.Windows.Forms.MenuItem("Изменить запись");
                    menuItemContextИзменитьЗапись.Click += new EventHandler(menuItemContextИзменитьЗапись_Click);
                    this.contextMenu1.MenuItems.Add(1, menuItemContextИзменитьЗапись);

                    MenuItem menuItemContextУдалитьЗапись = new System.Windows.Forms.MenuItem("Удалить запись");
                    menuItemContextУдалитьЗапись.Click += new EventHandler(menuItemContextУдалитьЗапись_Click);
                    this.contextMenu1.MenuItems.Add(2, menuItemContextУдалитьЗапись);

                    MenuItem menuItemContextПечатьКарточки = new System.Windows.Forms.MenuItem("Печать карточки");
                    menuItemContextПечатьКарточки.Click += new EventHandler(menuItemContextПечатьКарточки_Click);
                    this.contextMenu1.MenuItems.Add(3, menuItemContextПечатьКарточки);

                }

                // показываем созданное контекстное меню:
                contextMenu1.Show(dataGridРабочиеДокументы, new System.Drawing.Point(e.X, e.Y));
            }

            if (myHitTest.Type == DataGrid.HitTestType.Cell)// если щелкнули в ячейке, без разницы какой кнопкой
            {
                // получаем данные отображаемые в выделенной строке:
                BindingManagerBase bmb = this.BindingContext[dataGridРабочиеДокументы.DataSource, dataGridРабочиеДокументы.DataMember];
                bmb.Position = dataGridРабочиеДокументы.CurrentCell.RowNumber;
                dataGridРабочиеДокументы.Select(dataGridРабочиеДокументы.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                // выводим полученные данные на информационный лэйбл:РезультатВыполнения
                labelИнфоTab1.Text = "Документ: " + drv["ОписаниеДокумента"].ToString() + Environment.NewLine +
                    "Корр-т: " + drv["ОписаниеКорреспондента"].ToString() + Environment.NewLine +
                    "Дата отпр.: " + Convert.ToDateTime(drv["ДатаИсхода"]).ToShortDateString() +
                    "  №исход.: " + drv["НомерИсход"].ToString() + Environment.NewLine +
                    "Дата пост.: " + Convert.ToDateTime(drv["ДатаПоступ"]).ToShortDateString() +
                    "  №вход.: " + drv["НомерВход"].ToString() + Environment.NewLine +
                    "Содержание: " + drv["КраткоеСодержание"].ToString() + Environment.NewLine +
                    "Кому отписано: " + drv["Резолюция"].ToString();
            }
            else if (dataGridРабочиеДокументы.CurrentRowIndex > -1)
            {
                // убираем текущее выделение:
                dataGridРабочиеДокументы.UnSelect(dataGridРабочиеДокументы.CurrentRowIndex);
                Статистика();
            }
        }

        /// <summary>
        /// Шелчёк мыши в таблице "Входящие рабочие документы"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridДокументыВДеле_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            мышь.X = e.X;
            мышь.Y = e.Y;

            // получаем область в которой щелкнули мышкой:
            DataGrid.HitTestInfo myHitTest = dataGridДокументыВДеле.HitTest(e.X, e.Y);

            if (e.Button == System.Windows.Forms.MouseButtons.Right)// если щелкнули правой кнопкой
            {
                // удаляем все пункты контекстного меню
                contextMenu1.MenuItems.Clear();

                if (myHitTest.Type == DataGrid.HitTestType.Cell)// если щелкнули в ячейке
                {
                    // убираем текущее выделение:
                    dataGridДокументыВДеле.UnSelect(dataGridДокументыВДеле.CurrentRowIndex);
                    // делаем текущей ячейкой:
                    dataGridДокументыВДеле.CurrentCell = new DataGridCell(myHitTest.Row, myHitTest.Column);
                    // выделяем эту строку:
                    dataGridДокументыВДеле.Select(myHitTest.Row);

                    MenuItem menuItemContextИзменитьЗапись = new System.Windows.Forms.MenuItem("Изменить запись");
                    menuItemContextИзменитьЗапись.Click += new EventHandler(menuItemContextИзменитьЗапись_Click2);
                    this.contextMenu1.MenuItems.Add(0, menuItemContextИзменитьЗапись);

                    MenuItem menuItemContextУдалитьЗапись = new System.Windows.Forms.MenuItem("Удалить запись");
                    menuItemContextУдалитьЗапись.Click += new EventHandler(menuItemContextУдалитьЗапись_Click2);
                    this.contextMenu1.MenuItems.Add(1, menuItemContextУдалитьЗапись);

                    MenuItem menuItemContextПечатьКарточки = new System.Windows.Forms.MenuItem("Печать карточки");
                    menuItemContextПечатьКарточки.Click += new EventHandler(menuItemContextПечатьКарточки_Click);
                    this.contextMenu1.MenuItems.Add(2, menuItemContextПечатьКарточки);

                    MenuItem menuItemContextФильтр = new MenuItem("Отфильтровать по времени");
                    menuItemContextФильтр.Click += new EventHandler(menuItemContextФильтр_Click);
                    this.contextMenu1.MenuItems.Add(3, menuItemContextФильтр);

                    //MenuItem menuItemContextПовторитьВложениеДокумента = new MenuItem("Отфильтровать по времени");
                    //menuItemContextПовторитьВложениеДокумента.Click += new EventHandler(menuItemContextПовторитьВложениеДокумента_Click);
                    //this.contextMenu1.MenuItems.Add(3, menuItemContextПовторитьВложениеДокумента);


                }

                // показываем созданное контекстное меню:
                contextMenu1.Show(dataGridДокументыВДеле, new System.Drawing.Point(e.X, e.Y));
            }

            if (myHitTest.Type == DataGrid.HitTestType.Cell)// если щелкнули в ячейке, без разницы какой кнопкой
            {
                // получаем данные отображаемые в выделенной строке:
                BindingManagerBase bmb = this.BindingContext[dataGridДокументыВДеле.DataSource, dataGridДокументыВДеле.DataMember];
                bmb.Position = dataGridДокументыВДеле.CurrentCell.RowNumber;
                dataGridДокументыВДеле.Select(dataGridДокументыВДеле.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                // выводим полученные данные на информационный лэйбл:РезультатВыполнения
                labelИнфоTab2.Text = "Документ: " + drv["ОписаниеДокумента"].ToString() + Environment.NewLine +
                                    "Корр-т: " + drv["ОписаниеКорреспондента"].ToString() + Environment.NewLine +
                                    "Дата отпр.: " + Convert.ToDateTime(drv["ДатаИсхода"]).ToShortDateString() +
                                    "  №исход.: " + drv["НомерИсход"].ToString() + Environment.NewLine +
                                    "Дата пост.: " + Convert.ToDateTime(drv["ДатаПоступ"]).ToShortDateString() +
                                    "  №вход.: " + drv["НомерВход"].ToString() + Environment.NewLine +
                                    "Содержание: " + drv["КраткоеСодержание"].ToString() + Environment.NewLine +
                                    "Кому отписано: " + drv["Резолюция"].ToString() + Environment.NewLine +
                                    "Результат выполнения: " + drv["РезультатВыполнения"].ToString();
            }
            else if (dataGridДокументыВДеле.CurrentRowIndex > -1)
            {
                // убираем текущее выделение:
                dataGridДокументыВДеле.UnSelect(dataGridДокументыВДеле.CurrentRowIndex);
                Статистика();
            }
        }

        void menuItemContextПовторитьВложениеДокумента_Click(object sender, EventArgs e)
        {
           //// Получим id текущей карточки.
           //int idCard = this.DataGr
        }

        void menuItemContextФильтр_Click(object sender, EventArgs e)
        {
            FormSelectMonth fsm = new FormSelectMonth();
            fsm.ВыбранныйГод = выбранныйГод;
            fsm.ShowDialog();

            if (fsm.DialogResult == DialogResult.OK)
            {
                // Получим первый и последний день месяца.
                string датаНачало = fsm.GetПервыйДень;
                string датаКонец = fsm.GetКрайнийДень;

                if (датаНачало != null)
                {

                    dataGridДокументыВДеле.DataSource = null;
                    //string фильтр = "ВДело=True AND ДатаПоступ >='01.12.2011' AND (КраткоеСодержание LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    //    " OR ОписаниеДокумента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    //    " OR ОписаниеКорреспондента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    //    " OR РезультатВыполнения LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    //    " OR Резолюция LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    //    " OR НомерВход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    //    " OR НомерИсход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%')";

                    //========================

                    //Объявим переменную - фильтр
                    string фильтр;
                    if (this.comboBoxКорреспонденты.Visible == true)
                    {
                        фильтр = "ВДело=True AND (ДатаПоступ >='01.12.2011' AND ДатаПоступ >='" + датаНачало + "' AND ДатаПоступ <='" + датаКонец + "') AND (КраткоеСодержание LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR ОписаниеДокумента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR ОписаниеКорреспондента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR РезультатВыполнения LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR Резолюция LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR НомерВход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR НомерИсход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%')" +
                            " AND ОписаниеКорреспондента = '" + this.comboBoxКорреспонденты.Text + "'";
                        //" AND ОписаниеКорреспондента = 'ГКУ СО \"КСПН г.Саратова\"'";
                    }
                    else
                    {
                        фильтр = "ВДело=True AND (ДатаПоступ >='01.12.2011' AND ДатаПоступ >='" + датаНачало + "' AND ДатаПоступ <='" + датаКонец + "') AND (КраткоеСодержание LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR ОписаниеДокумента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR ОписаниеКорреспондента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR РезультатВыполнения LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR Резолюция LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR НомерВход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR НомерИсход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%')";
                    }

                    if (this.textBoxСтрокаПоискаTab2.Text == "")
                    {
                        фильтр = "ВДело=True AND (ДатаПоступ >='01.12.2011' AND ДатаПоступ >='" + датаНачало + "' AND ДатаПоступ <='" + датаКонец + "') AND (КраткоеСодержание LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR ОписаниеДокумента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR ОписаниеКорреспондента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR РезультатВыполнения LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR Резолюция LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR НомерВход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                            " OR НомерИсход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%')";
                    }
                    //=====================
                    dataViewВыборкаДокументыВДеле.RowFilter = фильтр;
                    dataGridДокументыВДеле.DataSource = dataViewВыборкаДокументыВДеле;
                    textBoxСтрокаПоискаTab2.Focus();
                    if (textBoxСтрокаПоискаTab2.Text != "")
                    {
                        labelОтобраноДокументовПоискомTab2.Text = "Отобрано документов: " + dataViewВыборкаДокументыВДеле.Count;
                    }
                    else
                    {
                        labelОтобраноДокументовПоискомTab2.Text = "";
                    }

                }

            }

        }

        /// <summary>
        /// Шелчёк мыши в таблице "Исходящие документы"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridИсходящиеДокументы_MouseUp(object sender, MouseEventArgs e)
        {
            мышь.X = e.X;
            мышь.Y = e.Y;

            // получаем область в которой щелкнули мышкой:
            DataGrid.HitTestInfo myHitTest = dataGridИсходящиеДокументы.HitTest(e.X, e.Y);

            if (e.Button == System.Windows.Forms.MouseButtons.Right)// если щелкнули правой кнопкой
            {
                // удаляем все пункты контекстного меню
                contextMenu1.MenuItems.Clear();

                // Создаем пункт меню Добавить посетителя:
                MenuItem menuItemContextДобавитьЗапись = new System.Windows.Forms.MenuItem("Добавить запись");
                menuItemContextДобавитьЗапись.Click += new EventHandler(menuItemContextДобавитьИсходящуюЗапись_Click);
                this.contextMenu1.MenuItems.Add(0, menuItemContextДобавитьЗапись);

                if (myHitTest.Type == DataGrid.HitTestType.Cell)// если щелкнули в ячейке
                {

                    // убираем текущее выделение:
                    dataGridИсходящиеДокументы.UnSelect(dataGridИсходящиеДокументы.CurrentRowIndex);
                    // делаем текущей ячейкой:
                    dataGridИсходящиеДокументы.CurrentCell = new DataGridCell(myHitTest.Row, myHitTest.Column);
                    // выделяем эту строку:
                    dataGridИсходящиеДокументы.Select(myHitTest.Row);

                    MenuItem menuItemContextИзменитьЗапись = new System.Windows.Forms.MenuItem("Изменить запись");
                    menuItemContextИзменитьЗапись.Click += new EventHandler(menuItemContextИзменитьИсходящуюЗапись_Click);
                    this.contextMenu1.MenuItems.Add(1, menuItemContextИзменитьЗапись);

                    MenuItem menuItemContextУдалитьЗапись = new System.Windows.Forms.MenuItem("Удалить запись");
                    menuItemContextУдалитьЗапись.Click += new EventHandler(menuItemContextУдалитьИсходящуюЗапись_Click);
                    this.contextMenu1.MenuItems.Add(2, menuItemContextУдалитьЗапись);

                }

                // показываем созданное контекстное меню:
                contextMenu1.Show(dataGridИсходящиеДокументы, new System.Drawing.Point(e.X, e.Y));
            }

            if (myHitTest.Type == DataGrid.HitTestType.Cell)// если щелкнули в ячейке, без разницы какой кнопкой
            {
                // убираем текущее выделение:
                dataGridИсходящиеДокументы.UnSelect(dataGridИсходящиеДокументы.CurrentRowIndex);
                // делаем текущей ячейкой:
                dataGridИсходящиеДокументы.CurrentCell = new DataGridCell(myHitTest.Row, myHitTest.Column);
                // выделяем эту строку:
                dataGridИсходящиеДокументы.Select(myHitTest.Row);

                // получаем данные отображаемые в выделенной строке:
                BindingManagerBase bmb = this.BindingContext[dataGridИсходящиеДокументы.DataSource, dataGridРабочиеДокументы.DataMember];
                bmb.Position = dataGridИсходящиеДокументы.CurrentCell.RowNumber;
                dataGridИсходящиеДокументы.Select(dataGridИсходящиеДокументы.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                // выводим полученные данные на информационный лэйбл:РезультатВыполнения
                labelИнфоTab3.Text = "Дата: " + Convert.ToDateTime(drv["Дата"]).ToShortDateString() + Environment.NewLine +
                    "Номер: " + drv["ТекстовыйНомер"].ToString() + Environment.NewLine +
                    "Адресат: " + drv["ОписаниеАдресата"].ToString() + Environment.NewLine +
                    "Отправитель: " + drv["ОписаниеПодразделения"].ToString() + Environment.NewLine +
                    "Содержание: " + drv["Содержание"].ToString() + Environment.NewLine +
                    "Ответ на документ: " + drv["НомерВходВходящегоДокумента"].ToString() + Environment.NewLine;
            }
            else if (dataGridИсходящиеДокументы.CurrentRowIndex > -1)
            {
                // убираем текущее выделение:
                dataGridИсходящиеДокументы.UnSelect(dataGridИсходящиеДокументы.CurrentRowIndex);
                Статистика();
            }
        }

        /// <summary>
        /// Изменение выделение в таблице Рабочие документы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridРабочиеДокументы_CurrentCellChanged(object sender, System.EventArgs e)
        {
            dataGridРабочиеДокументы.Select(dataGridРабочиеДокументы.CurrentCell.RowNumber);
            // получаем данные отображаемые в выделенной строке:
            BindingManagerBase bmb = this.BindingContext[dataGridРабочиеДокументы.DataSource, dataGridРабочиеДокументы.DataMember];
            bmb.Position = dataGridРабочиеДокументы.CurrentCell.RowNumber;
            dataGridРабочиеДокументы.Select(dataGridРабочиеДокументы.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            // выводим полученные данные на информационный лэйбл:РезультатВыполнения
            labelИнфоTab1.Text = "Документ: " + drv["ОписаниеДокумента"].ToString() + Environment.NewLine +
                "Корр-т: " + drv["ОписаниеКорреспондента"].ToString() + Environment.NewLine +
                "Дата отпр.: " + Convert.ToDateTime(drv["ДатаИсхода"]).ToShortDateString() +
                "  №исход.: " + drv["НомерИсход"].ToString() + Environment.NewLine +
                "Дата пост.: " + Convert.ToDateTime(drv["ДатаПоступ"]).ToShortDateString() +
                "  №вход.: " + drv["НомерВход"].ToString() + Environment.NewLine +
                "Содержание: " + drv["КраткоеСодержание"].ToString() + Environment.NewLine +
                "Кому отписано: " + drv["Резолюция"].ToString();
        }

        /// <summary>
        /// Изменение выделение в таблице Документы "в деле"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridДокументыВДеле_CurrentCellChanged(object sender, System.EventArgs e)
        {
            dataGridДокументыВДеле.Select(dataGridДокументыВДеле.CurrentCell.RowNumber);
            // получаем данные отображаемые в выделенной строке:
            BindingManagerBase bmb = this.BindingContext[dataGridДокументыВДеле.DataSource, dataGridДокументыВДеле.DataMember];
            bmb.Position = dataGridДокументыВДеле.CurrentCell.RowNumber;
            dataGridДокументыВДеле.Select(dataGridДокументыВДеле.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            // выводим полученные данные на информационный лэйбл:РезультатВыполнения
            labelИнфоTab2.Text = "Документ: " + drv["ОписаниеДокумента"].ToString() + Environment.NewLine +
                "Корр-т: " + drv["ОписаниеКорреспондента"].ToString() + Environment.NewLine +
                "Дата отпр.: " + Convert.ToDateTime(drv["ДатаИсхода"]).ToShortDateString() +
                "  №исход.: " + drv["НомерИсход"].ToString() + Environment.NewLine +
                "Дата пост.: " + Convert.ToDateTime(drv["ДатаПоступ"]).ToShortDateString() +
                "  №вход.: " + drv["НомерВход"].ToString() + Environment.NewLine +
                "Содержание: " + drv["КраткоеСодержание"].ToString() + Environment.NewLine +
                "Кому отписано: " + drv["Резолюция"].ToString() + Environment.NewLine +
                "Результат выполнения: " + drv["РезультатВыполнения"].ToString();
        }

        /// <summary>
        /// Изменение выделение в таблице Исходящие документы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGridИсходящиеДокументы_CurrentCellChanged(object sender, EventArgs e)
        {
            dataGridИсходящиеДокументы.Select(dataGridИсходящиеДокументы.CurrentCell.RowNumber);

            // получаем данные отображаемые в выделенной строке:
            BindingManagerBase bmb = this.BindingContext[dataGridИсходящиеДокументы.DataSource, dataGridРабочиеДокументы.DataMember];
            bmb.Position = dataGridИсходящиеДокументы.CurrentCell.RowNumber;
            dataGridИсходящиеДокументы.Select(dataGridИсходящиеДокументы.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            // выводим полученные данные на информационный лэйбл:РезультатВыполнения
            labelИнфоTab3.Text = "Дата: " + Convert.ToDateTime(drv["Дата"]).ToShortDateString() + Environment.NewLine +
                "Номер: " + drv["ТекстовыйНомер"].ToString() + Environment.NewLine +
                "Адресат: " + drv["ОписаниеАдресата"].ToString() + Environment.NewLine +
                "Отправитель: " + drv["ОписаниеПодразделения"].ToString() + Environment.NewLine +
                "Содержание: " + drv["Содержание"].ToString() + Environment.NewLine +
                "Ответ на документ: " + drv["НомерВходВходящегоДокумента"].ToString() + Environment.NewLine;
        }

        private void dataGridРабочиеДокументы_Leave(object sender, System.EventArgs e)
        {
            if (labelИнфоTab1.Focused)
            {
                return;
            }
            Статистика();
        }

        private void dataGridДокументыВДеле_Leave(object sender, System.EventArgs e)
        {
            if (labelИнфоTab2.Focused)
            {
                return;
            }
            Статистика();
        }


        private void dataGridИсходящиеДокументы_Leave(object sender, EventArgs e)
        {
            if (labelИнфоTab3.Focused)
            {
                return;
            }
            Статистика();
        }

        private void checkBoxKontrolFilter_CheckedChanged(object sender, System.EventArgs e)
        {
            if (checkBoxKontrolFilter.Checked)
            {
                textBoxСтрокаПоискаTab1.Text = "";
                textBoxСтрокаПоискаTab1.Enabled = false;
                dataViewВыборкаРабДокументы.RowFilter = "ВДело=False AND НаКонтроле=True AND ДатаПоступ >='01.12.2011'";
                dataGridРабочиеДокументы.DataSource = dataViewВыборкаРабДокументы;
            }
            else
            {
                textBoxСтрокаПоискаTab1.Text = "";
                textBoxСтрокаПоискаTab1.Enabled = true;
                dataViewВыборкаРабДокументы.RowFilter = "ВДело=False AND ДатаПоступ >='01.12.2011'";
                dataGridРабочиеДокументы.DataSource = dataViewВыборкаРабДокументы;
            }
            Статистика();

        }

        /// <summary>
        /// Изменение текста в textBoxСтрокаПоиска
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxСтрокаПоиска_TextChanged(object sender, System.EventArgs e)
        {
            dataGridРабочиеДокументы.DataSource = null;
            string фильтр = "ВДело=False AND ДатаПоступ >='01.12.2011' AND (КраткоеСодержание LIKE '%" + textBoxСтрокаПоискаTab1.Text + "%'" +
                            " OR ОписаниеДокумента LIKE '%" + textBoxСтрокаПоискаTab1.Text + "%'" +
                            " OR ОписаниеКорреспондента LIKE '%" + textBoxСтрокаПоискаTab1.Text + "%'" +
                            " OR Резолюция LIKE '%" + textBoxСтрокаПоискаTab1.Text + "%'" +
                            " OR НомерВход LIKE '%" + textBoxСтрокаПоискаTab1.Text + "%'" +
                            " OR НомерИсход LIKE '%" + textBoxСтрокаПоискаTab1.Text + "%')";
            dataViewВыборкаРабДокументы.RowFilter = фильтр;
            dataGridРабочиеДокументы.DataSource = dataViewВыборкаРабДокументы;
            textBoxСтрокаПоискаTab1.Focus();
            if (textBoxСтрокаПоискаTab1.Text != "")
            {
                labelОтобраноДокументовПоискомTab1.Text = "Отобрано документов: " + dataViewВыборкаРабДокументы.Count;
            }
            else
            {
                labelОтобраноДокументовПоискомTab1.Text = "";
            }
        }

        /// <summary>
        /// Изменение текста в строке поиска по документам "в деле"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBoxСтрокаПоискаTab2_TextChanged(object sender, System.EventArgs e)
        {
            dataGridДокументыВДеле.DataSource = null;
            //string фильтр = "ВДело=True AND ДатаПоступ >='01.12.2011' AND (КраткоеСодержание LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
            //    " OR ОписаниеДокумента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
            //    " OR ОписаниеКорреспондента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
            //    " OR РезультатВыполнения LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
            //    " OR Резолюция LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
            //    " OR НомерВход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
            //    " OR НомерИсход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%')";

            //========================

            //Объявим переменную - фильтр
            string фильтр;
            if (this.comboBoxКорреспонденты.Visible == true)
            {
                фильтр = "ВДело=True AND ДатаПоступ >='01.12.2011' AND (КраткоеСодержание LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR ОписаниеДокумента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR ОписаниеКорреспондента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR РезультатВыполнения LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR Резолюция LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR НомерВход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR НомерИсход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%')" +
                    " AND ОписаниеКорреспондента = '" + this.comboBoxКорреспонденты.Text + "'";
                //" AND ОписаниеКорреспондента = 'ГКУ СО \"КСПН г.Саратова\"'";
            }
            else
            {
                фильтр = "ВДело=True AND ДатаПоступ >='01.12.2011' AND (КраткоеСодержание LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR ОписаниеДокумента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR ОписаниеКорреспондента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR РезультатВыполнения LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR Резолюция LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR НомерВход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR НомерИсход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%')";
            }
            
            if(this.textBoxСтрокаПоискаTab2.Text == "")
            {
                фильтр = "ВДело=True AND ДатаПоступ >='01.12.2011' AND (КраткоеСодержание LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR ОписаниеДокумента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR ОписаниеКорреспондента LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR РезультатВыполнения LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR Резолюция LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR НомерВход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%'" +
                    " OR НомерИсход LIKE '%" + textBoxСтрокаПоискаTab2.Text + "%')";
            }
            //=====================
            dataViewВыборкаДокументыВДеле.RowFilter = фильтр;
            dataGridДокументыВДеле.DataSource = dataViewВыборкаДокументыВДеле;
            textBoxСтрокаПоискаTab2.Focus();
            if (textBoxСтрокаПоискаTab2.Text != "")
            {
                labelОтобраноДокументовПоискомTab2.Text = "Отобрано документов: " + dataViewВыборкаДокументыВДеле.Count;
            }
            else
            {
                labelОтобраноДокументовПоискомTab2.Text = "";
            }
        }

        private void textBoxСтрокаПоискаИсходящихДокументов_TextChanged(object sender, EventArgs e)
        {
            DataView view = (DataView)dataGridИсходящиеДокументы.DataSource;
            view.RowFilter = ФильтрИД;
            textBoxСтрокаПоискаИсходящихДокументов.Focus();
            labelОтобраноДокументовПоискомИсходящихДокументов.Text = "Отобрано документов: " + view.Count;
        }


        /// <summary>
        /// Очищает результаты поиска по рабочим документам
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonОчиститьСтрокуПоискаTab1_Click(object sender, System.EventArgs e)
        {
            checkBoxKontrolFilter.Checked = false;
            textBoxСтрокаПоискаTab1.Text = "";
            Статистика();
        }

        /// <summary>
        /// Сбрасывает условия поиска
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonОчиститьСтрокуПоиска_Click(object sender, System.EventArgs e)
        {
            textBoxСтрокаПоискаTab1.Text = "";
        }

        /// <summary>
        /// Очистить поиск по документам "в деле"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonОчиститьСтрокуПоискаTab2_Click(object sender, System.EventArgs e)
        {
            textBoxСтрокаПоискаTab2.Text = "";
        }

        private void buttonОчиститьСтрокуПоискаTab2_Click_1(object sender, System.EventArgs e)
        {
            textBoxСтрокаПоискаTab2.Text = "";
        }

        /// <summary>
        /// Очищает строку поиска Исходящих документов
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonОчиститьСтрокуПоискаИсходящихДокументов_Click(object sender, EventArgs e)
        {
            textBoxСтрокаПоискаИсходящихДокументов.Text = "";
            comboBoxФильтрИДПоДате.SelectedItem = "Весь год";
        }

        private void labelИнфоTab1_Leave(object sender, System.EventArgs e)
        {
            Статистика();
        }

        private void labelИнфоTab2_Leave(object sender, System.EventArgs e)
        {
            Статистика();
        }

        /// <summary>
        /// Контекстное меню ДОБАВИТЬ ЗАПИСЬ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContextДобавитьЗапись_Click(object sender, EventArgs e)
        {

               string iTest = выбранныйГод;

            // Переменная для хранения id основания передачи персональных.
            int idPersonDate = 0;

            int seletYear = Convert.ToInt16(this.выбранныйГод) + 1;

            string filePatchLog = Application.StartupPath + @"\fileLog.txt";

            if (File.Exists(filePatchLog) == true)
            {
                File.Delete(filePatchLog);
                Log.WriteLine(filePatchLog, "Создадим лог");
            }
            else
            {
                Log.WriteLine(filePatchLog, "Создали лог файл");
            }
            

            // переменная для хранения прирощения дат.
            int inc = 0;
            StringBuilder builder = new StringBuilder();

            // Строка для хранения запроса к БД для получения id получателей.
            StringBuilder buildКор = new StringBuilder();

            //FormКарточка form = new FormКарточка(ds11, выбранныйГод,false);
            FormКарточка form = new FormКарточка(ds11, seletYear.ToString(), false);

            // Добавим в карточку текущий год.
            form.CurrentYear = this.selectedYear;

            form.ShowDialog(this);

            // Возможен касяк.-
            НомерДокумента docNumNext = form.СледующийНомерДокумента;

            if (form.DialogResult == DialogResult.OK)
            {

                DS1.КарточкаRow row = form.строкаКарточки;

                // Сгенерируем ГУИД для идентификации документа.
                Guid guidCard = Guid.NewGuid();

                inc = form.IncrementDate;

                string patchToServer = string.Empty;

                // Переменная для хранения имени файла на сервере.
                string namFileServer = string.Empty;

                // Получим выбранный способ поступления документа выбрал пользователь.
                ItemСпособПоступленияДокумента способПоступленияДокумента = form.СпособПоступления;

                // Пролучим список начальников управлений и отделов которым отписан документ.
                this.ListPerson = form.ListPerson;

                // Архивируем файл.

                //Если установлен флаг сохранения ксерокопии документа на сервере.
                if (form.SaveDocServer == true)
                {
                    if (form.ФлагЗаписиАрхива == true)
                    {
                        // Получим путь к файлу.
                        string filePatch = form.PathFileServer;

                        // Имя программы архиватора.
                        //string archiver = @"C:\Program Files\7-Zip\7z.exe";

                        // Получим имя папки которую нужно заархивировать.
                        string archive = form.FileName;// +@"\*.*";

                        // GUID составляющая названия файла.
                        string file = form.PathFileServer;

                        // Создадим своё имя файла архива содержащего архивируемую папку.
                        string namFileS = docNumNext.Номер.ToString() + "-" + docNumNext.Префикс + "_" + file;
                        string namFile = docNumNext.Номер.ToString() + "-" + docNumNext.Префикс;

                        // Путь к временному размещению папки с архивом.
                        string patch = Application.StartupPath + @"\Archive\" + namFile + ".7z";

                        fileName = patch;

                        namFileServer = namFile;// +".7z";

                        // Директоря куда будем архивировать файл.
                        string patchDir = Application.StartupPath + @"\Archive\";

                        // Архивируем папку. (Старая реализация)
                        //Archiver.AddToArchive(archiver, archive, patch,patchDir);

                        Log.WriteLine(filePatchLog, "Архивирование файла начало");

                        // Путь к 7z.dll.
                        string sevenZipDll = Application.StartupPath + @"\7z.dll";
                        if (archive.Length > 0)
                        {
                            // Пометим, что документ помечен для записи на сервер.
                            flagInsertCopyDoc = true;

                            Archiver.AddToArchive(sevenZipDll, archive, patch, patchDir);
                        }
                        else
                        {
                            // Пометим, что документ для записи на сервер не посмечен.
                            flagInsertCopyDoc = false;

                            MessageBox.Show("Вы не указали какую папку с документами записать на сервер","Внимание",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                        }


                        Log.WriteLine(filePatchLog, "Архивирование файла конец");

                        // Путь куда будем архивировать папку.
                        patchToServer = patchServerFile + @"\" + namFileS.Trim();

                        fileNameCopy = patchToServer;
                    }
                    else
                    {
                        return;
                    }
                }

                // ===Begin========Запишем фамилии кому отписано документ в базу данных.
                // Разобъём строку на фамилии. (символ ,)
                string[] sКоррs = row["Резолюция"].ToString().Split(',');

                int id_карточки = Convert.ToInt32(row["id_карточки"]);

                // Полоучим время записи.
                DateTime todoy = DateTime.Now;

                // Счётчик циклов.
                int iCount = 1;

                //// Сформируем строку 
                //foreach (string str in sКоррs)
                //{
                //    string insert = "declare @id_" + iCount + "  int " +
                //                    "SELECT @id_" + iCount + " = id_получателя " +
                //                    "FROM [Получатели] " +
                //                    "where [ОписаниеПолучателя] = '" + str.Trim() + "' " +
                //                    "INSERT INTO [ПолучателДокументовУправление] " +
                //                               "([idПолучатель] " +
                //                               ",[ДатаВремяЗаписи] " +
                //                               ",[ОтметкаПрочтение] " +
                //                               ",[ОтметкаИсполнение] " +
                //                               ",[idКарточки] " +
                //                               ",[РезультатВыполнения]) " +
                //                         "VALUES " +
                //                               "(@id_" + iCount + " " +
                //                               ",'" + todoy + "' " +
                //                               ",NULL " +
                //                               ",NULL " +
                //                               ","+ id_карточки +" " +
                //                               ",NULL) ";

                //    // Добавим в запрос.
                //    builder.Append(insert);

                //    iCount++;
                //}
                //=============End=====================

                // Если пользователь не указал документ который нужно архивировать и отправлять на сервер
                // тогла установим флаг записи документа на запись без привязки номера документа.
                if (flagInsertCopyDoc == false)
                {
                    form.ФлагЗаписиАрхива = false;
                }
               
                // По умолчанию, добавляем новую карточку письмо НЕ ИНИЦИАТИВНОЕ.
                if (form.FlagRecordRepeet == false)
                {
                    string queryInsert = string.Empty;

                    // Строка для хранения хеш длинны файла.
                    string md5 = string.Empty;


                    if (form.ФлагЗаписиАрхива == true)
                    {
                        if (form.FlagAddDoc == true)
                        {

                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          " declare @id_карточки int " +
                                          " declare @номерПП int  " +
                                        " select top 1 @номерПП = номерПП from Карточка " +
                " where ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' and ДатаПоступ >= '" + выбранныйГод + "1231' and FlagAuto is null " +
                  "order by id_карточки desc " +
                                         //" where ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                         // " and id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                                // " where FlagAuto is null) " + 
                                                //"order by номерПП desc " +
                                                 "INSERT INTO Карточка " +
                                                 "([id_документа] " +
                                                 ",[id_корреспондента] " +
                                                 ",[ВДело] " +
                                                ",[ДатаИсхода] " +
                                                ",[ДатаПоступ] " +
                                                ",[КраткоеСодержание] " +
                                                ",[НаКонтроле] " +
                                                ",[НомерВход] " +
                                                ",[НомерИсход] " +
                                                ",[Резолюция] " +
                                                ",[РезультатВыполнения] " +
                                                ",[СрокВыполнения] " +
                                                ",[номерПП] " +
                                                ",[ОписаниеКорреспондента] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                ",MD5 " +
                                                ",idВидПоступленияДокумента  " +
                                                ",DataWriterServerDoc " +
                                                ",NameFileDocumentVipNetEmailTitlePage " +
                                                //",FileDate " +
                                                //",FileDateTitlePage " +
                                                ",FlagAuto " +
                                                ",ДСП ) " +
                                                "VALUES " +
                                                "( " + row["id_документа"] + " " +
                                                "," + row["id_корреспондента"] + " " +
                                                ",'" + row["ВДело"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["КраткоеСодержание"] + "' " +
                                                ",'" + row["НаКонтроле"] + "' " +
                                                ",'" + row["НомерВход"] + "' " +
                                                ",'" + row["НомерИсход"] + "' " +
                                                ",'" + row["Резолюция"] + "' " +
                                                ",'" + row["РезультатВыполнения"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["номерПП"] + " " +
                                                //"," + docNumNext.Номер + " " +
                                                ", @номерПП + 1 " +
                                                ",'" + row["ОписаниеКорреспондента"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",'" + namFileServer + "'  " +
                                                ",'" + form.PathFileServer + "' " +
                                                ",'md5' " +
                                                "," + способПоступленияДокумента.Id + "  " +
                                                ", NULL " +
                                                ", NULL " +
                                                //", NULL " +
                                                //", NULL " +
                                                ", NULL " +
                                                ", '"+ form.FlagDsp +"' ) " +
                                                "SELECT @id_карточки = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                        else
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          " declare @id_карточки int " +
                                         " declare @номерПП int  " +
                                          " select top 1 @номерПП = номерПП from Карточка " +
                                    " where ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' and ДатаПоступ >= '" + выбранныйГод + "1231' and FlagAuto is null " +
                                      "order by id_карточки desc " +
                                          ////" select top 1 @номерПП = номерПП from Карточка " +
                                          ////" where FlagAuto is null and ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' " +
                                          ////"order by id_карточки desc " +
                                         // " select top 1 @номерПП = номерПП from Карточка " +
                                         // " where ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                         //// " where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                         // " and id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                                         // " where FlagAuto is null) " +
                                         //        "order by номерПП desc " +
                                                 "INSERT INTO Карточка " +
                                                 "([id_документа] " +
                                                 ",[id_корреспондента] " +
                                                 ",[ВДело] " +
                                                ",[ДатаИсхода] " +
                                                ",[ДатаПоступ] " +
                                                ",[КраткоеСодержание] " +
                                                ",[НаКонтроле] " +
                                                ",[НомерВход] " +
                                                ",[НомерИсход] " +
                                                ",[Резолюция] " +
                                                ",[РезультатВыполнения] " +
                                                ",[СрокВыполнения] " +
                                                ",[номерПП] " +
                                                ",[ОписаниеКорреспондента] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                 ",MD5 " +
                                                ",idВидПоступленияДокумента  " +
                                                 ",DataWriterServerDoc " +
                                                ",NameFileDocumentVipNetEmailTitlePage " +
                                                //",FileDate " +
                                                //",FileDateTitlePage " +
                                                ",FlagAuto " +
                                                ",ДСП ) " +
                                                "VALUES " +
                                                "( " + row["id_документа"] + " " +
                                                "," + row["id_корреспондента"] + " " +
                                                ",'" + row["ВДело"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["КраткоеСодержание"] + "' " +
                                                ",'" + row["НаКонтроле"] + "' " +
                                                ",'" + row["НомерВход"] + "' " +
                                                ",'" + row["НомерИсход"] + "' " +
                                                ",'" + row["Резолюция"] + "' " +
                                                ",'" + row["РезультатВыполнения"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["номерПП"] + " " +
                                 //"," + docNumNext.Номер + " " +
                                                ", @номерПП + 1 " +
                                                ",'" + row["ОписаниеКорреспондента"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",'" + namFileServer + "'  " +
                                                ",'" + form.PathFileServer + "' " +
                                                ",NULL " + 
                                                 "," + способПоступленияДокумента.Id + "  " +
                                                 ", NULL " +
                                                ", NULL " +
                                                //", NULL " +
                                                //", NULL " +
                                                ", NULL " +
                                                ", '" + form.FlagDsp + "' ) " +
                                                "SELECT @id_карточки = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                    }
                    else
                    {

                        if (form.FlagAddDoc == true)
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                            " declare @id_карточки int " +
                                         " declare @номерПП int  " +
                                          " select top 1 @номерПП = номерПП from Карточка " +
                                        " where ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' and ДатаПоступ >= '" + выбранныйГод + "1231' and FlagAuto is null " +
                                          "order by id_карточки desc " +
                                                                  ////" select top 1 @номерПП = номерПП from Карточка " +
                                          ////" where FlagAuto is null and ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' " +
                                          ////"order by id_карточки desc " +
                                          //" select top 1 @номерПП = номерПП from Карточка " +
                                          ////" where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                          //" where ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                          //" and id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                                          //" where FlagAuto is null) " +
                                          //       "order by номерПП desc " +
                                                "INSERT INTO Карточка " +
                                                 "([id_документа] " +
                                                 ",[id_корреспондента] " +
                                                 ",[ВДело] " +
                                                ",[ДатаИсхода] " +
                                                ",[ДатаПоступ] " +
                                                ",[КраткоеСодержание] " +
                                                ",[НаКонтроле] " +
                                                ",[НомерВход] " +
                                                ",[НомерИсход] " +
                                                ",[Резолюция] " +
                                                ",[РезультатВыполнения] " +
                                                ",[СрокВыполнения] " +
                                                ",[номерПП] " +
                                                ",[ОписаниеКорреспондента] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                ",MD5 " +
                                                ",idВидПоступленияДокумента  " +
                                                  ",DataWriterServerDoc " +
                                                ",NameFileDocumentVipNetEmailTitlePage " +
                                                //",FileDate " +
                                                //",FileDateTitlePage " +
                                                ",FlagAuto " +
                                                ",ДСП ) " +
                                                "VALUES " +
                                                "( " + row["id_документа"] + " " +
                                                "," + row["id_корреспондента"] + " " +
                                                ",'" + row["ВДело"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["КраткоеСодержание"] + "' " +
                                                ",'" + row["НаКонтроле"] + "' " +
                                                ",'" + row["НомерВход"] + "' " +
                                                ",'" + row["НомерИсход"] + "' " +
                                                ",'" + row["Резолюция"] + "' " +
                                                ",'" + row["РезультатВыполнения"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["номерПП"] + " " +
                                 //"," + docNumNext.Номер + " " +
                                                ", @номерПП + 1 " +
                                                ",'" + row["ОписаниеКорреспондента"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",NULL  " +
                                                ",'" + guidCard + "' " +
                                                ",'md5' " +
                                                "," + способПоступленияДокумента.Id + "  " +
                                                ", NULL " +
                                                ", NULL " +
                                                //", NULL " +
                                                //", NULL " +
                                                ", NULL " +
                                                ", '" + form.FlagDsp + "' ) " +
                                                "SELECT @id_карточки = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                        else
                        {

                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                            " declare @id_карточки int " +
                                         " declare @номерПП int  " +
                                         " select top 1 @номерПП = номерПП from Карточка " +
                                        " where ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' and ДатаПоступ >= '" + выбранныйГод + "1231' and FlagAuto is null " +
                                          "order by id_карточки desc " +
                                          ////" select top 1 @номерПП = номерПП from Карточка " +
                                          ////" where FlagAuto is null and ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' " +
                                          ////"order by id_карточки desc " +
                                          //" select top 1 @номерПП = номерПП from Карточка " +
                                          ////" where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                          //" where ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                          //" and id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                                          //" where FlagAuto is null) " +
                                          //       "order by номерПП desc " +
                                                "INSERT INTO Карточка " +
                                                 "([id_документа] " +
                                                 ",[id_корреспондента] " +
                                                 ",[ВДело] " +
                                                ",[ДатаИсхода] " +
                                                ",[ДатаПоступ] " +
                                                ",[КраткоеСодержание] " +
                                                ",[НаКонтроле] " +
                                                ",[НомерВход] " +
                                                ",[НомерИсход] " +
                                                ",[Резолюция] " +
                                                ",[РезультатВыполнения] " +
                                                ",[СрокВыполнения] " +
                                                ",[номерПП] " +
                                                ",[ОписаниеКорреспондента] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                               ",MD5 " +
                                                ",idВидПоступленияДокумента  " +
                                                ",DataWriterServerDoc " +
                                                ",NameFileDocumentVipNetEmailTitlePage " +
                                                //",FileDate " +
                                                //",FileDateTitlePage " +
                                                ",FlagAuto " +
                                                ",ДСП ) " +
                                                "VALUES " +
                                                "( " + row["id_документа"] + " " +
                                                "," + row["id_корреспондента"] + " " +
                                                ",'" + row["ВДело"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["КраткоеСодержание"] + "' " +
                                                ",'" + row["НаКонтроле"] + "' " +
                                                ",'" + row["НомерВход"] + "' " +
                                                ",'" + row["НомерИсход"] + "' " +
                                                ",'" + row["Резолюция"] + "' " +
                                                ",'" + row["РезультатВыполнения"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["номерПП"] + " " +
                                                //"," + docNumNext.Номер + " " +
                                                ", @номерПП + 1 " +
                                                ",'" + row["ОписаниеКорреспондента"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",NULL  " +
                                                ",'" + guidCard + "' " +
                                                ",NULL " +
                                                  "," + способПоступленияДокумента.Id + "  " +
                                                   ", NULL " +
                                                ", NULL " +
                                                //", NULL " +
                                                //", NULL " +
                                                ", NULL " +
                                                ", '" + form.FlagDsp + "' ) " +
                                                "SELECT @id_карточки = @@IDENTITY  ";

                                               

                            builder.Append(queryInsert);
                        }
                    }

                    // Сформируем строку связывающую номер каротчки с пользователем кому отписан документ.
                    foreach (string str in sКоррs)
                    {
                        string insert = "declare @id_" + iCount + "  int " +
                                        "SELECT @id_" + iCount + " = id_получателя " +
                                        "FROM [Получатели] " +
                                        "where [ОписаниеПолучателя] = '" + str.Trim() + "' " +
                                        "INSERT INTO [ПолучателДокументовУправление] " +
                                                   "([idПолучатель] " +
                                                   ",[ДатаВремяЗаписи] " +
                                                   ",[ОтметкаПрочтение] " +
                                                   ",[ОтметкаИсполнение] " +
                                                   ",[idКарточки] " +
                                                   ",[РезультатВыполнения]) " +
                                             "VALUES " +
                                                   "(@id_" + iCount + " " +
                                                   //",'" + ДатаSQL.Дата(todoy.ToShortDateString()) + "' " +
                                                   ",GETDATE() " +
                                                   ",NULL " +
                                                   ",NULL " +
                                                   ",@id_карточки " +
                                                   ",NULL) ";

                        // Добавим в запрос.
                        builder.Append(insert);

                        iCount++;
                    }

                    // Сформируем запись в связующую таблицу документа, вида получения документа и начальниками отделов и управлений которым отписан текущий документ.
                    foreach (PersonRecepient person in this.ListPerson)
                    {
                        string insert = "INSERT INTO [СвязующаяВидПоступленияДокПолучатели] " +
                                        "([id_person] " +
                                       ",[id_ВидПоступленияДок] " +
                                       ",[id_карточки]) " +
                                       "VALUES " +
                                       "(" + person.ID + " " +
                                       "," + способПоступленияДокумента.Id + " " +
                                       ",@id_карточки ) ";

                        // Добавим в запрос.
                        builder.Append(insert);
                    }

                    // Добавим основания для передачи персональных данных.
                    builder.Append(form.QueryPersonDateForCardInput);
                                                
                    //builder.Append(queryInsert + "COMMIT TRANSACTION ");
                    builder.Append("COMMIT TRANSACTION ");

                    string sTest = builder.ToString().Trim();
                }

                // Добавляем новое письмо ИНИЦИАТИВНОЕ.
                if (form.FlagRecordRepeet == true)
                {
                    string queryInsert = string.Empty;
                    if (form.ФлагЗаписиАрхива == true)
                    {
                        queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                      "begin transaction  " +
                                        " declare @id_карточки int " +
                                      " declare @номерПП int " +
                                       " select top 1 @номерПП = номерПП from Карточка " +
                " where ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' and ДатаПоступ >= '" + выбранныйГод + "1231' and FlagAuto is null " +
                  "order by id_карточки desc " +
                            ////" select top 1 @номерПП = номерПП from Карточка " +
                            ////   " where FlagAuto is null and ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' " +
                            ////   "order by id_карточки desc " +
                            //               "select top 1 @номерПП = номерПП from Карточка " +
                            //              //"where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                            //              " where ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                            //              " and id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                            //              " where FlagAuto is null) " +
                            ////"select top 1 @номерПП = [номерПП] from Карточка " +
                            ////"where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' " +  and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                            //                     "order by номерПП desc " +
                                            " INSERT INTO Карточка " +
                                             "([id_документа] " +
                                             ",[id_корреспондента] " +
                                             ",[ВДело] " +
                                            ",[ДатаИсхода] " +
                                            ",[ДатаПоступ] " +
                                            ",[КраткоеСодержание] " +
                                            ",[НаКонтроле] " +
                                            ",[НомерВход] " +
                                            ",[НомерИсход] " +
                                            ",[Резолюция] " +
                                            ",[РезультатВыполнения] " +
                                            ",[СрокВыполнения] " +
                                            ",[номерПП] " +
                                            ",[ОписаниеКорреспондента] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                             ",[FlagCardRepeet] " +
                                            ",NameFileDocument ) " +
                                            "VALUES " +
                                            "( " + row["id_документа"] + " " +
                                            "," + row["id_корреспондента"] + " " +
                                            ",'" + row["ВДело"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["КраткоеСодержание"] + "' " +
                                            ",'" + row["НаКонтроле"] + "' " +
                                            ",'" + row["НомерВход"] + "' " +
                                            ",'" + row["НомерИсход"] + "' " +
                                            ",'" + row["Резолюция"] + "' " +
                                            ",'" + row["РезультатВыполнения"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["номерПП"] + " " +
                            //"," + docNumNext.Номер + " " +
                                            ", @номерПП + 1 " +
                                            ",'" + row["ОписаниеКорреспондента"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                             ",'" + namFileServer + "'  " +
                                            ",'" + form.PathFileServer + "' ) " +
                                           "INSERT INTO КарточкаПовтор " +
                                             "([id_документа] " +
                                             ",[id_корреспондента] " +
                                             ",[ВДело] " +
                                            ",[ДатаИсхода] " +
                                            ",[ДатаПоступ] " +
                                            ",[КраткоеСодержание] " +
                                            ",[НаКонтроле] " +
                                            ",[НомерВход] " +
                                            ",[НомерИсход] " +
                                            ",[Резолюция] " +
                                            ",[РезультатВыполнения] " +
                                            ",[СрокВыполнения] " +
                                            ",[номерПП] " +
                                            ",[ОписаниеКорреспондента] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                            ",id_карточкиВходящей  " +
                                            ",ДатаПрирощение " +
                                            ",FlagControl)" +
                                            "VALUES " +
                                            "( " + row["id_документа"] + " " +
                                            "," + row["id_корреспондента"] + " " +
                                            ",'" + row["ВДело"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["КраткоеСодержание"] + "' " +
                                            ",'" + row["НаКонтроле"] + "' " +
                                            ",'" + row["НомерВход"] + "' " +
                                            ",'" + row["НомерИсход"] + "' " +
                                            ",'" + row["Резолюция"] + "' " +
                                            ",'" + row["РезультатВыполнения"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["номерПП"] + " " +
                                            "," + docNumNext.Номер + " " +
                                            ",'" + row["ОписаниеКорреспондента"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                            ",@@IDENTITY " +
                                            "," + inc + " " +
                                            ",'False') ";                     }
                    else
                    {
                        queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                      "begin transaction  " +
                                        " declare @id_карточки int " +
                                      "declare @номерПП int " +
                                       " select top 1 @номерПП = номерПП from Карточка " +
                " where ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' and ДатаПоступ >= '" + выбранныйГод + "1231' and FlagAuto is null " +
                  "order by id_карточки desc " +
                            ////" select top 1 @номерПП = номерПП from Карточка " +
                            ////   " where FlagAuto is null and ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' " +
                            ////   "order by id_карточки desc " +
                            //              "select top 1 @номерПП = номерПП from Карточка " +
                            //              //"where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                            //              " where ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                            //              " and id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                            //              " where FlagAuto is null) " +
                            ////"select top 1 @номерПП = [номерПП] from Карточка " +
                            ////"where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' " +  and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                            //                     "order by номерПП desc " +
                                      "INSERT INTO Карточка " +
                                             "([id_документа] " +
                                             ",[id_корреспондента] " +
                                             ",[ВДело] " +
                                            ",[ДатаИсхода] " +
                                            ",[ДатаПоступ] " +
                                            ",[КраткоеСодержание] " +
                                            ",[НаКонтроле] " +
                                            ",[НомерВход] " +
                                            ",[НомерИсход] " +
                                            ",[Резолюция] " +
                                            ",[РезультатВыполнения] " +
                                            ",[СрокВыполнения] " +
                                            ",[номерПП] " +
                                            ",[ОписаниеКорреспондента] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                             ",[FlagCardRepeet] " +
                                            ",NameFileDocument ) " +
                                            "VALUES " +
                                            "( " + row["id_документа"] + " " +
                                            "," + row["id_корреспондента"] + " " +
                                            ",'" + row["ВДело"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["КраткоеСодержание"] + "' " +
                                            ",'" + row["НаКонтроле"] + "' " +
                                            ",'" + row["НомерВход"] + "' " +
                                            ",'" + row["НомерИсход"] + "' " +
                                            ",'" + row["Резолюция"] + "' " +
                                            ",'" + row["РезультатВыполнения"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["номерПП"] + " " +
                            //"," + docNumNext.Номер + " " +
                                            ", @номерПП + 1 " +
                                            ",'" + row["ОписаниеКорреспондента"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                             ",NULL  " +
                                             ",NULL ) " +
                                           "INSERT INTO КарточкаПовтор " +
                                             "([id_документа] " +
                                             ",[id_корреспондента] " +
                                             ",[ВДело] " +
                                            ",[ДатаИсхода] " +
                                            ",[ДатаПоступ] " +
                                            ",[КраткоеСодержание] " +
                                            ",[НаКонтроле] " +
                                            ",[НомерВход] " +
                                            ",[НомерИсход] " +
                                            ",[Резолюция] " +
                                            ",[РезультатВыполнения] " +
                                            ",[СрокВыполнения] " +
                                            ",[номерПП] " +
                                            ",[ОписаниеКорреспондента] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                            ",id_карточкиВходящей  " +
                                            ",ДатаПрирощение " +
                                            ",FlagControl)" +
                                            "VALUES " +
                                            "( " + row["id_документа"] + " " +
                                            "," + row["id_корреспондента"] + " " +
                                            ",'" + row["ВДело"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["КраткоеСодержание"] + "' " +
                                            ",'" + row["НаКонтроле"] + "' " +
                                            ",'" + row["НомерВход"] + "' " +
                                            ",'" + row["НомерИсход"] + "' " +
                                            ",'" + row["Резолюция"] + "' " +
                                            ",'" + row["РезультатВыполнения"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["номерПП"] + " " +
                                            "," + docNumNext.Номер + " " +
                                            ",'" + row["ОписаниеКорреспондента"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                            ",@@IDENTITY " +
                                            "," + inc + " " +
                                            ",'False') ";
                                           
                    }

                    // Добавим строку запроса на Insert.
                    builder.Append(queryInsert);

                    // Добавим основания для передачи персональных данных.
                    builder.Append(form.QueryPersonDateForCardInput);

                    // Закроем транзакцию.
                    builder.Append("COMMIT TRANSACTION");

                    string iTest2 = "";
                }

                //// Сохраним данные.
                ПодключитьБД connectBD = new ПодключитьБД();
                string sCon = connectBD.СтрокаПодключения();

                // Флаг проверки успешной копии файла.
                bool flagCopyServer = false;

                string sTest2 = builder.ToString().Trim();

                // Выполним запрос на вставку (к сожалению не в единой транзакции.
                using (SqlConnection con = new SqlConnection(sCon))
                {

                    //Log.WriteLine(filePatchLog, "Начнём копровать файл на сервер");

                    //if (form.SaveDocServer == true)
                    //{
                    //    try
                    //    {
                    //        if (form.ФлагЗаписиАрхива == true)
                    //        {

                    //            Log.WriteLine(filePatchLog, "Копируем файл на сервер");

                    //            // Проверим помечен ли документ для записи на сервер.
                    //            if (flagInsertCopyDoc == true)
                    //            {
                    //                //Скопируем файл на сервер хранения документов.
                    //                //File.Copy(fileName, fileNameCopy, true);
                    //            }
                    //            Log.WriteLine(filePatchLog, "Закончим копировать файл на сервер");

                    //            //Если файл скопировался успешно постави флаг в true.
                    //            flagCopyServer = true;
                    //        }
                    //    }
                    //    catch(Exception exp)
                    //    {
                    //        Log.WriteLine(filePatchLog, "Ошибка при копировании - ");
                    //        Log.WriteLine(filePatchLog, exp.Message);
                    //        MessageBox.Show("Ошибка при копировании файла");

                    //        flagCopyServer = false;

                    //        return;
                    //    }

                    //    string fileTest = fileNameCopy;
                    //    if (File.Exists(fileNameCopy) == true)
                    //    {
                    //        Log.WriteLine(filePatchLog, "Выполним запись на сервер");

                    //        con.Open();
                    //        SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    //        com.ExecuteNonQuery();
                    //    }
                    //    else
                    //    {
                    //        con.Open();
                    //        SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    //        com.ExecuteNonQuery();
                    //    }
                    //}
                    //else
                    //{
                            // Если файл скопировался успешно постави флаг в true.
                            flagCopyServer = true;

                            con.Open();
                            SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                            com.ExecuteNonQuery();
                    //}
                }


                //ds11.Карточка.AddКарточкаRow(row);
                ОбновитьДанные();

                // Пойдём по тупому варианту и получим номер документа.
                string queryNumDoc = "select id_карточки,номерПП,НомерВход from [Карточка] " +
                                     "where GuidName = '" + guidCard + "' ";

                string номерДок = string.Empty;

                DataTable tabNum;

                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();

                    SqlDataAdapter da = new SqlDataAdapter(queryNumDoc, con);

                    DataSet ds = new DataSet();

                    da.Fill(ds, "numDoc");

                    tabNum = ds.Tables["numDoc"];
                }

                номерДок = tabNum.Rows[0]["номерПП"].ToString().Trim() + "/" + tabNum.Rows[0]["НомерВход"].ToString().Trim();

                string номер = номерДок;

                // Получим номер id карточки.
                string idCard = tabNum.Rows[0]["id_карточки"].ToString().Trim();
                
                // Выводит номер зарегистрированного документа.
                FormMessage frmMessage = new FormMessage(номер);
                frmMessage.NumCardDoc = idCard.Trim();
                frmMessage.НомерДокумента = номерДок;
                frmMessage.СпособПоступленияДокумента = способПоступленияДокумента;
                frmMessage.TopMost = true;
                frmMessage.ShowDialog();

            }
        }

        /// <summary>
        /// Контекстное меню ДОБАВИТЬ ИСХОДЯЩУЮ ЗАПИСЬ
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContextДобавитьИсходящуюЗапись_Click(object sender, EventArgs e)
        {
            string iTest = выбранныйГод;

            int seletYear = Convert.ToInt16(this.выбранныйГод) + 1;

            
            FormКарточкаИсходящая form = new FormКарточкаИсходящая(ds11, выбранныйГод, false);

            // Установим флаг в false.
            form.FlagОтветПисьмо = false;

            // Установим адресат.
            form.Адресат = "";
            
            DialogResult result = form.ShowDialog(this);
            if (result == DialogResult.OK)
            {

                // Получим выбранный способ поступления документа выбрал пользователь.
                ItemСпособПоступленияДокумента способПоступленияДокумента = form.СпособПоступления;


                DS1.КарточкаИсходящаяRow row = form.строкаИсходящейКарточки;

                НомерДокумента doc = new НомерДокумента();
                doc.Номер = Convert.ToInt16(form.НомерDoc.Номер);

                // Обнулим переменную.
                numberPrefix = string.Empty;

                // Запишем префикс номера документа.
                numberPrefix = form.ПрефиксНомерИсходящий;

                //ds11.КарточкаИсходящая.AddКарточкаИсходящаяRow(row); 

                List<int> listIdВходДок = form.ListIDКарточки;
                //List<ОснованиеПередачи> listOP = form.ListОснованиеПередачи;

                // Если мы пишем инициативное письмо.
                if (form.FlagОтветПисьмо == false)
                {

                  // Строка для хранения SQL инструкции, для выполнения в одной транзакции.
                    StringBuilder buildInsert = new StringBuilder();

                    // Переменная для хранения номера
                    string numDirect = string.Empty;
                    
                    // Проверим работаем с ДСП или нет.
                    string query = string.Empty;

                    if (Convert.ToBoolean(form.FlagDsp) == false)
                    {
                        query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                       "begin transaction  " +
                                       "declare @numDoc int " +
                                        "select top 1 @numDoc = НомерПорядковый from КарточкаИсходящая " +
                                       " where Дата >= '" + выбранныйГод + "1201' and Дата <= '" + (Convert.ToInt32(выбранныйГод) + 1).ToString().Trim() + "1231' " +
                            "order by id_карточки desc " +
                            ////           "select top 1 @numDoc = НомерПорядковый from КарточкаИсходящая " +
                            //////"where Дата >= '" + seletYear.ToString().Trim() + "0101' and Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                            ////"where Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                            ////" and " +
                            ////" id_карточки in (SELECT MAX(id_карточки) FROM [КарточкаИсходящая] " +
                            ////" where FlagAutho is null) " +
                            ////" order by id_карточки desc " +
                                       "declare @key int " +
                                       "INSERT INTO КарточкаИсходящая " +
                                       "([Дата] " +
                                       ",[НомерКомитета] " +
                                       ",[id_Подразделения] " +
                                       ",[НомерНоменклатурный] " +
                                       ",[НомерПорядковый] " +
                                       ",[id_Адресата] " +
                                       ",[Содержание] " +
                                       ",[id_ВходящегоДокумента] " +
                                       ",[ОписаниеКорреспондента] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ",idВидПоступленияДокумента " +
                                       ",FlagAutho " +
                                       ",ДСП ) " +
                                       "VALUES " +
                                       "('" + ДатаSQL.Дата(Convert.ToDateTime(row["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row["НомерКомитета"] + "' " +
                                       "," + row["id_Подразделения"] + " " +
                                       ",'" + row["НомерНоменклатурный"] + "' " +
                            //"," + row["НомерПорядковый"] + " " +
                            //", "+ doc.Номер + " " +
                                       ", @numDoc + 1  " +
                                       "," + row["id_Адресата"] + " " +
                                       ",'" + row["Содержание"] + "' " +
                            //","+ row["id_ВходящегоДокумента"]+" " +
                                       ",NULL " +
                            //",'"+ form.Адресат.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       // ",NULL " +
                                       //",NULL " +
                                       ", " + способПоступленияДокумента.Id + " " +
                                       ", NULL " +
                                       ",'" + form.FlagDsp + "' ) " +
                                       "set @key = @@IDENTITY ";
                    }
                    else
                    {
                        query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                       "begin transaction  " +
                                       "declare @numDoc int " +
                                        "select top 1 @numDoc = НомерПорядковый from КарточкаИсходящая " +
                                       " where Дата >= '" + выбранныйГод + "1201' and Дата <= '" + (Convert.ToInt32(выбранныйГод) + 1).ToString().Trim() + "1231' " +
                            "order by id_карточки desc " + 
                            //  "select top 1 @numDoc = НомерПорядковый from КарточкаИсходящая " +
                            ////"where Дата >= '" + seletYear.ToString().Trim() + "0101' and Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                            //"where Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                            //" and " +
                            //" id_карточки in (SELECT MAX(id_карточки) FROM [КарточкаИсходящая] " +
                            //" where FlagAutho is null) " +
                            //" order by id_карточки desc " +

                                       "declare @key int " +
                                       "INSERT INTO КарточкаИсходящая " +
                                       "([Дата] " +
                                       ",[НомерКомитета] " +
                                       ",[id_Подразделения] " +
                                       ",[НомерНоменклатурный] " +
                                       ",[НомерПорядковый] " +
                                       ",[id_Адресата] " +
                                       ",[Содержание] " +
                                       ",[id_ВходящегоДокумента] " +
                                       ",[ОписаниеКорреспондента] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ",idВидПоступленияДокумента " +
                                       ",FlagAutho " +
                                       ",ДСП  " +
                                       ",ДспDesc ) " +
                                       "VALUES " +
                                       "('" + ДатаSQL.Дата(Convert.ToDateTime(row["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row["НомерКомитета"] + "' " +
                                       "," + row["id_Подразделения"] + " " +
                                       ",'" + row["НомерНоменклатурный"] + "' " +
                            //"," + row["НомерПорядковый"] + " " +
                            //", "+ doc.Номер + " " +
                                       ", @numDoc + 1  " +
                                       "," + row["id_Адресата"] + " " +
                                       ",'" + row["Содержание"] + "' " +
                            //","+ row["id_ВходящегоДокумента"]+" " +
                                       ",NULL " +
                            //",'"+ form.Адресат.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       // ",NULL " +
                                       //",NULL " +
                                       ", " + способПоступленияДокумента.Id + " " +
                                       ", NULL " +
                                       ",'" + form.FlagDsp + "'  " +
                                       ",'ДСП' )" +
                                       "set @key = @@IDENTITY ";
                    }

                    buildInsert.Append(query);

                    // Передадим в строку запроса на вставку SQL инструкцию на вставку в таблицу [СвязующаяЦельПолучениперсональныхДанных].

                    string sInsert = string.Empty;
                    sInsert = String.Format(form.QueryInsert.Trim(), "@key");


                    buildInsert.Append(sInsert.Trim());

                    // Завершим транзакцию.
                    buildInsert.Append("COMMIT TRANSACTION ");

                    string sTestInsertCardInput = buildInsert.ToString();

                    string sTest = "";

                    // Обнулим список для хранения оснований передачи перед использованием.
                    //form.ListОснованиеПередачи.Clear();

                    ПодключитьБД connBD = new ПодключитьБД();
                    string sCon = connBD.СтрокаПодключения();

                    SqlConnection con = new SqlConnection(sCon);
                    con.Open();
                    SqlCommand com = new SqlCommand(buildInsert.ToString(), con);
                    //com.ExecuteNonQuery();
                    con.Close();
                }
                else
                {

                   
                    // Тестируем.
                    DS1.КарточкаИсходящаяRow row2 = form.строкаИсходящейКарточки;

                    int i = row2.id_ВходящегоДокумента;

                    // Строка для хранения запроса.
                    System.Text.StringBuilder builder = new System.Text.StringBuilder();

                    // Установим флаг в FALSE.
                    string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                    "begin transaction  " +
                                    " declare @numCard  int " +
                                    //" select @numCard = MAX(НомерПорядковый) from КарточкаИсходящая " +
                                    ////" where Дата >= '20170101' and FlagAutho is null " +
                                    //"where Дата >= '" + seletYear.ToString().Trim() + "0101' and FlagAutho is null " +
                         "select top 1 @numCard = НомерПорядковый  from КарточкаИсходящая " +
                        //"where Дата >= '" + seletYear.ToString().Trim() + "0101' and Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                         "where Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                          " and " +
                        " id_карточки in (SELECT MAX(id_карточки) FROM [КарточкаИсходящая] " +
                        " where FlagAutho is null) " +
                         "order by id_карточки desc " +
                                        //////" select top 1 @номерПП = номерПП from Карточка " +
                                        //////  " where FlagAuto is null and ДатаПоступ <= '" + seletYear.ToString().Trim() + "1231' " +
                                        //////  "order by id_карточки desc " +
                                           "declare @key int " +
                                   "INSERT INTO КарточкаИсходящая " +
                                   "([Дата] " +
                                   ",[НомерКомитета] " +
                                   ",[id_Подразделения] " +
                                   ",[НомерНоменклатурный] " +
                                   ",[НомерПорядковый] " +
                                   ",[id_Адресата] " +
                                   ",[Содержание] " +
                                   ",[id_ВходящегоДокумента] " +
                                   ",[ОписаниеКорреспондента] " +
                                   ",[FlagPersonData] " +
                                   ",[GUID] " +
                                   //",FileData " +
                                   //",FileDateTitlePage " +
                                   ", idВидПоступленияДокумента)" +
                                   "VALUES " +
                                   "('" + ДатаSQL.Дата(Convert.ToDateTime(row2["Дата"]).ToShortDateString().Trim()) + "' " +
                                   ",'" + row2["НомерКомитета"] + "' " +
                                   "," + row2["id_Подразделения"] + " " +
                                   ",'" + row2["НомерНоменклатурный"] + "' " +
                        //"," + row2["НомерПорядковый"] + " " +
                                    //", " + doc.Номер + " " +
                                    ", @numCard + 1  " +
                                   "," + row2["id_Адресата"] + " " +
                                   ",'" + row2["Содержание"] + "' " +
                                   "," + row2["id_ВходящегоДокумента"] + " " +
                        //",NULL " +
                        //",'"+ form.Адресат.Trim() +"' " +
                                   ",NULL" +
                                   ",'" + row2["FlagPersonData"] + "' " +
                                   ",'" + form.StrGuid.Trim() + "'  " +
                                   //",NULL " +
                                   //",NULL " +
                                   ", " + способПоступленияДокумента.Id + " )" +
                                   " declare @idCard int " +
                                   "select top 1 @idCard = id_карточки  from КарточкаИсходящая " +
                                   "order by id_карточки desc ";
                    

                    builder.Append(query);

                    string номерПодразделения = string.Empty;

                    DataRow[] rowsSelect = ds11.ПодразделенияКомитета.Select("id_подразделения= "+ Convert.ToInt32(row2["id_Подразделения"]) +" ");
                    foreach (DataRow item in rowsSelect)
                    {
                        номерПодразделения = item["НомерПодразделения"].ToString().Trim();
                    }

                    string результатВыполнения = "Дан ответ. № исх. документа " + row2["НомерКомитета"].ToString().Trim() + "-" + row2["НомерНоменклатурный"].ToString().Trim() + "-" + номерПодразделения + "/" + doc.Номер.ToString().Trim();// row2["НомерПорядковый"].ToString().Trim();
                    //string результатВыполнения = "Дан ответ. № исх. документа " + row2["НомерКомитета"].ToString().Trim() + "-" + row2["НомерНоменклатурный"].ToString().Trim() + "-" + номерПодразделения + "/CAST(@numCard + 1 AS nvarchar) ";// +row2["НомерПорядковый"].ToString().Trim();

                    // Установим флаг в TRUE.
                    string queryUpdate = "UPDATE [Карточка] " +
                                         "SET РезультатВыполнения = '" + результатВыполнения + "' " + //' + CAST(@numCard + 1 AS nvarchar) " +
                                         //"FlagPersonData = '" + row["FlagPersonData"] + "' " +
                                         ",ВДело = 'True' " +
                                         "where id_карточки = " + row["id_ВходящегоДокумента"] + " ";
                    // Соберём строки запросаов на добавление записи и на редактирование в единую строку, чтобы выполнить всё в одной транзакции.
                    builder.Append(queryUpdate);

                    string sTestNum = builder.ToString().Trim();

                    // Запрос на всатвку id в связующую таблицу ЦельПолученияПерсДанных.
                    foreach (ОснованиеПередачи itm in form.ListОснованиеПередачи)
                    {
                        string queryIns = "INSERT INTO [СвязующаяЦельПолучениперсональныхДанных] " +
                                       "([id_карточки] " +
                                       ",[id_ОснованиеПередачи]) " +
                                       "VALUES " +
                                       //"('" + row.id_карточки + "' " +
                                       "( @idCard " +
                                       ",'" + itm.Id_основаниеПередачи + "' ) ";

                        builder.Append(queryIns);
                    }


                    // Заполним связующую таблицу Карточка входящаяИсходящая.
                    foreach (int idВх in listIdВходДок)
                    {

                        string queryIdВх = "INSERT INTO [СвязующаяКарточкаВходящаяИсходящая] " +
                                           "([id_карточкаВходящая] " +
                                           ",[id_карточкаИсходящая]) " +
                                           "VALUES " +
                                           "(" + idВх + " " +
                                           //"," + row.id_карточки + " ) " +
                                            ",@idCard ) " + 
                                           "update Карточка " +
                                           "set РезультатВыполнения = '" + form.НомерИсходящий.Trim() + "' " + " + CAST(@numCard + 1 AS nvarchar) " +
                                           "where id_карточки = "+ idВх +" ";
                        
                        builder.Append(queryIdВх);
                    }

                    // Проверим, что документ на который мы отвечаем стоит в стстусе повторных ответов.
                    СтатусКарточка card = new СтатусКарточка(Convert.ToInt32(row["id_ВходящегоДокумента"]));
                    bool flagStatusRepeet = card.СтатусПовторяющийсяОтвет();

                    // Если статус = true значит мы имеем дело с документом на который периодически необходимо довать ответ.
                    if (flagStatusRepeet == true)
                    {
                        // Строка для выполнения запроса в одной транзакции.
                        //StringBuilder querTransact = new StringBuilder();

                        /*
                         * Проверим отвечаем мы на этот документ впервые или нет.
                         * Для этого узнаем значение в поле ВДело в таблице Карточка, если установлено значение False 
                         * тогда на документ мы отвечаем впервые в противном случае нет.
                        */
                        ПодключитьБД bdConnect = new ПодключитьБД();
                        using (SqlConnection conn = new SqlConnection(bdConnect.СтрокаПодключения().Trim()))
                        {
                            conn.Open();
                            СтатусКарточка card2 = new СтатусКарточка(Convert.ToInt32(row["id_ВходящегоДокумента"]));
                            bool flagVD = card2.GetОтветПовторный(conn);

                            // Если на входящую карточку отвечают впервые.
                            if (flagVD == false)
                            {
                                // Установим значение поля ВДело таблицы Карточки в True, а так же поставим флаг в КарточкеПовтор указывающий, что уже один раз на данное письмо был ответ.
                                //string queryUp = " update Карточка " +
                                //               "set ВДело = 'True' " +
                                //               "where id_карточки = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " " + 
                                //               "update КарточкаПовтор " +
                                //               "set FlagControl = 'True' " + 
                                //               "where id_карточкиВходящей = "+ Convert.ToInt32(row["id_ВходящегоДокумента"]) +" ";


                                string queryUp = " update Карточка " +
                                              "set ВДело = 'True' " +
                                              "where id_карточки = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " " +
                                              " declare @date datetime " +
                                              "declare @day int " +
                                              "declare @SetDate datetime " +
                                              "select @date = СрокВыполнения,@day = ДатаПрирощение from КарточкаПовтор " +
                                              "where id_карточкиВходящей = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " " +
                                              "SELECT @SetDate = DATEADD(day, @day, @date); " +
                                              "update КарточкаПовтор " +
                                              "set FlagControl = 'True' " +
                                              ",СрокВыполнения = @SetDate " +
                                              "where id_карточкиВходящей = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " ";

                                //string queryUp = " update Карточка " +
                                //              "set ВДело = 'True' " +
                                //              "where id_карточки = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " ";
                             
                                
                                //querTransact.Append(query);
                                builder.Append(queryUp);
                            }

                            // Если ответ повторный.
                            if (flagVD == true)
                            {
                                // Увеличим значение в поле СрокИсполнения в таблице КарточкаПовтор на количесвто дней указанных в поле ДатаПрирощения.
                                string queryUpdatDate = " declare @date datetime " +
                                                        "declare @day int " +
                                                        "declare @SetDate datetime " +
                                                        "select @date = СрокВыполнения,@day = ДатаПрирощение from КарточкаПовтор " +
                                                        "where id_карточкиВходящей = "+ Convert.ToInt32(row["id_ВходящегоДокумента"]) +" " +
                                                        "SELECT @SetDate = DATEADD(day, @day, @date); " +
                                                        "update КарточкаПовтор " +
                                                        "set СрокВыполнения = @SetDate " +
                                                        "where id_карточкиВходящей = "+ Convert.ToInt32(row["id_ВходящегоДокумента"]) +" ";

                                //querTransact.Append(queryUpdatDate);
                                builder.Append(queryUpdatDate);
                            }
                            

                        }

                    }

                    // Завершим транзакцию.
                    builder.Append("COMMIT TRANSACTION ");

           

                    string queryTest = builder.ToString().Trim();

                    // Выполним запрос.
                    ПодключитьБД strConnectBD = new ПодключитьБД();
                    string strConn = strConnectBD.СтрокаПодключения();

                    // Откроем соединение и выполним запрос.
                    SqlConnection con = new SqlConnection(strConn);
                    con.Open();
                    SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    com.ExecuteNonQuery();

                    // Закроме соединение.
                    con.Close();

                }

                ОбновитьДанные();

                string iTest2 = "test";

                // Получим номер документа.
                NumOutputCardVipNet numDoc = GetNumDocOutVipNet(form.StrGuid);

                //string numberDocument = form.ПрефиксНомерИсходящий.Trim() + "/" + numDoc.Trim();

                string numberDocument = numberPrefix + "/" + numDoc.НомерПорядковый.Trim();

                // Выведим сообщение с новым номером.
                FormMessage message = new FormMessage(numberDocument.Trim());
                message.TopMost = true;
                message.СпособПоступленияДокумента = способПоступленияДокумента;
                message.NumCardDoc = numDoc.Id.ToString().Trim();
                message.НомерДокумента = numberDocument;
                message.ShowDialog();

            }
        }

        /// <summary>
        /// Контекстное меню ИЗМЕНИТЬ ЗАПИСЬ в таблице Рабочие документы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContextИзменитьЗапись_Click(object sender, EventArgs e)
        {
            // Переменная для хранения прирощения дат.
            int inc = 0;

            // Переменная для хранения имения файла.
            string namFile = string.Empty;

            string namFileServer = string.Empty;
            string patchToServer = string.Empty;

            // Строка для хранения запроса.
            StringBuilder builderUpdate = new StringBuilder();

            // Установим уровни изоляции транзакций.
            builderUpdate.Append("SET TRANSACTION ISOLATION LEVEL serializable begin transaction ");

            // получаем данные отображаемые в выделенной строке:
            BindingManagerBase bmb = this.BindingContext[dataGridРабочиеДокументы.DataSource, dataGridРабочиеДокументы.DataMember];
            bmb.Position = dataGridРабочиеДокументы.CurrentCell.RowNumber;
            dataGridРабочиеДокументы.Select(dataGridРабочиеДокументы.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            int idКарточки = (int)drv["id_карточки"];
            FormКарточка form = new FormКарточка(ds11, idКарточки, выбранныйГод);

            string testDsp = form.FlagDsp;

            form.ShowDialog(this);

            НомерДокумента docNum = new НомерДокумента();
            docNum = form.СледующийНомерДокумента;

            //doc.Номер = Convert.ToInt16(form.СледующийНомерДокумента.Номер);

            string strMd5 = string.Empty;

            if(form.FlagAddDoc == true)
            {
                strMd5 = "md5";
            }
            else
            {
                strMd5 = "0";
            }
                       

            // Получим номер документа который мы редактируем.
            НомерДокумента docNumNext = form.СледующийНомерДокумента;

            // Имя документа.
            string имяДокумента = form.ИмяДокумента;

            if (form.DialogResult == DialogResult.OK)
            {
                inc = form.IncrementDate;

                string sTestСтереть = form.FlagDsp;

                // Получим способ поступления документа.
                ItemСпособПоступленияДокумента способПоступленияДок = form.СпособПоступления;

                //Если установлен флаг сохранения ксерокопии документа на сервере.
                if (form.SaveDocServer == true)
                {
                    if (form.ФлагЗаписиАрхива == true)
                    {
                        // Получим путь к файлу.
                        string filePatch = form.PathFileServer;

                        // Имя программы архиватора.
                        string archiver = @"C:\Program Files\7-Zip\7z.exe";

                        // Получим имя папки которую нужно заархивировать.
                        string archive = form.FileName;// +@"\*.*";

                        // GUID составляющая названия файла.
                        string file = form.PathFileServer;

                        //string namFileS = docNumNext.Номер.ToString() + "-" + docNumNext.Префикс + "_" + file;

                        string namFileS = docNumNext.Префикс.Trim().Replace("/","-") + "_" + file;

                        //string namFile = docNumNext.Номер.ToString() + "-" + docNumNext.Префикс;

                        // Создадим своё имя файла архива содержащего архивируемую папку.
                        //string namFile = Guid.NewGuid().ToString();
                        
                        namFile = docNumNext.Префикс;

                        // Путь к временному размещению папки с архивом.
                        string patch = Application.StartupPath + @"\Archive\" + namFileS + ".7z";

                        fileName = patch;

                        namFileServer = namFile;// +".7z";

                        // Путь к папке для временного хранения архива
                        string patchDir = Application.StartupPath + @"\Archive\";

                        // Архивируем папку. (Старая реализация)
                        //Archiver.AddToArchive(archiver, archive, patch,patchDir);

                        // Путь к 7z.dll.
                        string sevenZipDll = Application.StartupPath + @"\7z.dll";
                        //Archiver.AddToArchive(sevenZipDll, archive, patch, patchDir);


                        // Архивируем папку новая реализация.


                        // Путь куда будем архивировать папку.
                        //patchToServer = patchServerFile + @"\" + namFile.Trim(); //
                        patchToServer = patchServerFile + @"\" + namFileS.Trim();

                        // Имя файла на сервере.
                        fileNameCopy = patchToServer;
                    }
                }
                else
                {
                    fileNameCopy = имяДокумента;
                    namFileServer = form.ИмяДокумента;
                }

                // ===Begin========Запишем фамилии кому отписано документ в базу данных.
                DS1.КарточкаRow row = form.строкаКарточки;

                // Разобъём строку на фамилии. (символ ,)
                string[] sКоррs = row["Резолюция"].ToString().Split(',');

                int id_карточки = idКарточки;

                string проверка = "delete ПолучателДокументовУправление " +
                                  "where idКарточки = " + id_карточки + " ";

                builderUpdate.Append(проверка);

                // Полоучим время записи.
                DateTime todoy = DateTime.Now;

                // Счётчик циклов.
                int iCount = 1;

                // Сформируем строку 
                foreach (string str in sКоррs)
                {
                    string insert = "declare @id_" + iCount + "  int " +
                                    "SELECT @id_" + iCount + " = id_получателя " +
                                    "FROM [Получатели] " +
                                    "where [ОписаниеПолучателя] = '" + str.Trim() + "' " +
                                    "INSERT INTO [ПолучателДокументовУправление] " +
                                               "([idПолучатель] " +
                                               ",[ДатаВремяЗаписи] " +
                                               ",[ОтметкаПрочтение] " +
                                               ",[ОтметкаИсполнение] " +
                                               ",[idКарточки] " +
                                               ",[РезультатВыполнения]) " +
                                         "VALUES " +
                                               "(@id_" + iCount + " " +
                                               ",'" + ДатаSQL.Дата(todoy.ToShortDateString()) + "' " +
                                               ",NULL " +
                                               ",NULL " +
                                               "," + id_карточки + " " +
                                               ",NULL) ";

                    // Добавим в запрос.
                    builderUpdate.Append(insert);

                    iCount++;
                }



                DS1TableAdapters.КарточкаTableAdapter адаптер = new RegKor.DS1TableAdapters.КарточкаTableAdapter();

                if (form.строкаКарточки.FlagCardRepeet == false)
                {

                    int Test = docNumNext.Номер;

                    ControlFlagRepeet cfr = new ControlFlagRepeet(form.строкаКарточки.id_карточки, form.строкаКарточки.FlagCardRepeet);
                    bool flag = cfr.CompareRepet();

                    if (flag == true && form.строкаКарточки.FlagCardRepeet == false)
                    {
                        НомерДокумента doc = new НомерДокумента();

                        string queryUpdate = "UPDATE [Карточка] " +
                                    "SET [id_документа] = " + form.строкаКарточки.id_документа + " " +
                                    ",[id_корреспондента] = " + form.строкаКарточки.id_корреспондента + " " +
                                    ",[ВДело] = '" + form.строкаКарточки.ВДело + "' " +
                                    ",[ДатаИсхода] = '" + ДатаSQL.Дата(form.строкаКарточки.ДатаИсхода.ToShortDateString()) + "' " +
                                    ",[ДатаПоступ] = '" + ДатаSQL.Дата(form.строкаКарточки.ДатаПоступ.ToShortDateString()) + "' " +
                                    ",[КраткоеСодержание] = '" + form.строкаКарточки.КраткоеСодержание.Trim() + "' " +
                                    ",[НаКонтроле] = '" + form.строкаКарточки.НаКонтроле + "' " +
                                    //",[НомерВход] = '" + form.строкаКарточки.НомерВход.Trim() + "' " +
                                     ",[НомерВход] = '" + docNumNext.Префикс + "' " +
                                    ",[НомерИсход] = '" + form.строкаКарточки.НомерИсход + "' " +
                                    ",[Резолюция] = '" + form.строкаКарточки.Резолюция.Trim() + "' " +
                                    ",[РезультатВыполнения] = '" + form.строкаКарточки.РезультатВыполнения.Trim() + "'  " +
                                    ",[СрокВыполнения] = '" + ДатаSQL.Дата(form.строкаКарточки.СрокВыполнения.ToShortDateString()) + "' " +
                                    //",[номерПП] = " + form.строкаКарточки.номерПП + " " +
                                    ",[номерПП] = " + docNumNext.Номер + " " +
                                    //",[номерПП] = " + docNum.Номер + " " +
                                    ",[ОписаниеКорреспондента] = '' " +
                                    ",[FlagPersonData] = '" + form.строкаКарточки.FlagPersonData + "' " +
                                    ",[FlagCardRepeet] = '" + form.строкаКарточки.FlagCardRepeet + "' " +
                                    ",[NameFileDocument] = '" + fileNameCopy + "' " +
                                    ",GuidName = '" + form.PathFileServer.Trim() + "' " +
                                     ",md5 = '" + strMd5.Trim() + "' " +
                                     ",idВидПоступленияДокумента = " + способПоступленияДок.Id + " " +
                                     ",ДСП = '" + form.FlagDsp + "' " +
                                    "WHERE id_карточки = " + form.строкаКарточки.id_карточки + " " +
                                    "DELETE FROM [КарточкаПовтор] " +
                                    "WHERE id_карточкиВходящей = " + form.строкаКарточки.id_карточки + " ";


                        builderUpdate.Append(queryUpdate);

                        //ExecuteQuery exe = new ExecuteQuery(builderUpdate.ToString().Trim());
                        //exe.Excecute();
                    }
                    if (flag == false && form.строкаКарточки.FlagCardRepeet == false)
                    {
                        адаптер.Update(form.строкаКарточки);

                        string queryUpdate = "UPDATE [Карточка] " +
                                             "set [NameFileDocument] = '" + namFileServer + "' " +
                                             ",GuidName = '"+ form.PathFileServer.Trim() +"' " +
                                             ",md5 = '" + strMd5.Trim() + "' " +
                                             ",idВидПоступленияДокумента = " + способПоступленияДок.Id + " " +
                                             ",ДСП = '" + form.FlagDsp + "' " +
                                             "WHERE id_карточки = " + form.строкаКарточки.id_карточки + " ";

                        builderUpdate.Append(queryUpdate);

                        //ExecuteQuery exe = new ExecuteQuery(builderUpdate.ToString().Trim());
                        //exe.Excecute();
                                              
                    }

                    // Получим выбранный способ поступления документа выбрал пользователь.
                    ItemСпособПоступленияДокумента способПоступленияДокумента = form.СпособПоступления;


                    // Обновим список начальников отделов и управлений которым отписан документ.
                    string queryDelete = "delete dbo.СвязующаяВидПоступленияДокПолучатели " +
                                         "where id_карточки = " + form.строкаКарточки.id_карточки + " ";

                    this.ListPerson = form.ListPerson;

                    builderUpdate.Append(queryDelete);

                    if (this.ListPerson != null)
                    {
                        // Сформируем запись в связующую таблицу документа, вида получения документа и начальниками отделов и управлений которым отписан текущий документ.
                        foreach (PersonRecepient person in this.ListPerson)
                        {
                            string insert = "INSERT INTO [СвязующаяВидПоступленияДокПолучатели] " +
                                            "([id_person] " +
                                           ",[id_ВидПоступленияДок] " +
                                           ",[id_карточки]) " +
                                           "VALUES " +
                                           "(" + person.ID + " " +
                                           "," + способПоступленияДокумента.Id + " " +
                                           "," + form.строкаКарточки.id_карточки + " ) ";

                            // Добавим в запрос.
                            builderUpdate.Append(insert);
                        }
                    }

                    //адаптер.Update(form.строкаКарточки);
                }

                if (form.строкаКарточки.FlagCardRepeet == true)
                {
                    ControlFlagRepeet cfr = new ControlFlagRepeet(form.строкаКарточки.id_карточки, form.строкаКарточки.FlagCardRepeet);
                    bool flag = cfr.CompareRepet();

                    string queryUpdate = string.Empty;

                    string stest = form.строкаКарточки.СрокВыполнения.ToShortDateString();

                    // Вносим изминение письмо было Не ИНИЦИАТИВНЫМ, а стало ИНИЦИАТИВНЫМ.
                    if (flag == false && form.строкаКарточки.FlagCardRepeet == true)
                    {

                        queryUpdate = "UPDATE [Карточка] " +
                                    "SET [id_документа] = " + form.строкаКарточки.id_документа + " " +
                                    ",[id_корреспондента] = " + form.строкаКарточки.id_корреспондента + " " +
                                    ",[ВДело] = '" + form.строкаКарточки.ВДело + "' " +
                                    ",[ДатаИсхода] = '" + ДатаSQL.Дата(form.строкаКарточки.ДатаИсхода.ToShortDateString()) + "' " +
                                    ",[ДатаПоступ] = '" + ДатаSQL.Дата(form.строкаКарточки.ДатаПоступ.ToShortDateString()) + "' " +
                                    ",[КраткоеСодержание] = '" + form.строкаКарточки.КраткоеСодержание.Trim() + "' " +
                                    ",[НаКонтроле] = '" + form.строкаКарточки.НаКонтроле + "' " +
                                    //",[НомерВход] = '" + form.строкаКарточки.НомерВход.Trim() + "' " +
                                     ",[НомерВход] = '" + docNumNext.Префикс + "' " +
                                    ",[НомерИсход] = '" + form.строкаКарточки.НомерИсход + "' " +
                                    ",[Резолюция] = '" + form.строкаКарточки.Резолюция.Trim() + "' " +
                                    ",[РезультатВыполнения] = '" + form.строкаКарточки.РезультатВыполнения.Trim() + "'  " +
                                    ",[СрокВыполнения] = '" + ДатаSQL.Дата(form.строкаКарточки.СрокВыполнения.ToShortDateString()) + "' " +
                                    //",[номерПП] = " + form.строкаКарточки.номерПП + " " +
                                    ",[номерПП] = " + docNumNext.Номер + " " +
                                     //",[номерПП] = " + docNum.Номер + " " +
                                    ",[ОписаниеКорреспондента] = '' " +
                                    ",[FlagPersonData] = '" + form.строкаКарточки.FlagPersonData + "' " +
                                    ",[FlagCardRepeet] = '" + form.строкаКарточки.FlagCardRepeet + "' " +
                                     ",[NameFileDocument] = '" + fileNameCopy + "' " +
                                     ",GuidName = '" + form.PathFileServer.Trim() + "' " +
                                      ",md5 = '" + strMd5.Trim() + "' " +
                                    "WHERE id_карточки = " + form.строкаКарточки.id_карточки + " " +
                                    " INSERT INTO [КарточкаПовтор] " +
                                    "([id_документа] " +
                                    ",[id_корреспондента] " +
                                    ",[ВДело] " +
                                    ",[ДатаИсхода] " +
                                    ",[ДатаПоступ] " +
                                    ",[КраткоеСодержание] " +
                                    ",[НаКонтроле] " +
                                    ",[НомерВход] " +
                                    ",[НомерИсход] " +
                                    ",[Резолюция] " +
                                    ",[РезультатВыполнения] " +
                                    ",[СрокВыполнения] " +
                                    ",[номерПП] " +
                                    ",[ОписаниеКорреспондента] " +
                                    ",[FlagPersonData] " +
                                    ",[FlagCardRepeet] " +
                                    ",[id_карточкиВходящей] " +
                                    ",ДатаПрирощение )" +
                                    "VALUES " +
                                    "(" + form.строкаКарточки.id_документа + " " +
                                    ", " + form.строкаКарточки.id_корреспондента + " " +
                                    ",'" + form.строкаКарточки.ВДело + "' " +
                                    ", '" + ДатаSQL.Дата(form.строкаКарточки.ДатаИсхода.ToShortDateString()) + "' " +
                                    ", '" + ДатаSQL.Дата(form.строкаКарточки.ДатаПоступ.ToShortDateString()) + "' " +
                                    ",'" + form.строкаКарточки.КраткоеСодержание.Trim() + "' " +
                                    ",'" + form.строкаКарточки.НаКонтроле + "' " +
                                    ",'" + form.строкаКарточки.НомерВход.Trim() + "' " +
                                    ",'" + form.строкаКарточки.НомерИсход + "' " +
                                    ",'" + form.строкаКарточки.Резолюция.Trim() + "' " +
                                    ",'" + form.строкаКарточки.РезультатВыполнения.Trim() + "'  " +
                                    ",'" + ДатаSQL.Дата(form.строкаКарточки.СрокВыполнения.ToShortDateString()) + "' " +
                                    //"," + form.строкаКарточки.номерПП + " " +
                                     ",[номерПП] = " + docNum.Номер + " " +
                                    ",'' " +
                                    ",'" + form.строкаКарточки.FlagPersonData + "' " +
                                    ", '" + form.строкаКарточки.FlagCardRepeet + "' " +
                                    "," + form.строкаКарточки.id_карточки + " " +
                                    ","+ inc +") ";

                        builderUpdate.Append(queryUpdate);

                        //ExecuteQuery exe = new ExecuteQuery(builderUpdate.ToString().Trim());
                        //exe.Excecute();

                        
                    }

                    // Вносим измининия в ИНИЦИАТИВНОЕ письмо.
                    if (flag == true && form.строкаКарточки.FlagCardRepeet == true)
                    {
                        queryUpdate = "UPDATE [Карточка] " +
                                    "SET [id_документа] = " + form.строкаКарточки.id_документа + " " +
                                    ",[id_корреспондента] = " + form.строкаКарточки.id_корреспондента + " " +
                                    ",[ВДело] = '" + form.строкаКарточки.ВДело + "' " +
                                    ",[ДатаИсхода] = '" + ДатаSQL.Дата(form.строкаКарточки.ДатаИсхода.ToShortDateString()) + "' " +
                                    ",[ДатаПоступ] = '" + ДатаSQL.Дата(form.строкаКарточки.ДатаПоступ.ToShortDateString()) + "' " +
                                    ",[КраткоеСодержание] = '" + form.строкаКарточки.КраткоеСодержание.Trim() + "' " +
                                    ",[НаКонтроле] = '" + form.строкаКарточки.НаКонтроле + "' " +
                                    //",[НомерВход] = '" + form.строкаКарточки.НомерВход.Trim() + "' " +
                                     ",[НомерВход] = '" + docNumNext.Префикс + "' " +
                                    ",[НомерИсход] = '" + form.строкаКарточки.НомерИсход + "' " +
                                    ",[Резолюция] = '" + form.строкаКарточки.Резолюция.Trim() + "' " +
                                    ",[РезультатВыполнения] = '" + form.строкаКарточки.РезультатВыполнения.Trim() + "'  " +
                                    ",[СрокВыполнения] = '" + ДатаSQL.Дата(form.строкаКарточки.СрокВыполнения.ToShortDateString()) + "' " +
                                    //",[номерПП] = " + form.строкаКарточки.номерПП + " " +
                            //",[номерПП] = " + docNumNext.Номер + " " +
                                     ",[номерПП] = " + docNum.Номер + " " +
                                    ",[ОписаниеКорреспондента] = '' " +
                                    ",[FlagPersonData] = '" + form.строкаКарточки.FlagPersonData + "' " +
                                    ",[FlagCardRepeet] = '" + form.строкаКарточки.FlagCardRepeet + "' " +
                                    ",[NameFileDocument] = '" + fileNameCopy + "' " +
                                    ",GuidName = '" + form.PathFileServer.Trim() + "' " +
                                     ",md5 = '" + strMd5.Trim() + "' " +
                                    "WHERE id_карточки = " + form.строкаКарточки.id_карточки + " " +
                                    " DELETE FROM [КарточкаПовтор] " +
                                    "WHERE id_карточкиВходящей = " + form.строкаКарточки.id_карточки + " " +
                                    " INSERT INTO [КарточкаПовтор] " +
                                    "([id_документа] " +
                                    ",[id_корреспондента] " +
                                    ",[ВДело] " +
                                    ",[ДатаИсхода] " +
                                    ",[ДатаПоступ] " +
                                    ",[КраткоеСодержание] " +
                                    ",[НаКонтроле] " +
                                    ",[НомерВход] " +
                                    ",[НомерИсход] " +
                                    ",[Резолюция] " +
                                    ",[РезультатВыполнения] " +
                                    ",[СрокВыполнения] " +
                                    ",[номерПП] " +
                                    ",[ОписаниеКорреспондента] " +
                                    ",[FlagPersonData] " +
                                    ",[FlagCardRepeet] " +
                                    ",[id_карточкиВходящей] " +
                                    ",ДатаПрирощение)" +
                                    "VALUES " +
                                    "(" + form.строкаКарточки.id_документа + " " +
                                    ", " + form.строкаКарточки.id_корреспондента + " " +
                                    ",'" + form.строкаКарточки.ВДело + "' " +
                                    ", '" + ДатаSQL.Дата(form.строкаКарточки.ДатаИсхода.ToShortDateString()) + "' " +
                                    ", '" + ДатаSQL.Дата(form.строкаКарточки.ДатаПоступ.ToShortDateString()) + "' " +
                                    ",'" + form.строкаКарточки.КраткоеСодержание.Trim() + "' " +
                                    ",'" + form.строкаКарточки.НаКонтроле + "' " +
                                    ",'" + form.строкаКарточки.НомерВход.Trim() + "' " +
                                    ",'" + form.строкаКарточки.НомерИсход + "' " +
                                    ",'" + form.строкаКарточки.Резолюция.Trim() + "' " +
                                    ",'" + form.строкаКарточки.РезультатВыполнения.Trim() + "'  " +
                                    ",'" + ДатаSQL.Дата(form.строкаКарточки.СрокВыполнения.ToShortDateString()) + "' " +
                                    //"," + form.строкаКарточки.номерПП + " " +
                                     ",[номерПП] = " + docNum.Номер + " " +
                                    ",'' " +
                                    ",'" + form.строкаКарточки.FlagPersonData + "' " +
                                    ", '" + form.строкаКарточки.FlagCardRepeet + "' " +
                                    "," + form.строкаКарточки.id_карточки + " " + 
                                    ","+ inc +") ";

                        builderUpdate.Append(queryUpdate);

                        //ExecuteQuery exe = new ExecuteQuery(builderUpdate.ToString().Trim());
                        //exe.Excecute();

                    }

                 }

                // Скрипт на обновления Основания передачи входящих документов СЭУ.

                 builderUpdate.Append(form.QueryPersonDateForCardInput);

                 // Завершим транзакцию.
                 builderUpdate.Append(" COMMIT TRANSACTION  ");

                 string sQueryTest = builderUpdate.ToString();


                 //// Сохраним данные.
                 ПодключитьБД connectBD = new ПодключитьБД();
                 string sCon = connectBD.СтрокаПодключения();

                // Флаг указывает успешно ли скопирован файл.
                 bool flagCopyServer = false;

                 using (SqlConnection con = new SqlConnection(sCon))
                 {
                     if (form.SaveDocServer == true)
                     {
                         try
                         {
                             if (File.Exists(fileNameCopy) == true)
                             {
                                 if (form.ФлагЗаписиАрхива == true)
                                 {
                                     //string asd = patchServerFile + @"\Move\" + namFile;

                                     //File.Move(fileNameCopy, asd);

                                     File.Delete(fileNameCopy);

                                     FileInfo file = new FileInfo(fileName);
                                     file.CopyTo(fileNameCopy, true);


                                     string patchDir = Application.StartupPath + @"\Archive\";

                                     // Удалим все файлы из директории
                                     DirectoryInfo dirInfo = new DirectoryInfo(patchDir);

                                     foreach (FileInfo fil in dirInfo.GetFiles())
                                     {
                                         fil.Delete();
                                     }
                                 }


                             }
                             else
                             {
                                 if (form.ФлагЗаписиАрхива == true)
                                 {
                                     // Скопируем файл на сервер хранения документов.
                                     File.Copy(fileName, fileNameCopy, true);
                                 }
                             }
                            
                             // Если файл скопировался успешно постави флаг в true.
                             flagCopyServer = true;
                         }
                         catch
                         {
                             MessageBox.Show("Ошибка при копировании файла");

                             flagCopyServer = false;
                         }

                         string fileTest = fileNameCopy;
                         if (File.Exists(fileNameCopy) == true)
                         {
                             con.Open();
                             SqlCommand com = new SqlCommand(builderUpdate.ToString().Trim(), con);
                             com.ExecuteNonQuery();
                         }
                     }
                     else
                     {
                         // Если файл скопировался успешно постави флаг в true.
                         flagCopyServer = true;

                         con.Open();
                         SqlCommand com = new SqlCommand(builderUpdate.ToString().Trim(), con);
                          com.ExecuteNonQuery();
                     }

                 }

                ОбновитьДанные();
            }
        }

        /// <summary>
        /// Контекстное меню ИЗМЕНИТЬ ЗАПИСЬ в таблице Исходящие документы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContextИзменитьИсходящуюЗапись_Click(object sender, EventArgs e)
        {
            // получаем данные отображаемые в выделенной строке:
            BindingManagerBase bmb = this.BindingContext[dataGridИсходящиеДокументы.DataSource, dataGridИсходящиеДокументы.DataMember];
            bmb.Position = dataGridИсходящиеДокументы.CurrentCell.RowNumber;
            dataGridИсходящиеДокументы.Select(dataGridИсходящиеДокументы.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            DataRow[] row = ds11.КарточкаИсходящая.Select("id_карточки=" + (int)drv["id_карточки"]);
            DS1.КарточкаИсходящаяRow строкаДляИзменения = (DS1.КарточкаИсходящаяRow)row[0];
            
            FormКарточкаИсходящая form = new FormКарточкаИсходящая(ds11, строкаДляИзменения, выбранныйГод);

            // Укажем что редактируем карточку.
            form.FlagEdit = true;
            
            // Передадим в форму id карточки исходящей.
            form.IdКарочкаИсходящая = строкаДляИзменения.id_карточки;

            form.ShowDialog(this);
            if (form.DialogResult == DialogResult.OK)
            {
                DS1TableAdapters.КарточкаИсходящаяTableAdapter адаптер = new RegKor.DS1TableAdapters.КарточкаИсходящаяTableAdapter();
                //адаптер.Update(form.строкаИсходящейКарточки);

                DataRow row2 = form.строкаИсходящейКарточки;

                // Получим значения выбранных id для оснований на передачу пер данныхз и входящих документов.
                List<int> listIdВходДок = form.ListIDКарточки;
                List<ОснованиеПередачи> listOP = form.ListОснованиеПередачи;

                List<int> ListIDСвязующаяКарточкаВходящаяИсходящая = form.ListIDСвязующаяКарточкаВходящаяИсходящая;
                List<int>ListIDСвязующаяЦельПолучениперсональныхДанных = form.ListIDСвязующаяЦельПолучениперсональныхДанных;

                StringBuilder builder = new StringBuilder();

                int id_ВходящегоДокумента;
                if (row2["id_ВходящегоДокумента"] == DBNull.Value)
                {
                    string query = string.Empty;

                    // Проверим работаем с ДСП или нет.
                    if (Convert.ToBoolean(form.FlagDsp) == false)
                    {
                        // Изменим исходящие документы.
                        query = "UPDATE [КарточкаИсходящая] " +
                                       "SET [Дата] = '" + ДатаSQL.Дата(Convert.ToDateTime(row2["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",[НомерКомитета] = '" + row2["НомерКомитета"] + "' " +
                                       ",[id_Подразделения] = " + row2["id_Подразделения"] + " " +
                                       ",[НомерНоменклатурный] = '" + row2["НомерНоменклатурный"] + "' " +
                                       ",[НомерПорядковый] = " + row2["НомерПорядковый"] + " " +
                                       ",[id_Адресата] = " + row2["id_Адресата"] + " " +
                                       ",[Содержание] = '" + row2["Содержание"] + "' " +
                            //",[id_ВходящегоДокумента] = " + id_ВходящегоДокумента + " " +
                                       ",[ОписаниеКорреспондента] = '" + row2["ОписаниеКорреспондента"] + "' " +
                                       ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                                       ",ДСП = '" + form.FlagDsp + "' " +
                                       "where id_карточки = " + Convert.ToInt32(row2["id_карточки"]) + " ";
                    }
                    else
                    {
                        // Изменим исходящие документы.
                        query = "UPDATE [КарточкаИсходящая] " +
                                       "SET [Дата] = '" + ДатаSQL.Дата(Convert.ToDateTime(row2["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",[НомерКомитета] = '" + row2["НомерКомитета"] + "' " +
                                       ",[id_Подразделения] = " + row2["id_Подразделения"] + " " +
                                       ",[НомерНоменклатурный] = '" + row2["НомерНоменклатурный"] + "' " +
                                       ",[НомерПорядковый] = " + row2["НомерПорядковый"] + " " +
                                       ",[id_Адресата] = " + row2["id_Адресата"] + " " +
                                       ",[Содержание] = '" + row2["Содержание"] + "' " +
                            //",[id_ВходящегоДокумента] = " + id_ВходящегоДокумента + " " +
                                       ",[ОписаниеКорреспондента] = '" + row2["ОписаниеКорреспондента"] + "' " +
                                       ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                                       ",ДСП = '" + form.FlagDsp + "' " +
                                       ", ДспDesc = 'ДСП' " +
                                       "where id_карточки = " + Convert.ToInt32(row2["id_карточки"]) + " ";
                    }

                    builder.Append(query);

                    // Счётчик для цикла.
                    //int iCountlistOP = 0;
                    string queryDelete = "DELETE FROM [СвязующаяЦельПолучениперсональныхДанных] " +
                                        "WHERE id_карточки = " + Convert.ToInt32(row2["id_карточки"]) + " ";

                    builder.Append(queryDelete);

                    // обновим данные по связующим таблицам.
                    // Запрос на всатвку id в связующую таблицу ЦельПолученияПерсДанных.
                    string sUpdate = string.Empty;
                    sUpdate = String.Format(form.QueryInsert.Trim(), " "+ Convert.ToInt32(row2["id_карточки"]) + " ");

                    builder.Append(sUpdate);

                  

                    // Удалим строки из связывающей таблицы.
                    string delete = "DELETE FROM СвязующаяКарточкаВходящаяИсходящая " +
                                    "WHERE id_карточкаИсходящая = " + Convert.ToInt32(row2["id_карточки"]) + " ";

                    builder.Append(delete);

                    // Заполним связующую таблицу Карточка входящаяИсходящая.
                    foreach (int idВх in listIdВходДок)
                    {
                        string queryIdВх = "INSERT INTO [СвязующаяКарточкаВходящаяИсходящая] " +
                                           "([id_карточкаВходящая] " +
                                           ",[id_карточкаИсходящая]) " +
                                           "VALUES " +
                                           "(" + idВх + " " +
                                           "," + Convert.ToInt32(row2["id_карточки"]) + " ) " +
                                           "update Карточка " +
                                           "set РезультатВыполнения = '" + form.НомерИсходящий.Trim() + "' " +
                                           "where id_карточки = " + idВх + " ";

                        //iiCount++;

                        builder.Append(queryIdВх);
                    }

                }
                else
                {
                    id_ВходящегоДокумента = Convert.ToInt32(row2["id_ВходящегоДокумента"]);

                    string query = string.Empty;

                    //// Изменим исходящие документы.
                    //string query = "UPDATE [КарточкаИсходящая] " +
                    //                "SET [Дата] = '" + ДатаSQL.Дата(Convert.ToDateTime(row2["Дата"]).ToShortDateString().Trim()) + "' " +
                    //                ",[НомерКомитета] = '" + row2["НомерКомитета"] + "' " +
                    //                ",[id_Подразделения] = " + row2["id_Подразделения"] + " " +
                    //                ",[НомерНоменклатурный] = '" + row2["НомерНоменклатурный"] + "' " +
                    //                ",[НомерПорядковый] = " + row2["НомерПорядковый"] + " " +
                    //                ",[id_Адресата] = " + row2["id_Адресата"] + " " +
                    //                ",[Содержание] = '" + row2["Содержание"] + "' " +
                    //                ",[id_ВходящегоДокумента] = " + id_ВходящегоДокумента + " " +
                    //                ",[ОписаниеКорреспондента] = '" + row2["ОписаниеКорреспондента"] + "' " +
                    //                ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                    //                 ",ДСП = '" + form.FlagDsp + "' " +
                    //                "where id_карточки = " + Convert.ToInt32(row2["id_карточки"]) + " ";

                    // Проверим работаем с ДСП или нет.
                    if (Convert.ToBoolean(form.FlagDsp) == false)
                    {
                        // Изменим исходящие документы.
                        query = "UPDATE [КарточкаИсходящая] " +
                                       "SET [Дата] = '" + ДатаSQL.Дата(Convert.ToDateTime(row2["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",[НомерКомитета] = '" + row2["НомерКомитета"] + "' " +
                                       ",[id_Подразделения] = " + row2["id_Подразделения"] + " " +
                                       ",[НомерНоменклатурный] = '" + row2["НомерНоменклатурный"] + "' " +
                                       ",[НомерПорядковый] = " + row2["НомерПорядковый"] + " " +
                                       ",[id_Адресата] = " + row2["id_Адресата"] + " " +
                                       ",[Содержание] = '" + row2["Содержание"] + "' " +
                            ",[id_ВходящегоДокумента] = " + id_ВходящегоДокумента + " " +
                                       ",[ОписаниеКорреспондента] = '" + row2["ОписаниеКорреспондента"] + "' " +
                                       ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                                       ",ДСП = '" + form.FlagDsp + "' " +
                                       "where id_карточки = " + Convert.ToInt32(row2["id_карточки"]) + " ";
                    }
                    else
                    {
                        // Изменим исходящие документы.
                        query = "UPDATE [КарточкаИсходящая] " +
                                       "SET [Дата] = '" + ДатаSQL.Дата(Convert.ToDateTime(row2["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",[НомерКомитета] = '" + row2["НомерКомитета"] + "' " +
                                       ",[id_Подразделения] = " + row2["id_Подразделения"] + " " +
                                       ",[НомерНоменклатурный] = '" + row2["НомерНоменклатурный"] + "' " +
                                       ",[НомерПорядковый] = " + row2["НомерПорядковый"] + " " +
                                       ",[id_Адресата] = " + row2["id_Адресата"] + " " +
                                       ",[Содержание] = '" + row2["Содержание"] + "' " +
                            ",[id_ВходящегоДокумента] = " + id_ВходящегоДокумента + " " +
                                       ",[ОписаниеКорреспондента] = '" + row2["ОписаниеКорреспондента"] + "' " +
                                       ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                                       ",ДСП = '" + form.FlagDsp + "' " +
                                       ", ДспDesc = 'ДСП' " +
                                       "where id_карточки = " + Convert.ToInt32(row2["id_карточки"]) + " ";
                    }

                    builder.Append(query);

                    string updateQuery = "UPDATE [Карточка] " +
                                         "SET [ВДело] = 'True' " +
                                         "WHERE id_карточки = "+ id_ВходящегоДокумента +" ";

                    builder.Append(updateQuery);

                    // Счётчик для цикла.
                    //int iCountlistOP = 0;
                    string queryDelete = "DELETE FROM [СвязующаяЦельПолучениперсональныхДанных] " +
                                       "WHERE id_карточки = " + Convert.ToInt32(row2["id_карточки"]) + " ";

                    builder.Append(queryDelete);

                    // обновим данные по связующим таблицам.
                    // Запрос на всатвку id в связующую таблицу ЦельПолученияПерсДанных.
                    foreach (ОснованиеПередачи itm in listOP)
                    {

                        string queryIns = "INSERT INTO [СвязующаяЦельПолучениперсональныхДанных] " +
                                       "([id_карточки] " +
                                       ",[id_ОснованиеПередачи]) " +
                                       "VALUES " +
                                       "('" + Convert.ToInt32(row2["id_карточки"]) + "' " +
                                       ",'" + itm.Id_основаниеПередачи + "' ) ";


                        //iCountlistOP++;

                        builder.Append(queryIns);
                    }

                    //int iiCount = 0;

                    // Удалим строки из связывающей таблицы.
                    string delete = "DELETE FROM СвязующаяКарточкаВходящаяИсходящая " +
                                    "WHERE id_карточкаИсходящая = " + Convert.ToInt32(row2["id_карточки"]) + " ";
                    
                    builder.Append(delete);

                    // Заполним связующую таблицу Карточка входящаяИсходящая.
                    foreach (int idВх in listIdВходДок)
                    {
                        string queryIdВх = "INSERT INTO [СвязующаяКарточкаВходящаяИсходящая] " +
                                           "([id_карточкаВходящая] " +
                                           ",[id_карточкаИсходящая]) " +
                                           "VALUES " +
                                           "(" + idВх + " " +
                                           "," + Convert.ToInt32(row2["id_карточки"]) + " ) ";

                        //iiCount++;

                        builder.Append(queryIdВх);
                    }
                }

                // Если письмо уже помечено как с передачей персональных данных.
                if (Convert.ToBoolean(row2["FlagPersonData"]) == true)
                {

                    string номерПодразделения = string.Empty;

                    DataRow[] rowsSelect = ds11.ПодразделенияКомитета.Select("id_подразделения= " + Convert.ToInt32(row2["id_Подразделения"]) + " ");
                    foreach (DataRow item in rowsSelect)
                    {
                        номерПодразделения = item["НомерПодразделения"].ToString().Trim();
                    }

                    string результатВыполнения = "Дан ответ. № исх. документа " + row2["НомерКомитета"].ToString().Trim() + "-" + row2["НомерНоменклатурный"].ToString().Trim() + "-" + номерПодразделения + "/" + row2["НомерПорядковый"].ToString().Trim();

                    if (row2["id_ВходящегоДокумента"] != DBNull.Value)
                    {
                        // Установим флаг в TRUE.
                        string queryUpdate = " UPDATE [Карточка] " +
                                             "SET РезультатВыполнения = '" + результатВыполнения + "' " +
                            //"FlagPersonData = '" + row["FlagPersonData"] + "' " +
                                             ",ВДело = 'True' " +
                                              ",ДСП = '" + form.FlagDsp + "' " +
                                             "where id_карточки = " + Convert.ToInt32(row2["id_ВходящегоДокумента"]) + " ";
                        // Соберём строки запросаов на добавление записи и на редактирование в единую строку, чтобы выполнить всё в одной транзакции.
                        builder.Append(queryUpdate);
                    }
                }
                else
                {
                    string номерПодразделения = string.Empty;

                    DataRow[] rowsSelect = ds11.ПодразделенияКомитета.Select("id_подразделения= " + Convert.ToInt32(row2["id_Подразделения"]) + " ");
                    foreach (DataRow item in rowsSelect)
                    {
                        номерПодразделения = item["НомерПодразделения"].ToString().Trim();
                    }

                    string результатВыполнения = "Дан ответ. № исх. документа " + row2["НомерКомитета"].ToString().Trim() + "-" + row2["НомерНоменклатурный"].ToString().Trim() + "-" + номерПодразделения + "/" + row2["НомерПорядковый"].ToString().Trim();

                    if (row2["id_ВходящегоДокумента"] != DBNull.Value)
                    {
                        // Установим флаг в TRUE.
                        string queryUpdate = " UPDATE [Карточка] " +
                                             "SET РезультатВыполнения = '" + результатВыполнения + "' " +
                            //"FlagPersonData = '" + row["FlagPersonData"] + "' " +
                                             ",ВДело = 'True' " +
                                              ",ДСП = '" + form.FlagDsp + "' " +
                                             "where id_карточки = " + Convert.ToInt32(row2["id_ВходящегоДокумента"]) + " ";
                        // Соберём строки запросаов на добавление записи и на редактирование в единую строку, чтобы выполнить всё в одной транзакции.
                        builder.Append(queryUpdate);
                    }
                }

              
                
                // Сохраним изменение в БД.
                ПодключитьБД коннект = new ПодключитьБД();
                string sConnect = коннект.СтрокаПодключения();

                SqlConnection con = new SqlConnection(sConnect);
                SqlCommand com = new SqlCommand(builder.ToString(), con);
                con.Open();
                com.ExecuteNonQuery();
                con.Close();
                    
                                   
                ОбновитьДанные();
            }
        }

        /// <summary>
        /// Контекстное меню ИЗМЕНИТЬ ЗАПИСЬ в таблице документы в деле
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContextИзменитьЗапись_Click2(object sender, EventArgs e)
        {
            // получаем данные отображаемые в выделенной строке:
            BindingManagerBase bmb = this.BindingContext[dataGridДокументыВДеле.DataSource, dataGridДокументыВДеле.DataMember];
            bmb.Position = dataGridДокументыВДеле.CurrentCell.RowNumber;
            dataGridДокументыВДеле.Select(dataGridДокументыВДеле.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            int idКарточки = (int)drv["id_карточки"];
            FormКарточка form = new FormКарточка(ds11, idКарточки, выбранныйГод);
            form.ShowDialog(this);
            if (form.DialogResult == DialogResult.OK)
            {

                DS1TableAdapters.КарточкаTableAdapter адаптер = new RegKor.DS1TableAdapters.КарточкаTableAdapter();

                // Стрка для хранения строки обращения к БД для обновления БД.
                 StringBuilder builderUpdate = new StringBuilder();

                // Получим способ поступления документа.
                ItemСпособПоступленияДокумента способПоступленияДок = form.СпособПоступления;

                // Построим строку запроса в единой транзакции.
                builderUpdate.Append("SET TRANSACTION ISOLATION LEVEL serializable begin transaction  ");

                string queryUpdate = "UPDATE [Карточка] " +
                                     "SET [id_документа] = " + form.строкаКарточки.id_документа + " " +
                                     ",[id_корреспондента] = " + form.строкаКарточки.id_корреспондента + " " +
                                     ",[ВДело] = '" + form.строкаКарточки.ВДело + "' " +
                                     ",[ДатаИсхода] = '" + ДатаSQL.Дата(form.строкаКарточки.ДатаИсхода.ToShortDateString()) + "' " +
                                     ",[ДатаПоступ] = '" + ДатаSQL.Дата(form.строкаКарточки.ДатаПоступ.ToShortDateString()) + "' " +
                                     ",[КраткоеСодержание] = '" + form.строкаКарточки.КраткоеСодержание.Trim() + "' " +
                                     ",[НаКонтроле] = '" + form.строкаКарточки.НаКонтроле + "' " +
                    //",[НомерВход] = '" + form.строкаКарточки.НомерВход.Trim() + "' " +
                    //",[НомерВход] = '" + docNumNext.Префикс + "' " +
                                     ",[НомерИсход] = '" + form.строкаКарточки.НомерИсход + "' " +
                                     ",[Резолюция] = '" + form.строкаКарточки.Резолюция.Trim() + "' " +
                                     ",[РезультатВыполнения] = '" + form.строкаКарточки.РезультатВыполнения.Trim() + "'  " +
                                     ",[СрокВыполнения] = '" + ДатаSQL.Дата(form.строкаКарточки.СрокВыполнения.ToShortDateString()) + "' " +
                    //",[номерПП] = " + form.строкаКарточки.номерПП + " " +
                    //",[номерПП] = " + docNumNext.Номер + " " +
                    //",[номерПП] = " + docNum.Номер + " " +
                                     ",[ОписаниеКорреспондента] = '' " +
                                     ",[FlagPersonData] = '" + form.строкаКарточки.FlagPersonData + "' " +
                                     ",[FlagCardRepeet] = '" + form.строкаКарточки.FlagCardRepeet + "' " +
                                      ",[NameFileDocument] = '" + fileNameCopy + "' " +
                                      ",GuidName = '" + form.PathFileServer.Trim() + "' " +
                                      ",idВидПоступленияДокумента = " + form.СпособПоступления.Id + " " +
                    // ",md5 = '" + strMd5.Trim() + "' " +
                                     "WHERE id_карточки = " + form.строкаКарточки.id_карточки + " ";

                // Добавим строку запроса на обновление.
                builderUpdate.Append(queryUpdate);

                // Проверим пустая строка или нет.


                // Запрос на удаление связанных и добавления новых запсией.
                builderUpdate.Append(form.QueryPersonDateForCardInput);

                builderUpdate.Append("COMMIT TRANSACTION ");

                // Строка на выполнение запроса.
                string queryUpdateCard = builderUpdate.ToString();


                ExecuteQuery exec = new ExecuteQuery(queryUpdateCard);
                exec.Excecute();

                //адаптер.Update(form.строкаКарточки);
                ОбновитьДанные();
            }
            this.Refresh();
        }

        /// <summary>
        /// Контекстное меню УДАЛИТЬ ЗАПИСЬ в таблице Рабочие документы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContextУдалитьЗапись_Click(object sender, EventArgs e)
        {
            // получаем данные отображаемые в выделенной строке:
            BindingManagerBase bmb = this.BindingContext[dataGridРабочиеДокументы.DataSource, dataGridРабочиеДокументы.DataMember];
            bmb.Position = dataGridРабочиеДокументы.CurrentCell.RowNumber;
            dataGridРабочиеДокументы.Select(dataGridРабочиеДокументы.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            DialogResult выборПользователя = MessageBox.Show(this, "Вы действительно хотите удалить документ '" + drv["ОписаниеДокумента"] + "' от корресподента '" + drv["ОписаниеКорреспондента"] + "'?", "Удаление записи", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2);
            if (выборПользователя == DialogResult.Yes)
            {
                int idКарточки = (int)drv["id_карточки"];

                DataRow[] rows = ds11.Карточка.Select("id_карточки = " + idКарточки);
                //rows[0].Delete();

                StringBuilder build = new StringBuilder();

                string queryDelete = "DELETE FROM [Карточка] " +
                                     "WHERE id_карточки = "+ idКарточки +" ";

                build.Append(queryDelete);

                string queryDeleteD = "DELETE FROM КарточкаПовтор " +
                                     "WHERE id_карточкиВходящей = " + idКарточки + " ";

                build.Append(queryDeleteD);

                 //а вот здесь затык в КарточкаПовтор нужен id из табл Карточка

                ExecuteQuery eq = new ExecuteQuery(build.ToString().Trim());
                eq.Excecute();

                ОбновитьДанные();
            }
            this.Refresh();
        }

        /// <summary>
        /// Контекстное меню УДАЛИТЬ ЗАПИСЬ в таблице документы в деле
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContextУдалитьЗапись_Click2(object sender, EventArgs e)
        {
            // получаем данные отображаемые в выделенной строке:
            BindingManagerBase bmb = this.BindingContext[dataGridДокументыВДеле.DataSource, dataGridДокументыВДеле.DataMember];
            bmb.Position = dataGridДокументыВДеле.CurrentCell.RowNumber;
            dataGridДокументыВДеле.Select(dataGridДокументыВДеле.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            DialogResult выборПользователя = MessageBox.Show(this, "Вы действительно хотите удалить документ '" + drv["ОписаниеДокумента"] + "' от корресподента '" + drv["ОписаниеКорреспондента"] + "'?", "Удаление записи", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2);
            if (выборПользователя == DialogResult.Yes)
            {
                int idКарточки = (int)drv["id_карточки"];

                DataRow[] rows = ds11.Карточка.Select("id_карточки = " + idКарточки);
                rows[0].Delete();

                ОбновитьДанные();
            }
            this.Refresh();
        }

        /// <summary>
        /// Контекстное меню УДАЛИТЬ ЗАПИСЬ в таблице исходящие документы
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContextУдалитьИсходящуюЗапись_Click(object sender, EventArgs e)
        {
            // получаем данные отображаемые в выделенной строке:
            BindingManagerBase bmb = this.BindingContext[dataGridИсходящиеДокументы.DataSource, dataGridИсходящиеДокументы.DataMember];
            bmb.Position = dataGridИсходящиеДокументы.CurrentCell.RowNumber;
            dataGridИсходящиеДокументы.Select(dataGridИсходящиеДокументы.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            DialogResult выборПользователя = MessageBox.Show(this, "Вы действительно хотите удалить документ от '" + drv["ОписаниеПодразделения"] + "' \nдля\n'" + drv["ОписаниеАдресата"] + "'?", "Удаление записи", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2);
            if (выборПользователя == DialogResult.Yes)
            {
                int idКарточки = (int)drv["id_карточки"];

                if (drv["id_ВходящегоДокумента"] != System.DBNull.Value)
                {
                    int idВходДокумента = (int)drv["id_ВходящегоДокумента"];
                    DataRow[] отмена = ds11.Карточка.Select("id_карточки=" + idВходДокумента);
                    if (отмена.Length > 0)
                    {
                        отмена[0]["ВДело"] = false;
                        отмена[0]["РезультатВыполнения"] = "";
                    }
                }


                DataRow[] rows = ds11.КарточкаИсходящая.Select("id_карточки = " + idКарточки);
                rows[0].Delete();

                string query = "DELETE FROM КарточкаИсходящая " +
                               "WHERE id_карточки = "+ idКарточки +" ";

                Classess.ПодключитьБД clConn = new ПодключитьБД();
                string sConn = clConn.СтрокаПодключения();

                SqlConnection con = new SqlConnection(sConn);
                con.Open();
                SqlCommand com = new SqlCommand(query, con);
                com.ExecuteNonQuery();
                con.Close();

                ОбновитьДанные();
            }
            this.Refresh();
        }

        private void menuItemContextПечатьКарточки_Click(object sender, EventArgs e)
        {
            ПечатьКарточки();
            this.Refresh();
        }

        /// <summary>
        /// Вызывает окно сохранения в файл
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemЗакрыть_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Вызывает справочник "Документы"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemСправочникиДокументы_Click(object sender, System.EventArgs e)
        {
            FormДокументы form = new FormДокументы();
            this.Enabled = false;
            form.ShowDialog(this);
            this.Refresh();
            ПодключитьсяПолучитьДанные();
            this.Enabled = true;
            this.Refresh();
        }

        /// <summary>
        /// Вызывает справочник "Корреспонденты"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemСправочникиКорреспонденты_Click(object sender, System.EventArgs e)
        {
            FormКорреспонденты form = new FormКорреспонденты();
            this.Refresh();
            this.Enabled = false;
            form.ShowDialog(this);
            this.Refresh();
            ПодключитьсяПолучитьДанные();
            this.Enabled = true;
            this.Refresh();
        }

        /// <summary>
        /// Вызывает справочник "Получатели"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemСправочникиПолучатели_Click(object sender, System.EventArgs e)
        {
            FormПолучатели form = new FormПолучатели();
            this.Refresh();
            this.Enabled = false;
            form.ShowDialog(this);
            this.Refresh();
            ПодключитьсяПолучитьДанные();
            this.Enabled = true;
            this.Refresh();
        }

        ///// <summary>
        ///// Вызывает справочник "Адресаты"
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void menuItemСправочникАдресаты_Click ( object sender, EventArgs e )
        //{
        //    FormАдресатыИсходящие form = new FormАдресатыИсходящие();
        //    this.Refresh( );
        //    this.Enabled = false;
        //    form.ShowDialog( this );
        //    this.Refresh( );
        //    ПодключитьсяПолучитьДанные( );
        //    this.Enabled = true;
        //    this.Refresh( );	
        //}

        private void menuItem4_Click(object sender, System.EventArgs e)
        {
            this.Enabled = false;
            FormДиапазонДат frm = new FormДиапазонДат(this.ds11);
            frm.ShowDialog(this);
            this.Enabled = true;
        }

        /// <summary>
        /// Событие Click меню "Статистика отправления корреспонденции"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem5_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            FormДиапазонДат2 frm = new FormДиапазонДат2(this.ds11);
            frm.ShowDialog(this);
            this.Enabled = true;
        }

        private void menuItemКонтрольныеУведомления_Click(object sender, EventArgs e)
        {
            ПечатьКонтрольныхУведомлений();
        }

        private void menuItemПросрочДокументы_Click(object sender, EventArgs e)
        {
            потокОжидания = new System.Threading.Thread(new System.Threading.ThreadStart(ЗапуститьФормуОжидания));
            потокОжидания.Start();

            ПечатьПросроченныхДокументов();
        }

        private void tabControlТипыДокументов_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControlТипыДокументов.SelectedTab.Text == "Исходящие")
            {
                //MessageBox.Show( "Активная вкладка " + tabControlТипыДокументов.SelectedTab.Name);
                menuItemКонтрольныеУведомления.Enabled = false;
                menuItemПросрочДокументы.Enabled = false;

                menuItemContextПечатьКарточки.Enabled = false;
                menuItem4.Enabled = false;
                menuItem6.Enabled = true;
            }
            if (tabControlТипыДокументов.SelectedTab.Text == "Входящие")
            {
                menuItemКонтрольныеУведомления.Enabled = true;
                menuItemПросрочДокументы.Enabled = true;
                
                menuItemContextПечатьКарточки.Enabled = true;
                menuItem4.Enabled = true;
                menuItem6.Enabled = false;
            }
        }

        #endregion

        private void menuItemСправочникиПодразделения_Click(object sender, EventArgs e)
        {
            FormПодразделения form = new FormПодразделения(this.ds11);
            this.Refresh();
            this.Enabled = false;
            form.ShowDialog(this);
            this.Refresh();
            ПодключитьсяПолучитьДанные();
            this.Enabled = true;
            this.Refresh();
        }

        private void dataGridРабочиеДокументы_Resize(object sender, EventArgs e)
        {
            int ширинаТаблицы = dataGridРабочиеДокументы.Width;

            int шДок = dataGridTextBoxColumn1.Width;
            int шКорр = dataGridTextBoxColumn2.Width;
            int шДт1 = dataGridTextBoxColumn3.Width;
            int шНом1 = dataGridTextBoxColumn4.Width;
            int шДт2 = dataGridTextBoxColumn5.Width;
            int шНом2 = dataGridTextBoxColumn6.Width;
            int шСодерж = dataGridTextBoxColumn7.Width;
            int шКонтр = dataGridTextBoxColumn8.Width;

            dataGridTextBoxColumn7.Width = ширинаТаблицы - 20 - шДок - шКорр - шДт1 - шНом1 - шДт2 - шНом2 - шКонтр;
        }

        private void dataGridДокументыВДеле_Resize(object sender, EventArgs e)
        {
            int ширинаТаблицы = dataGridДокументыВДеле.Width;

            int шДок = dataGridTextBoxColumn9.Width;
            int шКорр = dataGridTextBoxColumn10.Width;
            int шДт1 = dataGridTextBoxColumn11.Width;
            int шНом1 = dataGridTextBoxColumn12.Width;
            int шДт2 = dataGridTextBoxColumn13.Width;
            int шНом2 = dataGridTextBoxColumn14.Width;
            int шСодерж = dataGridTextBoxColumn15.Width;
            int шРезол = dataGridTextBoxColumn16.Width;

            dataGridTextBoxColumn15.Width = ширинаТаблицы - 20 - шДок - шКорр - шДт1 - шНом1 - шДт2 - шНом2 - шРезол;
        }

        private void dataGridИсходящиеДокументы_Resize(object sender, EventArgs e)
        {
            int ширинаТаблицы = dataGridИсходящиеДокументы.Width;

            int шДаты = dataGridTextBoxColumnИсхДокДатаИсхода.Width;
            int шНомер = dataGridTextBoxColumnИсхДокНомер.Width;
            int шАдрес = dataGridTextBoxColumnИсхДокОписаниеАдресата.Width;
            int шСодерж = dataGridTextBoxColumnИсхДокСодержание.Width;
            int шВхДокт = dataGridTextBoxColumnИсхДокНомерВходДокта.Width;

            dataGridTextBoxColumnИсхДокСодержание.Width = ширинаТаблицы - 20 - шДаты - шНомер - шАдрес - шВхДокт;
        }



        /// <summary>
        /// Конструирует строку для фильтрации исходящих документов
        /// </summary>
        private string ФильтрИД
        {
            get
            {
                string фильтр = string.Empty;

                if (textBoxСтрокаПоискаИсходящихДокументов.Text.Trim().ToLower() != "ДСП".ToLower().Trim())
                {
                    фильтр = "(ТекстовыйНомер LIKE '%" + textBoxСтрокаПоискаИсходящихДокументов.Text + "%'" +
                                                        " OR Содержание LIKE '%" + textBoxСтрокаПоискаИсходящихДокументов.Text + "%'" +
                                                        " OR ОписаниеПодразделения LIKE '%" + textBoxСтрокаПоискаИсходящихДокументов.Text + "%'" +
                                                        " OR ОписаниеРуководителя LIKE '%" + textBoxСтрокаПоискаИсходящихДокументов.Text + "%'" +
                                                        " OR ОписаниеАдресата LIKE '%" + textBoxСтрокаПоискаИсходящихДокументов.Text + "%')";
                }
                else
                {
                    фильтр = "ДспDesc = 'ДСП' ";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Весь год")
                {
                    DateTime min = Convert.ToDateTime("01.12." + выбранныйГод + "");
                    DateTime max = Convert.ToDateTime("31.12." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Январь")
                {
                    DateTime min = Convert.ToDateTime("01.12." + выбранныйГод + "");
                    DateTime max = Convert.ToDateTime("31.01." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Февраль")
                {
                    DateTime min = Convert.ToDateTime("01.02." + selectedYear.ToString() + "");
                    DateTime max;
                    if (DateTime.IsLeapYear(DateTime.Now.Year))
                    {
                        max = Convert.ToDateTime("29.02." + selectedYear.ToString() + "");
                    }
                    else
                    {
                        max = Convert.ToDateTime("28.02." + selectedYear.ToString() + "");
                    }
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Март")
                {
                    DateTime min = Convert.ToDateTime("01.03." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.03." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Апрель")
                {
                    DateTime min = Convert.ToDateTime("01.04." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("30.04." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Май")
                {
                    DateTime min = Convert.ToDateTime("01.05." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.05." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Июнь")
                {
                    DateTime min = Convert.ToDateTime("01.06." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("30.06." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Июль")
                {
                    DateTime min = Convert.ToDateTime("01.07." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.07." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Август")
                {
                    DateTime min = Convert.ToDateTime("01.08." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.08." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Сентябрь")
                {
                    DateTime min = Convert.ToDateTime("01.09." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("30.09." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Октябрь")
                {
                    DateTime min = Convert.ToDateTime("01.10." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.10." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Ноябрь")
                {
                    DateTime min = Convert.ToDateTime("01.11." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("30.11." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }
                if (comboBoxФильтрИДПоДате.SelectedItem.ToString() == "Декабрь")
                {
                    DateTime min = Convert.ToDateTime("01.12." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.12." + selectedYear.ToString() + "");
                    фильтр += " AND Дата>='" + min + "' AND Дата<='" + max + "'";
                }

                return фильтр;
            }
        }

        private void comboBoxФильтрИДПоДате_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataView view = (DataView)dataGridИсходящиеДокументы.DataSource;
            view.RowFilter = ФильтрИД;
            labelОтобраноДокументовПоискомИсходящихДокументов.Text = "Отобрано документов: " + view.Count;
        }

        //private void dataGridРабочиеДокументы_Navigate(object sender, NavigateEventArgs ne)
        //{

        //}

        //private void dataGridИсходящиеДокументы_Navigate(object sender, NavigateEventArgs ne)
        //{

        //}

        

        private void checkBoxКорреспонденты_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBoxКорреспонденты.Checked == true)
            {
                this.comboBoxКорреспонденты.Visible = false;
            }
            else
            {
                this.comboBoxКорреспонденты.Visible = true;
            }
        }

        private void menuItem6_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            //FormДиапазонДатОтправка frm = new FormДиапазонДатОтправка(this.ds11);
            FormДиапазонДатОтправка frm = new FormДиапазонДатОтправка();
            frm.ShowDialog(this);
            this.Enabled = true;
        }

        private void menuItem8_Click(object sender, EventArgs e)
        {
            FormГодКонфигурации годКонфигураци = new FormГодКонфигурации();
            годКонфигураци.ShowDialog(this);
        }

        private void FormГлавная_Load(object sender, EventArgs e)
        {
            //При загрузки сделаем пункт меню с отчётами не активным
            menuItem6.Enabled = false;

            if (ConfigurationSettings.AppSettings["AddDubleNumberDoc"] == "1")
            {
                this.menuItem17.Visible = true;
            }
            else
            {
                this.menuItem17.Visible = false;
            }
        }

        private void menuItem10_Click(object sender, EventArgs e)
        {
            FormSelectDatePerson personDate = new FormSelectDatePerson();
            //personDate.MdiParent = this;
            personDate.ShowDialog();
        }

        private void menuItem11_Click(object sender, EventArgs e)
        {
            // Откроем окно редактирования персональных данных.
            FormPD formPD = new FormPD();
            formPD.Show();
        }

        private void menuItem12_Click(object sender, EventArgs e)
        {
            FormПолучениеПерсональныхДанных formPD = new FormПолучениеПерсональныхДанных();
            formPD.Show();
        }

        private void menuItem13_Click(object sender, EventArgs e)
        {
            FormПечатьКарточки form = new FormПечатьКарточки();
            form.ТекущийГод = Дата.ПервыйДень(selectedYear.ToString());
            form.БудущийГод = следующаяДата;
            form.ShowDialog(this);
            this.Enabled = true;
        }

        private void FormГлавная_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Очистим директорию Журнал.

            // Получим путь к исполняемому файлу.
            string pathExe = Application.StartupPath;

            string patch = pathExe;
        }

        private void contextMenu1_Popup(object sender, EventArgs e)
        {

        }

        private void menuItem15_Click(object sender, EventArgs e)
        {
            FormДокументооборот form = new FormДокументооборот();
            form.Show();
        }

        private NumOutputCardVipNet GetNumDocOutVipNet(string strGuid)
        {

            // Экземпляр вспомогательного класса.
            NumOutputCardVipNet numCard = new NumOutputCardVipNet();

            string num = string.Empty;

            string query = "select id_карточки,НомерПорядковый from dbo.КарточкаИсходящая " +
                           "where [GUID] = '" + strGuid.Trim() + "' ";

            ПодключитьБД strCon = new ПодключитьБД();
            SqlConnection con = new SqlConnection(strCon.СтрокаПодключения());
            con.Open();

            SqlCommand com = new SqlCommand(query, con);
            SqlDataReader read = com.ExecuteReader();

            while (read.Read())
            {
                numCard.Id = Convert.ToInt32(read["id_карточки"]);
                numCard.НомерПорядковый = read["НомерПорядковый"].ToString().Trim();

            }

            return numCard;

        }

        /// <summary>
        /// Получим номер исходящего документа.
        /// </summary>
        /// <returns></returns>
        private string GetNumDocOut(string strGuid)
        {
            
            // Экземпляр вспомогательного класса.
            NumOutputCardVipNet numCard = new NumOutputCardVipNet();
            
            string num = string.Empty;

            string query = "select id_карточки,НомерПорядковый from dbo.КарточкаИсходящая " +
                           "where [GUID] = '" + strGuid.Trim() + "' ";

            ПодключитьБД strCon = new ПодключитьБД();
            SqlConnection con = new SqlConnection(strCon.СтрокаПодключения());
            con.Open();

            SqlCommand com = new SqlCommand(query, con);
            SqlDataReader read = com.ExecuteReader();

            while (read.Read())
            {
                num = read["НомерПорядковый"].ToString().Trim();
            }

            return num.Trim();

        }

        private void menuItem16_Click(object sender, EventArgs e)
        {
            FormОтчетИсполнитетлей form = new FormОтчетИсполнитетлей();
            form.YearSelect = выбранныйГод;
            form.Show();
        }

        /// <summary>
        /// Вставляем забытый документ.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem18_Click(object sender, EventArgs e)
        {
            string iTest = выбранныйГод;

            int seletYear = Convert.ToInt16(this.выбранныйГод) + 1;

            string filePatchLog = Application.StartupPath + @"\fileLog.txt";

            if (File.Exists(filePatchLog) == true)
            {
                File.Delete(filePatchLog);
                Log.WriteLine(filePatchLog, "Создадим лог");
            }
            else
            {
                Log.WriteLine(filePatchLog, "Создали лог файл");
            }


            // переменная для хранения прирощения дат.
            int inc = 0;
            StringBuilder builder = new StringBuilder();

            // Строка для хранения запроса к БД для получения id получателей.
            StringBuilder buildКор = new StringBuilder();

            FormКарточка form = new FormКарточка(ds11, seletYear.ToString(), true);

            form.ShowDialog(this);

            // Возможен касяк.-
            НомерДокумента docNumNext = form.СледующийНомерДокумента;

            if (form.DialogResult == DialogResult.OK)
            {
                DS1.КарточкаRow row = form.строкаКарточки;

                // Сгенерируем ГУИД для идентификации документа.
                Guid guidCard = Guid.NewGuid();

                inc = form.IncrementDate;

                string patchToServer = string.Empty;

                // Переменная для хранения имени файла на сервере.
                string namFileServer = string.Empty;

                // Получим выбранный способ поступления документа выбрал пользователь.
                ItemСпособПоступленияДокумента способПоступленияДокумента = form.СпособПоступления;

                // Пролучим список начальников управлений и отделов которым отписан документ.
                this.ListPerson = form.ListPerson;

                // Архивируем файл.

                //Если установлен флаг сохранения ксерокопии документа на сервере.
                if (form.SaveDocServer == true)
                {
                    if (form.ФлагЗаписиАрхива == true)
                    {
                        // Получим путь к файлу.
                        string filePatch = form.PathFileServer;

                        // Имя программы архиватора.
                        //string archiver = @"C:\Program Files\7-Zip\7z.exe";

                        // Получим имя папки которую нужно заархивировать.
                        string archive = form.FileName;// +@"\*.*";

                        // GUID составляющая названия файла.
                        string file = form.PathFileServer;

                        // Создадим своё имя файла архива содержащего архивируемую папку.
                        string namFileS = docNumNext.Номер.ToString() + "-" + docNumNext.Префикс + "_" + file;
                        string namFile = docNumNext.Номер.ToString() + "-" + docNumNext.Префикс;

                        // Путь к временному размещению папки с архивом.
                        string patch = Application.StartupPath + @"\Archive\" + namFile + ".7z";

                        fileName = patch;

                        namFileServer = namFile;// +".7z";

                        // Директоря куда будем архивировать файл.
                        string patchDir = Application.StartupPath + @"\Archive\";

                        // Архивируем папку. (Старая реализация)
                        //Archiver.AddToArchive(archiver, archive, patch,patchDir);

                        Log.WriteLine(filePatchLog, "Архивирование файла начало");

                        // Путь к 7z.dll.
                        string sevenZipDll = Application.StartupPath + @"\7z.dll";
                        if (archive.Length > 0)
                        {
                            // Пометим, что документ помечен для записи на сервер.
                            flagInsertCopyDoc = true;

                            Archiver.AddToArchive(sevenZipDll, archive, patch, patchDir);
                        }
                        else
                        {
                            // Пометим, что документ для записи на сервер не посмечен.
                            flagInsertCopyDoc = false;

                            MessageBox.Show("Вы не указали какую папку с документами записать на сервер", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }


                        Log.WriteLine(filePatchLog, "Архивирование файла конец");

                        // Путь куда будем архивировать папку.
                        patchToServer = patchServerFile + @"\" + namFileS.Trim();

                        fileNameCopy = patchToServer;
                    }
                    else
                    {
                        return;
                    }
                }

                // ===Begin========Запишем фамилии кому отписано документ в базу данных.
                // Разобъём строку на фамилии. (символ ,)
                string[] sКоррs = row["Резолюция"].ToString().Split(',');

                int id_карточки = Convert.ToInt32(row["id_карточки"]);

                // Полоучим время записи.
                DateTime todoy = DateTime.Now;

                // Счётчик циклов.
                int iCount = 1;

                //// Сформируем строку 
                //foreach (string str in sКоррs)
                //{
                //    string insert = "declare @id_" + iCount + "  int " +
                //                    "SELECT @id_" + iCount + " = id_получателя " +
                //                    "FROM [Получатели] " +
                //                    "where [ОписаниеПолучателя] = '" + str.Trim() + "' " +
                //                    "INSERT INTO [ПолучателДокументовУправление] " +
                //                               "([idПолучатель] " +
                //                               ",[ДатаВремяЗаписи] " +
                //                               ",[ОтметкаПрочтение] " +
                //                               ",[ОтметкаИсполнение] " +
                //                               ",[idКарточки] " +
                //                               ",[РезультатВыполнения]) " +
                //                         "VALUES " +
                //                               "(@id_" + iCount + " " +
                //                               ",'" + todoy + "' " +
                //                               ",NULL " +
                //                               ",NULL " +
                //                               ","+ id_карточки +" " +
                //                               ",NULL) ";

                //    // Добавим в запрос.
                //    builder.Append(insert);

                //    iCount++;
                //}
                //=============End=====================

                // Если пользователь не указал документ который нужно архивировать и отправлять на сервер
                // тогла установим флаг записи документа на запись без привязки номера документа.
                if (flagInsertCopyDoc == false)
                {
                    form.ФлагЗаписиАрхива = false;
                }

                // По умолчанию, добавляем новую карточку письмо НЕ ИНИЦИАТИВНОЕ.
                if (form.FlagRecordRepeet == false)
                {
                    string queryInsert = string.Empty;

                    // Строка для хранения хеш длинны файла.
                    string md5 = string.Empty;

                    if (form.ФлагЗаписиАрхива == true)
                    {
                        if (form.FlagAddDoc == true)
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          " declare @номерПП int  " +
                                          "select top 1 @номерПП = номерПП from Карточка " +
                                          "where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                //"select top 1 @номерПП = [номерПП] from Карточка " +
                                //"where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' " +  and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                                 "order by номерПП desc " +
                                                 "INSERT INTO Карточка " +
                                                 "([id_документа] " +
                                                 ",[id_корреспондента] " +
                                                 ",[ВДело] " +
                                                ",[ДатаИсхода] " +
                                                ",[ДатаПоступ] " +
                                                ",[КраткоеСодержание] " +
                                                ",[НаКонтроле] " +
                                                ",[НомерВход] " +
                                                ",[НомерИсход] " +
                                                ",[Резолюция] " +
                                                ",[РезультатВыполнения] " +
                                                ",[СрокВыполнения] " +
                                                ",[номерПП] " +
                                                ",[ОписаниеКорреспондента] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                ",MD5 " +
                                                ",idВидПоступленияДокумента  " +
                                                ", FlagAuto )" +
                                                "VALUES " +
                                                "( " + row["id_документа"] + " " +
                                                "," + row["id_корреспондента"] + " " +
                                                ",'" + row["ВДело"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["КраткоеСодержание"] + "' " +
                                                ",'" + row["НаКонтроле"] + "' " +
                                                //",'" + row["НомерВход"] + "' " +
                                                ",'"+ docNumNext.Префикс +"'" +
                                                ",'" + row["НомерИсход"] + "' " +
                                                ",'" + row["Резолюция"] + "' " +
                                                ",'" + row["РезультатВыполнения"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["номерПП"] + " " +
                                                "," + docNumNext.Номер + " " +
                                                //", @номерПП + 1 " +
                                                ",'" + row["ОписаниеКорреспондента"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",'" + namFileServer + "'  " +
                                                ",'" + form.PathFileServer + "' " +
                                                ",'md5' " +
                                                "," + способПоступленияДокумента.Id + "  " +
                                                ",'True' ) " +
                                                "SELECT @id_карточки = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                        else
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          " declare @номерПП int  " +
                                              "select top 1 @номерПП = номерПП from Карточка " +
                                          "where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                //"select top 1 @номерПП = [номерПП] from Карточка " +
                                //"where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' " +  and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                                 "order by номерПП desc " +
                                                 "INSERT INTO Карточка " +
                                                 "([id_документа] " +
                                                 ",[id_корреспондента] " +
                                                 ",[ВДело] " +
                                                ",[ДатаИсхода] " +
                                                ",[ДатаПоступ] " +
                                                ",[КраткоеСодержание] " +
                                                ",[НаКонтроле] " +
                                                ",[НомерВход] " +
                                                ",[НомерИсход] " +
                                                ",[Резолюция] " +
                                                ",[РезультатВыполнения] " +
                                                ",[СрокВыполнения] " +
                                                ",[номерПП] " +
                                                ",[ОписаниеКорреспондента] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                 ",MD5 " +
                                                ",idВидПоступленияДокумента  " +
                                                ", FlagAuto )" +
                                                "VALUES " +
                                                "( " + row["id_документа"] + " " +
                                                "," + row["id_корреспондента"] + " " +
                                                ",'" + row["ВДело"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["КраткоеСодержание"] + "' " +
                                                ",'" + row["НаКонтроле"] + "' " +
                                //",'" + row["НомерВход"] + "' " +
                                                ",'" + docNumNext.Префикс + "'" +
                                                ",'" + row["НомерИсход"] + "' " +
                                                ",'" + row["Резолюция"] + "' " +
                                                ",'" + row["РезультатВыполнения"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["номерПП"] + " " +
                                                "," + docNumNext.Номер + " " +
                                                //", @номерПП + 1 " +
                                                ",'" + row["ОписаниеКорреспондента"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",'" + namFileServer + "'  " +
                                                ",'" + form.PathFileServer + "' " +
                                                ",NULL " +
                                                 "," + способПоступленияДокумента.Id + "  " +
                                                ",'True' ) " +
                                                "SELECT @id_карточки = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                    }
                    else
                    {

                        if (form.FlagAddDoc == true)
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                        " declare @номерПП int  " +
                                          " select top 1 @номерПП = номерПП from Карточка " +
                                          " where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                          " id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                                          " where FlagAuto is null) " +
                                                 "order by номерПП desc " +
                                                "INSERT INTO Карточка " +
                                                 "([id_документа] " +
                                                 ",[id_корреспондента] " +
                                                 ",[ВДело] " +
                                                ",[ДатаИсхода] " +
                                                ",[ДатаПоступ] " +
                                                ",[КраткоеСодержание] " +
                                                ",[НаКонтроле] " +
                                //",'" + row["НомерВход"] + "' " +
                                                 ",[НомерВход] " +
                                                ",[НомерИсход] " +
                                                ",[Резолюция] " +
                                                ",[РезультатВыполнения] " +
                                                ",[СрокВыполнения] " +
                                                ",[номерПП] " +
                                                ",[ОписаниеКорреспондента] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                ",MD5 " +
                                                  ",idВидПоступленияДокумента  " +
                                                ", FlagAuto )" +
                                                "VALUES " +
                                                "( " + row["id_документа"] + " " +
                                                "," + row["id_корреспондента"] + " " +
                                                ",'" + row["ВДело"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["КраткоеСодержание"] + "' " +
                                                ",'" + row["НаКонтроле"] + "' " +
                                                //",'" + row["НомерВход"] + "' " +
                                                ",'" + docNumNext.Префикс + "'" +
                                                ",'" + row["НомерИсход"] + "' " +
                                                ",'" + row["Резолюция"] + "' " +
                                                ",'" + row["РезультатВыполнения"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["номерПП"] + " " +
                                                "," + docNumNext.Номер + " " +
                                                ", @номерПП + 1 " +
                                                ",'" + row["ОписаниеКорреспондента"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",NULL  " +
                                                ",'" + guidCard + "' " +
                                                ",'md5' " +
                                                "," + способПоступленияДокумента.Id + "  " +
                                                ",'True' ) " +
                                                "SELECT @id_карточки = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                        else
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          "declare @id_карточки int " +
                                          "declare @номерПП int " +
                                              "select top 1 @номерПП = номерПП from Карточка " +
                                          "where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                //"select top 1 @номерПП = [номерПП] from Карточка " +
                                //"where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' " +  and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                                 "order by номерПП desc " +
                                                "INSERT INTO Карточка " +
                                                 "([id_документа] " +
                                                 ",[id_корреспондента] " +
                                                 ",[ВДело] " +
                                                ",[ДатаИсхода] " +
                                                ",[ДатаПоступ] " +
                                                ",[КраткоеСодержание] " +
                                                ",[НаКонтроле] " +
                                //",'" + row["НомерВход"] + "' " +
                                                  ",[НомерВход] " +
                                                ",[НомерИсход] " +
                                                ",[Резолюция] " +
                                                ",[РезультатВыполнения] " +
                                                ",[СрокВыполнения] " +
                                                ",[номерПП] " +
                                                ",[ОписаниеКорреспондента] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                               ",MD5 " +
                                                 ",idВидПоступленияДокумента  " +
                                                ", FlagAuto )" +
                                                "VALUES " +
                                                "( " + row["id_документа"] + " " +
                                                "," + row["id_корреспондента"] + " " +
                                                ",'" + row["ВДело"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["КраткоеСодержание"] + "' " +
                                                ",'" + row["НаКонтроле"] + "' " +
                                                ",'" + docNumNext.Префикс + "'" +
                                                ",'" + row["НомерИсход"] + "' " +
                                                ",'" + row["Резолюция"] + "' " +
                                                ",'" + row["РезультатВыполнения"] + "' " +
                                                ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["номерПП"] + " " +
                                                "," + docNumNext.Номер + " " +
                                                //", @номерПП + 1 " +
                                                ",'" + row["ОписаниеКорреспондента"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",NULL  " +
                                                ",'" + guidCard + "' " +
                                                ",NULL " +
                                                 "," + способПоступленияДокумента.Id + "  " +
                                                ",'True' ) " +
                                                "SELECT @id_карточки = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                    }

                    // Сформируем строку связывающую номер каротчки с пользователем кому отписан документ.
                    foreach (string str in sКоррs)
                    {
                        string insert = "declare @id_" + iCount + "  int " +
                                        "SELECT @id_" + iCount + " = id_получателя " +
                                        "FROM [Получатели] " +
                                        "where [ОписаниеПолучателя] = '" + str.Trim() + "' " +
                                        "INSERT INTO [ПолучателДокументовУправление] " +
                                                   "([idПолучатель] " +
                                                   ",[ДатаВремяЗаписи] " +
                                                   ",[ОтметкаПрочтение] " +
                                                   ",[ОтметкаИсполнение] " +
                                                   ",[idКарточки] " +
                                                   ",[РезультатВыполнения]) " +
                                             "VALUES " +
                                                   "(@id_" + iCount + " " +
                            //",'" + ДатаSQL.Дата(todoy.ToShortDateString()) + "' " +
                                                   ",GETDATE() " +
                                                   ",NULL " +
                                                   ",NULL " +
                                                   ",@id_карточки " +
                                                   ",NULL) ";

                        // Добавим в запрос.
                        builder.Append(insert);

                        iCount++;
                    }

                    // Сформируем запись в связующую таблицу документа, вида получения документа и начальниками отделов и управлений которым отписан текущий документ.
                    foreach (PersonRecepient person in this.ListPerson)
                    {
                        string insert = "INSERT INTO [СвязующаяВидПоступленияДокПолучатели] " +
                                        "([id_person] " +
                                       ",[id_ВидПоступленияДок] " +
                                       ",[id_карточки]) " +
                                       "VALUES " +
                                       "(" + person.ID + " " +
                                       "," + способПоступленияДокумента.Id + " " +
                                       ",@id_карточки ) ";

                        // Добавим в запрос.
                        builder.Append(insert);
                    }


                    //builder.Append(queryInsert + "COMMIT TRANSACTION ");
                    builder.Append("COMMIT TRANSACTION ");

                    string sTest = builder.ToString().Trim();
                }

                // Добавляем новое письмо ИНИЦИАТИВНОЕ.
                if (form.FlagRecordRepeet == true)
                {
                    string queryInsert = string.Empty;
                    if (form.ФлагЗаписиАрхива == true)
                    {
                        queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                      "begin transaction  " +
                                   " declare @номерПП int  " +
                                          " select top 1 @номерПП = номерПП from Карточка " +
                                          " where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                          " id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                                          " where FlagAuto is null) " +
                                                 "order by номерПП desc " +
                                            " INSERT INTO Карточка " +
                                             "([id_документа] " +
                                             ",[id_корреспондента] " +
                                             ",[ВДело] " +
                                            ",[ДатаИсхода] " +
                                            ",[ДатаПоступ] " +
                                            ",[КраткоеСодержание] " +
                                            ",[НаКонтроле] " +
                            //",'" + row["НомерВход"] + "' " +
                                              ",[НомерВход] " +
                                            ",[НомерИсход] " +
                                            ",[Резолюция] " +
                                            ",[РезультатВыполнения] " +
                                            ",[СрокВыполнения] " +
                                            ",[номерПП] " +
                                            ",[ОписаниеКорреспондента] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                             ",[FlagCardRepeet] " +
                                            ",NameFileDocument  " +
                                              ",idВидПоступленияДокумента  " +
                                                ", FlagAuto )" +
                                            "VALUES " +
                                            "( " + row["id_документа"] + " " +
                                            "," + row["id_корреспондента"] + " " +
                                            ",'" + row["ВДело"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["КраткоеСодержание"] + "' " +
                                            ",'" + row["НаКонтроле"] + "' " +
                                           ",'" + docNumNext.Префикс + "'" +
                                            ",'" + row["НомерИсход"] + "' " +
                                            ",'" + row["Резолюция"] + "' " +
                                            ",'" + row["РезультатВыполнения"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["номерПП"] + " " +
                                            "," + docNumNext.Номер + " " +
                                            //", @номерПП + 1 " +
                                            ",'" + row["ОписаниеКорреспондента"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                             ",'" + namFileServer + "'  " +
                                            ",'" + form.PathFileServer + "'  " +
                                             "," + способПоступленияДокумента.Id + "  " +
                                                ",'True' ) " +
                                           "INSERT INTO КарточкаПовтор " +
                                             "([id_документа] " +
                                             ",[id_корреспондента] " +
                                             ",[ВДело] " +
                                            ",[ДатаИсхода] " +
                                            ",[ДатаПоступ] " +
                                            ",[КраткоеСодержание] " +
                                            ",[НаКонтроле] " +
                                            ",[НомерВход] " +
                                            ",[НомерИсход] " +
                                            ",[Резолюция] " +
                                            ",[РезультатВыполнения] " +
                                            ",[СрокВыполнения] " +
                                            ",[номерПП] " +
                                            ",[ОписаниеКорреспондента] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                            ",id_карточкиВходящей  " +
                                            ",ДатаПрирощение " +
                                            ",FlagControl)" +
                                            "VALUES " +
                                            "( " + row["id_документа"] + " " +
                                            "," + row["id_корреспондента"] + " " +
                                            ",'" + row["ВДело"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["КраткоеСодержание"] + "' " +
                                            ",'" + row["НаКонтроле"] + "' " +
                                            ",'" + row["НомерВход"] + "' " +
                                            ",'" + row["НомерИсход"] + "' " +
                                            ",'" + row["Резолюция"] + "' " +
                                            ",'" + row["РезультатВыполнения"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["номерПП"] + " " +
                                            "," + docNumNext.Номер + " " +
                                            ",'" + row["ОписаниеКорреспондента"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                            ",@@IDENTITY " +
                                            "," + inc + " " +
                                            ",'False') " +
                                            "COMMIT TRANSACTION ";
                    }
                    else
                    {
                        queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                      "begin transaction  " +
                                  " declare @номерПП int  " +
                                          " select top 1 @номерПП = номерПП from Карточка " +
                                          " where ДатаИсхода >= '" + seletYear.ToString().Trim() + "0101' and ДатаИсхода <= '" + seletYear.ToString().Trim() + "1231' " +
                                          " id_карточки in (SELECT MAX(id_карточки) FROM [Карточка] " +
                                          " where FlagAuto is null) " +
                                                 "order by номерПП desc " +
                                      "INSERT INTO Карточка " +
                                             "([id_документа] " +
                                             ",[id_корреспондента] " +
                                             ",[ВДело] " +
                                            ",[ДатаИсхода] " +
                                            ",[ДатаПоступ] " +
                                            ",[КраткоеСодержание] " +
                                            ",[НаКонтроле] " +
                            //",'" + row["НомерВход"] + "' " +
                                            ",[НомерВход] " +
                                            ",[НомерИсход] " +
                                            ",[Резолюция] " +
                                            ",[РезультатВыполнения] " +
                                            ",[СрокВыполнения] " +
                                            ",[номерПП] " +
                                            ",[ОписаниеКорреспондента] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                             ",[FlagCardRepeet] " +
                                            ",NameFileDocument  " +
                                               ",idВидПоступленияДокумента  " +
                                                ", FlagAuto )" +
                                            "VALUES " +
                                            "( " + row["id_документа"] + " " +
                                            "," + row["id_корреспондента"] + " " +
                                            ",'" + row["ВДело"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["КраткоеСодержание"] + "' " +
                                            ",'" + row["НаКонтроле"] + "' " +
                                          ",'" + docNumNext.Префикс + "'" +
                                            ",'" + row["НомерИсход"] + "' " +
                                            ",'" + row["Резолюция"] + "' " +
                                            ",'" + row["РезультатВыполнения"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["номерПП"] + " " +
                                            "," + docNumNext.Номер + " " +
                                            //", @номерПП + 1 " +
                                            ",'" + row["ОписаниеКорреспондента"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                             ",NULL  " +
                                             ",NULL  " +
                                               "," + способПоступленияДокумента.Id + "  " +
                                                ",'True' ) " +
                                           "INSERT INTO КарточкаПовтор " +
                                             "([id_документа] " +
                                             ",[id_корреспондента] " +
                                             ",[ВДело] " +
                                            ",[ДатаИсхода] " +
                                            ",[ДатаПоступ] " +
                                            ",[КраткоеСодержание] " +
                                            ",[НаКонтроле] " +
                                            ",[НомерВход] " +
                                            ",[НомерИсход] " +
                                            ",[Резолюция] " +
                                            ",[РезультатВыполнения] " +
                                            ",[СрокВыполнения] " +
                                            ",[номерПП] " +
                                            ",[ОписаниеКорреспондента] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                            ",id_карточкиВходящей  " +
                                            ",ДатаПрирощение " +
                                            ",FlagControl)" +
                                            "VALUES " +
                                            "( " + row["id_документа"] + " " +
                                            "," + row["id_корреспондента"] + " " +
                                            ",'" + row["ВДело"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаИсхода"]).ToShortDateString()) + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["КраткоеСодержание"] + "' " +
                                            ",'" + row["НаКонтроле"] + "' " +
                                            ",'" + row["НомерВход"] + "' " +
                                            ",'" + row["НомерИсход"] + "' " +
                                            ",'" + row["Резолюция"] + "' " +
                                            ",'" + row["РезультатВыполнения"] + "' " +
                                            ",'" + ДатаSQL.Дата(Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["номерПП"] + " " +
                                            "," + docNumNext.Номер + " " +
                                            ",'" + row["ОписаниеКорреспондента"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                            ",@@IDENTITY " +
                                            "," + inc + " " +
                                            ",'False') " +
                                            "COMMIT TRANSACTION ";
                    }

                    builder.Append(queryInsert);
                }

                string strBuild = builder.ToString();

                //// Сохраним данные.
                ПодключитьБД connectBD = new ПодключитьБД();
                string sCon = connectBD.СтрокаПодключения();

                // Флаг проверки успешной копии файла.
                bool flagCopyServer = false;

                // Выполним запрос на вставку (к сожалению не в единой транзакции.
                using (SqlConnection con = new SqlConnection(sCon))
                {

                    //Log.WriteLine(filePatchLog, "Начнём копровать файл на сервер");

                    //if (form.SaveDocServer == true)
                    //{
                    //    try
                    //    {
                    //        if (form.ФлагЗаписиАрхива == true)
                    //        {

                    //            Log.WriteLine(filePatchLog, "Копируем файл на сервер");

                    //            // Проверим помечен ли документ для записи на сервер.
                    //            if (flagInsertCopyDoc == true)
                    //            {
                    //                //Скопируем файл на сервер хранения документов.
                    //                //File.Copy(fileName, fileNameCopy, true);
                    //            }
                    //            Log.WriteLine(filePatchLog, "Закончим копировать файл на сервер");

                    //            //Если файл скопировался успешно постави флаг в true.
                    //            flagCopyServer = true;
                    //        }
                    //    }
                    //    catch(Exception exp)
                    //    {
                    //        Log.WriteLine(filePatchLog, "Ошибка при копировании - ");
                    //        Log.WriteLine(filePatchLog, exp.Message);
                    //        MessageBox.Show("Ошибка при копировании файла");

                    //        flagCopyServer = false;

                    //        return;
                    //    }

                    //    string fileTest = fileNameCopy;
                    //    if (File.Exists(fileNameCopy) == true)
                    //    {
                    //        Log.WriteLine(filePatchLog, "Выполним запись на сервер");

                    //        con.Open();
                    //        SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    //        com.ExecuteNonQuery();
                    //    }
                    //    else
                    //    {
                    //        con.Open();
                    //        SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    //        com.ExecuteNonQuery();
                    //    }
                    //}
                    //else
                    //{
                    // Если файл скопировался успешно постави флаг в true.
                    flagCopyServer = true;

                    con.Open();
                    SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    com.ExecuteNonQuery();
                    //}
                }


                //ds11.Карточка.AddКарточкаRow(row);
                ОбновитьДанные();

                // Пойдём по тупому варианту и получим номер документа.
                string queryNumDoc = "select id_карточки,номерПП,НомерВход from [Карточка] " +
                                     "where GuidName = '" + guidCard + "' ";

                string номерДок = string.Empty;

                DataTable tabNum;

                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();

                    SqlDataAdapter da = new SqlDataAdapter(queryNumDoc, con);

                    DataSet ds = new DataSet();

                    da.Fill(ds, "numDoc");

                    tabNum = ds.Tables["numDoc"];
                }

                номерДок = tabNum.Rows[0]["номерПП"].ToString().Trim() + "/" + tabNum.Rows[0]["НомерВход"].ToString().Trim();

                string номер = номерДок;

                // Получим номер id карточки.
                string idCard = tabNum.Rows[0]["id_карточки"].ToString().Trim();

                // Выводит номер зарегистрированного документа.
                FormMessage frmMessage = new FormMessage(номер);
                frmMessage.NumCardDoc = idCard.Trim();
                frmMessage.НомерДокумента = номерДок;
                frmMessage.СпособПоступленияДокумента = способПоступленияДокумента;
                frmMessage.TopMost = true;
                frmMessage.ShowDialog();

            }
        }


        /// <summary>
        /// Обработка карточки исходящей.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem19_Click(object sender, EventArgs e)
        {
            string iTest = выбранныйГод;

            int seletYear = Convert.ToInt16(this.выбранныйГод) + 1;

            FormКарточкаИсходящая form = new FormКарточкаИсходящая(ds11, выбранныйГод, true);
            //FormКарточкаИсходящая form = new FormКарточкаИсходящая(ds11, seletYear.ToString(), true);

            // Установим флаг в false.
            form.FlagОтветПисьмо = false;

            // Установим адресат.
            form.Адресат = "";

            DialogResult result = form.ShowDialog(this);
            if (result == DialogResult.OK)
            {

                // Получим выбранный способ поступления документа выбрал пользователь.
                ItemСпособПоступленияДокумента способПоступленияДокумента = form.СпособПоступления;

                DS1.КарточкаИсходящаяRow row = form.строкаИсходящейКарточки;

                НомерДокумента doc = new НомерДокумента();

                if (form.FlagNumStopDoc == false)
                {
                    doc.Номер = Convert.ToInt16(form.НомерDoc.Номер);
                }
                else
                {
                    // Если мы отключили автоматическое генерирование номера.
                    doc.Номер = form.NumDocNoAutomat;
                }


                // Обнулим переменную.
                numberPrefix = string.Empty;

                // Запишем префикс номера документа.
                numberPrefix = form.ПрефиксНомерИсходящий;

                //ds11.КарточкаИсходящая.AddКарточкаИсходящаяRow(row); 

                List<int> listIdВходДок = form.ListIDКарточки;
                //List<ОснованиеПередачи> listOP = form.ListОснованиеПередачи;

                // Если мы пишем инициативное письмо.
                if (form.FlagОтветПисьмо == false)
                {
                    // Строка для хранения SQL инструкции, для выполнения в одной транзакции.
                    StringBuilder buildInsert = new StringBuilder();

                    // Переменная для хранения номера
                    string numDirect = string.Empty;

                    if (form.FlagNumStopDoc == false)
                    {
                        string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                       "begin transaction  " +
                                       "declare @numDoc int " +
                                       "select top 1 @numDoc = НомерПорядковый from КарточкаИсходящая " +
                                       "where Дата >= '" + seletYear.ToString().Trim() + "0101' and Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                                       "order by id_карточки desc " +
                                       "declare @key int " +
                                       "INSERT INTO КарточкаИсходящая " +
                                       "([Дата] " +
                                       ",[НомерКомитета] " +
                                       ",[id_Подразделения] " +
                                       ",[НомерНоменклатурный] " +
                                       ",[НомерПорядковый] " +
                                       ",[id_Адресата] " +
                                       ",[Содержание] " +
                                       ",[id_ВходящегоДокумента] " +
                                       ",[ОписаниеКорреспондента] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ",idВидПоступленияДокумента ) " +
                                       "VALUES " +
                                       "('" + ДатаSQL.Дата(Convert.ToDateTime(row["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row["НомерКомитета"] + "' " +
                                       "," + row["id_Подразделения"] + " " +
                                       ",'" + row["НомерНоменклатурный"] + "' " +
                            //"," + row["НомерПорядковый"] + " " +
                            //", "+ doc.Номер + " " +
                                       ", @numDoc + 1  " +
                                       "," + row["id_Адресата"] + " " +
                                       ",'" + row["Содержание"] + "' " +
                            //","+ row["id_ВходящегоДокумента"]+" " +
                                       ",NULL " +
                            //",'"+ form.Адресат.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       // ",NULL " +
                                       //",NULL " +
                                       ", " + способПоступленияДокумента.Id + " ) " +
                                       "set @key = @@IDENTITY ";

                        buildInsert.Append(query);
                    }
                    else
                    {
                        string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                       "begin transaction  " +
                                       "declare @numDoc int " +
                                       "select top 1 @numDoc = НомерПорядковый from КарточкаИсходящая " +
                                       "where Дата >= '" + seletYear.ToString().Trim() + "0101' and Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                                       "order by id_карточки desc " +
                                       "declare @key int " +
                                       "INSERT INTO КарточкаИсходящая " +
                                       "([Дата] " +
                                       ",[НомерКомитета] " +
                                       ",[id_Подразделения] " +
                                       ",[НомерНоменклатурный] " +
                                       ",[НомерПорядковый] " +
                                       ",[id_Адресата] " +
                                       ",[Содержание] " +
                                       ",[id_ВходящегоДокумента] " +
                                       ",[ОписаниеКорреспондента] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ",idВидПоступленияДокумента " +
                                       ", FlagAutho ) " + 
                                       "VALUES " +
                                       "('" + ДатаSQL.Дата(Convert.ToDateTime(row["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row["НомерКомитета"] + "' " +
                                       "," + row["id_Подразделения"] + " " +
                                       ",'" + row["НомерНоменклатурный"] + "' " +
                            //"," + row["НомерПорядковый"] + " " +
                                       ", "+ doc.Номер + " " +
                                       //", @numDoc + 1  " +
                                       "," + row["id_Адресата"] + " " +
                                       ",'" + row["Содержание"] + "' " +
                            //","+ row["id_ВходящегоДокумента"]+" " +
                                       ",NULL " +
                            //",'"+ form.Адресат.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       // ",NULL " +
                                       //",NULL " +
                                       ", " + способПоступленияДокумента.Id + " " +
                                       ", 'True' ) " + 
                                       "set @key = @@IDENTITY ";

                        buildInsert.Append(query);
                    }

                    // Передадим в строку запроса на вставку SQL инструкцию на вставку в таблицу [СвязующаяЦельПолучениперсональныхДанных].

                    string sInsert = string.Empty;
                    sInsert = String.Format(form.QueryInsert.Trim(), "@key");

                    buildInsert.Append(sInsert.Trim());

                    // Завершим транзакцию.
                    buildInsert.Append("COMMIT TRANSACTION ");

                    // Обнулим список для хранения оснований передачи перед использованием.
                    //form.ListОснованиеПередачи.Clear();

                    ПодключитьБД connBD = new ПодключитьБД();
                    string sCon = connBD.СтрокаПодключения();

                    SqlConnection con = new SqlConnection(sCon);
                    con.Open();
                    SqlCommand com = new SqlCommand(buildInsert.ToString(), con);
                    com.ExecuteNonQuery();
                    con.Close();
                }
                else
                {
                    // Тестируем.
                    DS1.КарточкаИсходящаяRow row2 = form.строкаИсходящейКарточки;

                    int i = row2.id_ВходящегоДокумента;

                    // Строка для хранения запроса.
                    System.Text.StringBuilder builder = new System.Text.StringBuilder();

                    if (form.FlagNumStopDoc == false)
                    {
                        // Установим флаг в FALSE.
                        string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                        "begin transaction  " +
                                        "declare @numCard  int " +
                                        "select top 1 @numCard = НомерПорядковый  from КарточкаИсходящая " +
                                        "where Дата >= '" + seletYear.ToString().Trim() + "0101' and Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                                        "order by id_карточки desc " +
                                       "INSERT INTO КарточкаИсходящая " +
                                       "([Дата] " +
                                       ",[НомерКомитета] " +
                                       ",[id_Подразделения] " +
                                       ",[НомерНоменклатурный] " +
                                       ",[НомерПорядковый] " +
                                       ",[id_Адресата] " +
                                       ",[Содержание] " +
                                       ",[id_ВходящегоДокумента] " +
                                       ",[ОписаниеКорреспондента] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ", idВидПоступленияДокумента)" +
                                       "VALUES " +
                                       "('" + ДатаSQL.Дата(Convert.ToDateTime(row2["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row2["НомерКомитета"] + "' " +
                                       "," + row2["id_Подразделения"] + " " +
                                       ",'" + row2["НомерНоменклатурный"] + "' " +
                            //"," + row2["НомерПорядковый"] + " " +
                            //", " + doc.Номер + " " +
                                        ", @numCard + 1  " +
                                       "," + row2["id_Адресата"] + " " +
                                       ",'" + row2["Содержание"] + "' " +
                                       "," + row2["id_ВходящегоДокумента"] + " " +
                            //",NULL " +
                            //",'"+ form.Адресат.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row2["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       //",NULL " +
                                       //",NULL " +
                                       ", " + способПоступленияДокумента.Id + " )" +
                                       " declare @idCard int " +
                                       "select top 1 @idCard = id_карточки  from КарточкаИсходящая " +
                                       "order by id_карточки desc ";

                        builder.Append(query);
                    }
                    else
                    {
                        // Установим флаг в FALSE.
                        string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                        "begin transaction  " +
                                        "declare @numCard  int " +
                                        "select top 1 @numCard = НомерПорядковый  from КарточкаИсходящая " +
                                        "where Дата >= '" + seletYear.ToString().Trim() + "0101' and Дата <= '" + seletYear.ToString().Trim() + "1231' " +
                                        "order by id_карточки desc " +
                                       "INSERT INTO КарточкаИсходящая " +
                                       "([Дата] " +
                                       ",[НомерКомитета] " +
                                       ",[id_Подразделения] " +
                                       ",[НомерНоменклатурный] " +
                                       ",[НомерПорядковый] " +
                                       ",[id_Адресата] " +
                                       ",[Содержание] " +
                                       ",[id_ВходящегоДокумента] " +
                                       ",[ОписаниеКорреспондента] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ", idВидПоступленияДокумента" +
                                       ",FlagAutho )" +
                                       "VALUES " +
                                       "('" + ДатаSQL.Дата(Convert.ToDateTime(row2["Дата"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row2["НомерКомитета"] + "' " +
                                       "," + row2["id_Подразделения"] + " " +
                                       ",'" + row2["НомерНоменклатурный"] + "' " +
                            //"," + row2["НомерПорядковый"] + " " +
                                        ", " + doc.Номер + " " +
                                        //", @numCard + 1  " +
                                       "," + row2["id_Адресата"] + " " +
                                       ",'" + row2["Содержание"] + "' " +
                                       "," + row2["id_ВходящегоДокумента"] + " " +
                            //",NULL " +
                            //",'"+ form.Адресат.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row2["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       //",NULL " +
                                       //",NULL " +
                                       ", " + способПоступленияДокумента.Id + " " +
                                       ", 'True' ) " +
                                       " declare @idCard int " +
                                       "select top 1 @idCard = id_карточки  from КарточкаИсходящая " +
                                       "order by id_карточки desc ";
                        builder.Append(query);
                    }

                    string номерПодразделения = string.Empty;

                    DataRow[] rowsSelect = ds11.ПодразделенияКомитета.Select("id_подразделения= " + Convert.ToInt32(row2["id_Подразделения"]) + " ");
                    foreach (DataRow item in rowsSelect)
                    {
                        номерПодразделения = item["НомерПодразделения"].ToString().Trim();
                    }

                    string результатВыполнения = "Дан ответ. № исх. документа " + row2["НомерКомитета"].ToString().Trim() + "-" + row2["НомерНоменклатурный"].ToString().Trim() + "-" + номерПодразделения + "/" + doc.Номер.ToString().Trim();// row2["НомерПорядковый"].ToString().Trim();
                    //string результатВыполнения = "Дан ответ. № исх. документа " + row2["НомерКомитета"].ToString().Trim() + "-" + row2["НомерНоменклатурный"].ToString().Trim() + "-" + номерПодразделения + "/CAST(@numCard + 1 AS nvarchar) ";// +row2["НомерПорядковый"].ToString().Trim();

                    // Установим флаг в TRUE.
                    string queryUpdate = "UPDATE [Карточка] " +
                                         "SET РезультатВыполнения = '" + результатВыполнения + "' " + //' + CAST(@numCard + 1 AS nvarchar) " +
                        //"FlagPersonData = '" + row["FlagPersonData"] + "' " +
                                         ",ВДело = 'True' " +
                                         "where id_карточки = " + row["id_ВходящегоДокумента"] + " ";
                    // Соберём строки запросаов на добавление записи и на редактирование в единую строку, чтобы выполнить всё в одной транзакции.
                    builder.Append(queryUpdate);

                    string sTestNum = builder.ToString().Trim();

                    // Запрос на всатвку id в связующую таблицу ЦельПолученияПерсДанных.
                    foreach (ОснованиеПередачи itm in form.ListОснованиеПередачи)
                    {
                        string queryIns = "INSERT INTO [СвязующаяЦельПолучениперсональныхДанных] " +
                                       "([id_карточки] " +
                                       ",[id_ОснованиеПередачи]) " +
                                       "VALUES " +
                            //"('" + row.id_карточки + "' " +
                                       "( @idCard " +
                                       ",'" + itm.Id_основаниеПередачи + "' ) ";

                        builder.Append(queryIns);
                    }

                    // Заполним связующую таблицу Карточка входящаяИсходящая.
                    foreach (int idВх in listIdВходДок)
                    {

                        string queryIdВх = "INSERT INTO [СвязующаяКарточкаВходящаяИсходящая] " +
                                           "([id_карточкаВходящая] " +
                                           ",[id_карточкаИсходящая]) " +
                                           "VALUES " +
                                           "(" + idВх + " " +
                            //"," + row.id_карточки + " ) " +
                                            ",@idCard ) " +
                                           "update Карточка " +
                                           "set РезультатВыполнения = '" + form.НомерИсходящий.Trim() + "' " + " + CAST(@numCard + 1 AS nvarchar) " +
                                           "where id_карточки = " + idВх + " ";

                        builder.Append(queryIdВх);
                    }

                    // Проверим, что документ на который мы отвечаем стоит в стстусе повторных ответов.
                    СтатусКарточка card = new СтатусКарточка(Convert.ToInt32(row["id_ВходящегоДокумента"]));
                    bool flagStatusRepeet = card.СтатусПовторяющийсяОтвет();

                    // Если статус = true значит мы имеем дело с документом на который периодически необходимо довать ответ.
                    if (flagStatusRepeet == true)
                    {
                        // Строка для выполнения запроса в одной транзакции.
                        //StringBuilder querTransact = new StringBuilder();

                        /*
                         * Проверим отвечаем мы на этот документ впервые или нет.
                         * Для этого узнаем значение в поле ВДело в таблице Карточка, если установлено значение False 
                         * тогда на документ мы отвечаем впервые в противном случае нет.
                        */
                        ПодключитьБД bdConnect = new ПодключитьБД();
                        using (SqlConnection conn = new SqlConnection(bdConnect.СтрокаПодключения().Trim()))
                        {
                            conn.Open();
                            СтатусКарточка card2 = new СтатусКарточка(Convert.ToInt32(row["id_ВходящегоДокумента"]));
                            bool flagVD = card2.GetОтветПовторный(conn);

                            // Если на входящую карточку отвечают впервые.
                            if (flagVD == false)
                            {
    
                                string queryUp = " update Карточка " +
                                              "set ВДело = 'True' " +
                                              "where id_карточки = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " " +
                                              " declare @date datetime " +
                                              "declare @day int " +
                                              "declare @SetDate datetime " +
                                              "select @date = СрокВыполнения,@day = ДатаПрирощение from КарточкаПовтор " +
                                              "where id_карточкиВходящей = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " " +
                                              "SELECT @SetDate = DATEADD(day, @day, @date); " +
                                              "update КарточкаПовтор " +
                                              "set FlagControl = 'True' " +
                                              ",СрокВыполнения = @SetDate " +
                                              "where id_карточкиВходящей = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " ";

                                builder.Append(queryUp);
                            }

                            // Если ответ повторный.
                            if (flagVD == true)
                            {
                                // Увеличим значение в поле СрокИсполнения в таблице КарточкаПовтор на количесвто дней указанных в поле ДатаПрирощения.
                                string queryUpdatDate = " declare @date datetime " +
                                                        "declare @day int " +
                                                        "declare @SetDate datetime " +
                                                        "select @date = СрокВыполнения,@day = ДатаПрирощение from КарточкаПовтор " +
                                                        "where id_карточкиВходящей = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " " +
                                                        "SELECT @SetDate = DATEADD(day, @day, @date); " +
                                                        "update КарточкаПовтор " +
                                                        "set СрокВыполнения = @SetDate " +
                                                        "where id_карточкиВходящей = " + Convert.ToInt32(row["id_ВходящегоДокумента"]) + " ";

                                //querTransact.Append(queryUpdatDate);
                                builder.Append(queryUpdatDate);
                            }
                        }
                    }

                    // Завершим транзакцию.
                    builder.Append("COMMIT TRANSACTION ");

                    string queryTest = builder.ToString().Trim();

                    // Выполним запрос.
                    ПодключитьБД strConnectBD = new ПодключитьБД();
                    string strConn = strConnectBD.СтрокаПодключения();

                    // Откроем соединение и выполним запрос.
                    SqlConnection con = new SqlConnection(strConn);
                    con.Open();
                    SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    com.ExecuteNonQuery();

                    // Закроме соединение.
                    con.Close();

                }

                ОбновитьДанные();

                string iTest2 = "test";

                // Получим номер документа.
                NumOutputCardVipNet numDoc = GetNumDocOutVipNet(form.StrGuid);

                //string numberDocument = form.ПрефиксНомерИсходящий.Trim() + "/" + numDoc.Trim();

                string numberDocument = numberPrefix + "/" + numDoc.НомерПорядковый.Trim();

                // Выведим сообщение с новым номером.
                FormMessage message = new FormMessage(numberDocument.Trim());
                message.TopMost = true;
                message.СпособПоступленияДокумента = способПоступленияДокумента;
                message.NumCardDoc = numDoc.Id.ToString().Trim();
                message.НомерДокумента = numberDocument;
                message.ShowDialog();

            }
        }

        private void menuItem24_Click(object sender, EventArgs e)
        {

            ПечатьКонтрольныхУведомлений();
            //потокОжидания = new System.Threading.Thread(new System.Threading.ThreadStart(ЗапуститьФормуОжидания));
            //потокОжидания.Start();

            //ПечатьПросроченныхДокументов();
        }

        private void menuItem25_Click(object sender, EventArgs e)
        {
            //потокОжидания = new System.Threading.Thread(new System.Threading.ThreadStart(ЗапуститьФормуОжидания));
            //потокОжидания.Start();

            //ПечатьПросроченныхДокументов();

            //string querySelect = "SELECT * FROM [Выборка] " +
            //                    "where СрокВыполнения<'" + ДатаSQL.Дата(DateTime.Today.ToShortDateString()) + "' AND ДатаПоступ >= '" + выбранныйГод + "0112' AND НаКонтроле='True' AND ВДело='False'";

            //string query = "SELECT     convert(VARCHAR,dbo.Карточка.номерПП) + '/' + dbo.Карточка.НомерВход as 'НомерВход',  dbo.Карточка.ДатаПоступ, dbo.Корреспонденты.ОписаниеКорреспондента, " +
            //               " dbo.Карточка.КраткоеСодержание, dbo.Карточка.ДатаИсхода, dbo.Карточка.НомерИсход, dbo.Карточка.СрокВыполнения, " +
            //               " dbo.Карточка.Резолюция " +
            //               "FROM         dbo.Карточка INNER JOIN " +
            //               "  dbo.Корреспонденты ON dbo.Карточка.id_корреспондента = dbo.Корреспонденты.id_корреспондента " +
            //               " where ВДело = 'False' and СрокВыполнения < CONVERT(DATE,GETDATE()) ";

            string query = "select НомерВход,ДатаПоступ,ОписаниеКорреспондента,КраткоеСодержание, ДатаИсхода, НомерИсход, СрокВыполнения,ОписаниеПолучателя as Резолюция from dbo.ViewВыборкаОтчет " +
                           "where ВДело = 'False' and СрокВыполнения < CONVERT(DATE,GETDATE()) ";

            GetDataTable getTable = new GetDataTable(query);
            DataTable tab = getTable.DataTable("Выборка");

            FormПросроченныеДок form = new FormПросроченныеДок();
            form.TopMost = false;
            form.TabDate = tab;
            //form.ListDoc = list;
            form.Show();

        }

        private void menuItem26_Click(object sender, EventArgs e)
        {
            FormSelectDate formSelD = new FormSelectDate();
            formSelD.ShowDialog();

            if (formSelD.DialogResult == DialogResult.OK)
            {
                RangeDate rd = formSelD.ДиапазоДат;

                FormStatInputKorr form = new FormStatInputKorr();
                form.TopMost = true;
                form.ДиапазонДат = rd;
                form.Show();
            }

        }

        private void menuItem27_Click(object sender, EventArgs e)
        {
            FormSelectDate formSelD = new FormSelectDate();
            formSelD.ShowDialog();

            if (formSelD.DialogResult == DialogResult.OK)
            {
                RangeDate rd = formSelD.ДиапазоДат;

                FormViewInputDoc form = new FormViewInputDoc();
                //form.TopMost = true;
                form.ДиапазонДат = rd;
                form.Show();
                
            }
        }

        private void menuItem28_Click(object sender, EventArgs e)
        {
            FormПечатьКарточки form = new FormПечатьКарточки();
            form.ТекущийГод = Дата.ПервыйДень(selectedYear.ToString());
            form.БудущийГод = следующаяДата;
            form.ShowDialog(this);
            this.Enabled = true;
        }

        private void menuItem29_Click(object sender, EventArgs e)
        {
            FormSelectDate formSelD = new FormSelectDate();
            formSelD.ShowDialog();

            if (formSelD.DialogResult == DialogResult.OK)
            {
                RangeDate rd = formSelD.ДиапазоДат;

                FormStatOutpotCorr form = new FormStatOutpotCorr();
                form.TopMost = true;
                form.ДиапазонДат = rd;
                form.Show();
            }
        }

        private void menuItem30_Click(object sender, EventArgs e)
        {
            FormSelectDate formSelD = new FormSelectDate();
            formSelD.ShowDialog();

            if (formSelD.DialogResult == DialogResult.OK)
            {
                RangeDate rd = formSelD.ДиапазоДат;

                FormViewOutputDoc form = new FormViewOutputDoc();
                form.TopMost = true;
                form.ДиапазонДат = rd;
                form.Show();

            }
        }

        private void menuItem31_Click(object sender, EventArgs e)
        {
            FormПодразделениеРуководитель form = new FormПодразделениеРуководитель(выбраннаяДата);
            form.Show();
        }

        private void menuItem20_Click(object sender, EventArgs e)
        {

        }

        private void menuItem32_Click(object sender, EventArgs e)
        {
            FormSelectDatePerson personDate = new FormSelectDatePerson();
            //personDate.MdiParent = this;
            personDate.ShowDialog();
        }

        private void menuItem21_Click(object sender, EventArgs e)
        {

        }

        private void menuItem33_Click(object sender, EventArgs e)
        {
            string query = "SELECT     convert(VARCHAR,dbo.Карточка.номерПП) + '/' + dbo.Карточка.НомерВход as 'НомерВход',  dbo.Карточка.ДатаПоступ, dbo.Корреспонденты.ОписаниеКорреспондента, " +
                         " dbo.Карточка.КраткоеСодержание, dbo.Карточка.ДатаИсхода, dbo.Карточка.НомерИсход, dbo.Карточка.СрокВыполнения, " +
                         " dbo.Карточка.Резолюция " +
                         "FROM         dbo.Карточка INNER JOIN " +
                         "  dbo.Корреспонденты ON dbo.Карточка.id_корреспондента = dbo.Корреспонденты.id_корреспондента " +
                         " where ВДело = 'False' and СрокВыполнения < CONVERT(DATE,GETDATE()) ";

            GetDataTable getTable = new GetDataTable(query);
            DataTable tab = getTable.DataTable("Выборка");

            FormПросроченныеДок form = new FormПросроченныеДок();
            form.TopMost = false;
            form.TabDate = tab;
            //form.ListDoc = list;
            form.Show();

        }

        private void menuItem23_Click(object sender, EventArgs e)
        {

        }

        private void menuItem22_Click(object sender, EventArgs e)
        {

        }

        private void menuItem33_Click_1(object sender, EventArgs e)
        {
            // Создадим отчет.
            FormSelectInputDatePerson forminput = new FormSelectInputDatePerson();
            forminput.Show();
        }

        

       

        

    }
}

