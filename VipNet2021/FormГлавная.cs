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
    /// Summary description for Form�������.
    /// </summary>
    public class Form������� : System.Windows.Forms.Form
    {
        #region ����������

        private System.Windows.Forms.DataGridTableStyle dataGridTableStyle2;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn��������;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn�������������;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn����������;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn����������;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn����������;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn���������;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn����������;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn��������;
        private System.Windows.Forms.DataGridBoolColumn dataGridBoolColumn�����;
        private System.Windows.Forms.DataGrid dataGrid����������������;
        private System.Windows.Forms.DataGrid dataGrid��������������;
        private System.Windows.Forms.DataGrid dataGrid������������������;
        private System.Windows.Forms.DataGridTableStyle dataGridTableStyle����������������;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn1;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn2;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn3;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn4;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn5;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn6;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn7;
        private System.Windows.Forms.DataGridTextBoxColumn dataGridTextBoxColumn8;
        private System.Windows.Forms.DataGridTableStyle dataGridTableStyle��������������;
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
        private System.Windows.Forms.RichTextBox label����Tab1;
        private System.Windows.Forms.RichTextBox label����Tab2;
        private System.Windows.Forms.TextBox textBox������������Tab2;
        private System.Windows.Forms.TextBox textBox������������Tab1;
        private System.Windows.Forms.Button button��������������������Tab1;
        private System.Windows.Forms.Button button��������������������Tab2;
        private System.Windows.Forms.Label label�������������������������Tab2;
        private System.Windows.Forms.Label label�������������������������Tab1;
        private System.Windows.Forms.MainMenu mainMenu1;
        private System.Windows.Forms.MenuItem menuItem1;
        private System.Windows.Forms.MenuItem menuItem2;
        private System.Windows.Forms.MenuItem menuItem3;
        private System.Windows.Forms.MenuItem menuItem4;
        private System.Windows.Forms.MenuItem menuItem��������������;
        private System.Windows.Forms.MenuItem menuItemContext��������������;
        private System.Windows.Forms.MenuItem menuItem��������������������;
        private System.Windows.Forms.MenuItem menuItem�������������������������;
        private System.Windows.Forms.MenuItem menuItem���������������������;
        private System.Windows.Forms.MenuItem menuItem����������������������;
        private System.Windows.Forms.MenuItem menuItem����������������;
        private System.Windows.Forms.ContextMenu contextMenu1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.CheckBox checkBoxKontrolFilter;
        private System.ComponentModel.IContainer components;
        System.Windows.Forms.TreeNode treeNode������;
        System.Windows.Forms.TreeNode treeNode�������;
        System.Windows.Forms.TreeNode treeNode����;
        System.Windows.Forms.TreeNode treeNode������;
        System.Windows.Forms.TreeNode treeNode���;
        System.Windows.Forms.TreeNode treeNode����;
        System.Windows.Forms.TreeNode treeNode����;
        System.Windows.Forms.TreeNode treeNode������;
        System.Windows.Forms.TreeNode treeNode��������;
        System.Windows.Forms.TreeNode treeNode�������;
        System.Windows.Forms.TreeNode treeNode������;
        System.Windows.Forms.TreeNode treeNode�������;
        System.Windows.Forms.TreeNode treeNode���;
        /// <summary>
        /// ���-������� �� ��������� �����������
        /// </summary>
        private System.Windows.Forms.TabControl tabControl�����������������;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;

        /// <summary>
        /// ���-������� � ������ ����������
        /// </summary>
        private System.Windows.Forms.TabControl tabControl��������������;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.TabPage tabPage3;

        /// <summary>
        /// ������������� ��� "������� ����������"
        /// </summary>
        private System.Data.DataView dataView�������������������;

        /// <summary>
        /// ������������� ��� "���������� � ����"
        /// </summary>
        private System.Data.DataView dataView���������������������;

        /// <summary>
        /// ������������� ��� ��������� ����������
        /// </summary>
        private System.Data.DataView dataView������������������;

        /// <summary>
        /// ��������� ��������� ������������ �� ����� App.config
        /// </summary>
        System.Configuration.AppSettingsReader configReader;

        /// <summary>
        /// ������ ����������� � ��������� ������
        /// </summary>
        string ����������������� = "";
        
        /// <summary>
        /// ������ ���������� ��������� ����������� "-"
        /// </summary>
        string[] TimeInterval;
        
        /// <summary>
        /// ����������� � ��������� ������
        /// </summary>
        SqlConnection �����������;

        /// <summary>
        /// ������������ �������� ����������, �������, �������� � ������� ��� ���������� ������
        /// </summary>
        SqlDataAdapter �����������;

        /// <summary>
        /// ������ � ������ ��������� ������
        /// </summary>
        string ��������������;

        /// <summary>
        /// ���������� ������ ��� �������� �� pegecr ������ ����� ���������
        /// </summary>
        static System.Threading.Mutex mutex;

        /// <summary>
        /// ��������� ���� � �������� ��������
        /// </summary>
        public System.Threading.Thread �������������;

        private DataGridTableStyle dataGridTableStyle������������������;
        private DataGridTextBoxColumn dataGridTextBoxColumn����������������;
        private DataGridTextBoxColumn dataGridTextBoxColumn�����������;
        private DataGridTextBoxColumn dataGridTextBoxColumn����������������������;
        private DataGridTextBoxColumn dataGridTextBoxColumn����������������;
        private DataGridTextBoxColumn dataGridTextBoxColumn��������������������;
        private SplitContainer splitContainer1;
        private SplitContainer splitContainer2;
        private RichTextBox label����Tab3;
        private TextBox textBox�������������������������������;
        private Label label��������������������������������������������;
        private Button button���������������������������������������;
        private MenuItem menuItem������������������������;
        private DS1 ds11;
        private SplitContainer splitContainer3;
        private TableLayoutPanel tableLayoutPanel2;
        private CheckBox checkBox1;
        private CheckBox checkBox2;
        private Label label2;
        private MenuItem menuItem5;
        private ComboBox comboBox��������������;
        private CheckBox checkBox��������������;
        private ComboBox comboBox��������������;
        private MenuItem menuItem6;
        private MenuItem menuItem7;
        private MenuItem menuItem8;
        private MenuItem menuItem9;

        //������ ��������� ���
        private int selectedYear;
        private string �������������;
        
        private string ������������;
        private string �������������;
        private MenuItem menuItem10;
        private MenuItem menuItem11;
        private MenuItem menuItem12;
        private MenuItem menuItem13;

        // ���������� ��� �������� ����� ����� ������� ���������� �� ������.
        private string fileName = string.Empty;
        private string fileNameCopy = string.Empty;

        // ���� ���������, ��� ������ � ������������� ����������� �������������� 1-�� ���.
        private bool flagFirstLoad = false;

        // ���������� ��� �������� �������� �������� ���������.
        private string numberPrefix = string.Empty;

        /// <summary>
        /// ������������ ���������� ����
        /// </summary>
        private struct ��������������
        {
            public int X;
            public int Y;
        }

        /// <summary>
        /// ��������� � ������������ ����
        /// </summary>
        �������������� ���� = new ��������������();
        private MenuItem menuItem14;
        private MenuItem menuItem15;

        /// <summary>
        /// ���� � ��������� ����� ����������� �� ������.
        /// </summary>
        private string patchServerFile = string.Empty;

        /// <summary>
        /// ���� ���������, ��� �������� ������� ��� ������ �� ������.
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
        /// �������� ������ ������ ����������� ������� � ���������� ������� ������� ��� ��������� ����������� ���������.
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
        /// ����������� �����
        /// </summary>
        public Form�������()
        {
            InitializeComponent();

            // ��������� ���� � ��������� ������:
            configReader = new AppSettingsReader();
            ����������������� = (string)configReader.GetValue("�����������������", typeof(System.String));
            //��������� ������ ������� �� ���� ������ 
            TimeInterval = configReader.GetValue("�����������������", typeof(System.String)).ToString().Split('-');

            //������� ���
            SelectYearForm selectYearForm = new SelectYearForm();
            selectYearForm.ShowDialog();

            //���� ������������ ����� �� ���������� ��������� ���
            if (selectYearForm.DialogResult == DialogResult.OK)
            {
                //������� ��������� ���
                selectedYear = selectYearForm.SelectedYear;
                ������������� = �������������������.����������������������(selectedYear);

                //������� ��� � �������� ��� ���������� ������ �������� ���������
                ������������ = �������������������.������������(selectedYear); //selectedYear.ToString();

                //������� ��������� ����
                ������������� = �������������������.����������������(selectedYear);
            }

            //���� ������������ ����� ������ ������� �� ����������
            if (selectYearForm.DialogResult == DialogResult.Cancel)
            {
                this.Close();
                Environment.Exit(0);
            }
            

            // ������� ����������� � ��������� ������:
            ����������� = new SqlConnection(�����������������);

            // ������� ����������� ��� �������� ��� ���������� ������:
            ����������� = new SqlDataAdapter("", �����������);

            �������������� = �����������.DataSource.ToString();

            string str = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Environment.CurrentDirectory + "\\RegKor.exe").FileVersion;

            this.Text = "����������� ���������������. ������: " + str + ". SQL Server: " + �����������.DataSource;

            // ��������� ������� ������� � ���������� �� �� �����:
            ��������������������������();

            // ���������, ���� �� ������������ ���������
            //DataRow[] rows = ds11.�������.Select("��������������<'" + DateTime.Now.ToString() + "' AND ���������� >='01.12.2011' AND ����������=True AND �����=False");

            List<���������������������> list = new List<���������������������>();

            string sTest = DateTime.Now.ToString();

            string querySelect = "SELECT * FROM [�������] " +
                                 "where ��������������<'"+ ����SQL.����(DateTime.Today.ToShortDateString()) +"' AND ���������� >= '"+ ������������ +"0112' AND ����������='True' AND �����='False'";

            GetDataTable getTable = new GetDataTable(querySelect);
            DataTable tab = getTable.DataTable("�������");

            int iCount = 1;

            // �������� ������� ������.
            foreach (DataRow row in tab.Rows)
            {
                ��������������������� item = new ���������������������();
                item.������� = iCount.ToString().Trim();
                item.������������������������ = row["���������"].ToString().Trim();
                item.��������������� = Convert.ToDateTime(row["����������"]).ToShortDateString();
                item.������������� = row["���������"].ToString().Trim();
                item.�������������� = Convert.ToDateTime(row["��������������"]).ToShortDateString();

                list.Add(item);

                iCount++;
            }



            // ������� ������ �� ������������� �������������.
            string querySelectP = "SELECT * FROM [�������������] " +
                                "where �������������� <'" + ����SQL.����(DateTime.Today.ToShortDateString()) + "' AND ���������� >= '" + ������������ + "0112' AND ����������='True' AND �����='False'";

            // ������� ������ �� ������������� �������������.
            GetDataTable getTableP = new GetDataTable(querySelectP);
            DataTable tabP = getTableP.DataTable("�������������");

            // �������� ������� ������.
            foreach (DataRow row in tabP.Rows)
            {
                ��������������������� item = new ���������������������();
                item.������� = iCount.ToString().Trim();
                item.������������������������ = row["���������"].ToString().Trim();
                item.��������������� = Convert.ToDateTime(row["����������"]).ToShortDateString();
                item.������������� = row["���������"].ToString().Trim();
                //item.�������������� = Convert.ToDateTime(row["��������������"]).ToShortDateString();
                item.�������������� = Convert.ToDateTime(row["��������������"]).AddDays(2).ToShortDateString();

                list.Add(item);

                iCount++;
            }

            int ii = list.Count;


            //if (tab.Rows.Count > 0)
            if (list.Count > 0)
            {
                if (flagFirstLoad == false)
                {
                    //����������������������������(list); -- ������� ���� ����� �����������.
                    flagFirstLoad = true;
                }
            }

            // ������� ������� ���.
            //DataRow[] rows = ds11.�������.Select("��������������<'" + DateTime.Now.ToShortDateString() + "' AND ���������� >= '01.12." + ������������ + "' AND ����������=True AND �����=False");
            //if (rows.Length > 0)
            //{
            //    ����������������������������();
            //}

            //������ ����� ����
            comboBox��������������.SelectedItem = "���� ���";
            string ������2 = "�����=False AND ���������� >='01.12." + ������������ + "'";
            dataView�������������������.RowFilter = ������2;
            dataGrid����������������.DataSource = dataView�������������������;

            string ������3 = "�����=True AND ���������� >='01.12." + ������������ + "'";
            dataView���������������������.RowFilter = ������3;
            dataGrid��������������.DataSource = dataView���������������������;

            // ������� ���� ��� ����������� �����.
            string queryPatchServer = "select top 1 PatchServer from ����������";

            // ������� ������ �� � ���� ����������� �����.
            GetDataTable getTablePatch = new GetDataTable(queryPatchServer);
            DataTable tabPatch = getTablePatch.DataTable("���������");

            // ������� � ���������� ����� ���� ������������ ������.
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form�������));
            this.dataView������������������� = new System.Data.DataView();
            this.contextMenu1 = new System.Windows.Forms.ContextMenu();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label�������������������������Tab1 = new System.Windows.Forms.Label();
            this.button��������������������Tab1 = new System.Windows.Forms.Button();
            this.checkBoxKontrolFilter = new System.Windows.Forms.CheckBox();
            this.panel3 = new System.Windows.Forms.Panel();
            this.textBox������������Tab1 = new System.Windows.Forms.TextBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.textBox������������Tab2 = new System.Windows.Forms.TextBox();
            this.button��������������������Tab2 = new System.Windows.Forms.Button();
            this.textBox������������������������������� = new System.Windows.Forms.TextBox();
            this.button��������������������������������������� = new System.Windows.Forms.Button();
            this.panel4Tab1 = new System.Windows.Forms.Panel();
            this.dataGrid���������������� = new System.Windows.Forms.DataGrid();
            this.dataGridTableStyle���������������� = new System.Windows.Forms.DataGridTableStyle();
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
            this.menuItem�������������� = new System.Windows.Forms.MenuItem();
            this.menuItem2 = new System.Windows.Forms.MenuItem();
            this.menuItem������������������������� = new System.Windows.Forms.MenuItem();
            this.menuItem������������������������ = new System.Windows.Forms.MenuItem();
            this.menuItem��������������������� = new System.Windows.Forms.MenuItem();
            this.menuItem�������������������� = new System.Windows.Forms.MenuItem();
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
            this.menuItem���������������� = new System.Windows.Forms.MenuItem();
            this.menuItem���������������������� = new System.Windows.Forms.MenuItem();
            this.menuItemContext�������������� = new System.Windows.Forms.MenuItem();
            this.menuItem5 = new System.Windows.Forms.MenuItem();
            this.menuItem4 = new System.Windows.Forms.MenuItem();
            this.menuItem7 = new System.Windows.Forms.MenuItem();
            this.menuItem6 = new System.Windows.Forms.MenuItem();
            this.menuItem10 = new System.Windows.Forms.MenuItem();
            this.menuItem13 = new System.Windows.Forms.MenuItem();
            this.menuItem14 = new System.Windows.Forms.MenuItem();
            this.menuItem15 = new System.Windows.Forms.MenuItem();
            this.menuItem16 = new System.Windows.Forms.MenuItem();
            this.tabControl����������������� = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.panel1Tab1 = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label����Tab1 = new System.Windows.Forms.RichTextBox();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.panel1Tab2 = new System.Windows.Forms.Panel();
            this.dataGrid�������������� = new System.Windows.Forms.DataGrid();
            this.dataGridTableStyle�������������� = new System.Windows.Forms.DataGridTableStyle();
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
            this.label����Tab2 = new System.Windows.Forms.RichTextBox();
            this.panel4 = new System.Windows.Forms.Panel();
            this.checkBox�������������� = new System.Windows.Forms.CheckBox();
            this.comboBox�������������� = new System.Windows.Forms.ComboBox();
            this.label�������������������������Tab2 = new System.Windows.Forms.Label();
            this.panel6 = new System.Windows.Forms.Panel();
            this.dataGridTableStyle2 = new System.Windows.Forms.DataGridTableStyle();
            this.dataGridTextBoxColumn�������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn������������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn���������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn���������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn���������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn��������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn���������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn�������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridBoolColumn����� = new System.Windows.Forms.DataGridBoolColumn();
            this.dataView��������������������� = new System.Data.DataView();
            this.tabControl�������������� = new System.Windows.Forms.TabControl();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.splitContainer3 = new System.Windows.Forms.SplitContainer();
            this.dataGrid������������������ = new System.Windows.Forms.DataGrid();
            this.dataGridTableStyle������������������ = new System.Windows.Forms.DataGridTableStyle();
            this.dataGridTextBoxColumn���������������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn����������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn���������������������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn���������������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.dataGridTextBoxColumn�������������������� = new System.Windows.Forms.DataGridTextBoxColumn();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.comboBox�������������� = new System.Windows.Forms.ComboBox();
            this.label�������������������������������������������� = new System.Windows.Forms.Label();
            this.label����Tab3 = new System.Windows.Forms.RichTextBox();
            this.dataView������������������ = new System.Data.DataView();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.label2 = new System.Windows.Forms.Label();
            this.ds11 = new RegKor.DS1();
            this.menuItem33 = new System.Windows.Forms.MenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.dataView�������������������)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4Tab1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid����������������)).BeginInit();
            this.tabControl�����������������.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel1Tab1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.panel1Tab2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid��������������)).BeginInit();
            this.panel5.SuspendLayout();
            this.panel7.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panel6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataView���������������������)).BeginInit();
            this.tabControl��������������.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.splitContainer3.Panel1.SuspendLayout();
            this.splitContainer3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid������������������)).BeginInit();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataView������������������)).BeginInit();
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
            this.panel2.Controls.Add(this.label�������������������������Tab1);
            this.panel2.Controls.Add(this.button��������������������Tab1);
            this.panel2.Controls.Add(this.checkBoxKontrolFilter);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(280, 110);
            this.panel2.TabIndex = 0;
            // 
            // label�������������������������Tab1
            // 
            this.label�������������������������Tab1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.label�������������������������Tab1.Location = new System.Drawing.Point(0, 86);
            this.label�������������������������Tab1.Name = "label�������������������������Tab1";
            this.label�������������������������Tab1.Size = new System.Drawing.Size(276, 20);
            this.label�������������������������Tab1.TabIndex = 5;
            // 
            // button��������������������Tab1
            // 
            this.button��������������������Tab1.Location = new System.Drawing.Point(208, 20);
            this.button��������������������Tab1.Name = "button��������������������Tab1";
            this.button��������������������Tab1.Size = new System.Drawing.Size(64, 22);
            this.button��������������������Tab1.TabIndex = 4;
            this.button��������������������Tab1.Text = "��������";
            this.toolTip1.SetToolTip(this.button��������������������Tab1, "�������� ������� ������.");
            this.button��������������������Tab1.Click += new System.EventHandler(this.button��������������������Tab1_Click);
            // 
            // checkBoxKontrolFilter
            // 
            this.checkBoxKontrolFilter.Location = new System.Drawing.Point(8, 28);
            this.checkBoxKontrolFilter.Name = "checkBoxKontrolFilter";
            this.checkBoxKontrolFilter.Size = new System.Drawing.Size(160, 27);
            this.checkBoxKontrolFilter.TabIndex = 3;
            this.checkBoxKontrolFilter.Text = "������ �� ��������";
            this.toolTip1.SetToolTip(this.checkBoxKontrolFilter, "������ - ������ �������������� ���������.");
            this.checkBoxKontrolFilter.CheckedChanged += new System.EventHandler(this.checkBoxKontrolFilter_CheckedChanged);
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.textBox������������Tab1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel3.Location = new System.Drawing.Point(0, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(276, 20);
            this.panel3.TabIndex = 2;
            // 
            // textBox������������Tab1
            // 
            this.textBox������������Tab1.Dock = System.Windows.Forms.DockStyle.Top;
            this.textBox������������Tab1.Location = new System.Drawing.Point(0, 0);
            this.textBox������������Tab1.Name = "textBox������������Tab1";
            this.textBox������������Tab1.Size = new System.Drawing.Size(276, 21);
            this.textBox������������Tab1.TabIndex = 0;
            this.toolTip1.SetToolTip(this.textBox������������Tab1, "������� ����� ��� ������.");
            this.textBox������������Tab1.TextChanged += new System.EventHandler(this.textBox������������_TextChanged);
            // 
            // textBox������������Tab2
            // 
            this.textBox������������Tab2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox������������Tab2.Location = new System.Drawing.Point(0, 0);
            this.textBox������������Tab2.Name = "textBox������������Tab2";
            this.textBox������������Tab2.Size = new System.Drawing.Size(276, 21);
            this.textBox������������Tab2.TabIndex = 0;
            this.toolTip1.SetToolTip(this.textBox������������Tab2, "������� ����� ��� ������.");
            this.textBox������������Tab2.TextChanged += new System.EventHandler(this.textBox������������Tab2_TextChanged);
            // 
            // button��������������������Tab2
            // 
            this.button��������������������Tab2.Location = new System.Drawing.Point(208, 51);
            this.button��������������������Tab2.Name = "button��������������������Tab2";
            this.button��������������������Tab2.Size = new System.Drawing.Size(64, 22);
            this.button��������������������Tab2.TabIndex = 3;
            this.button��������������������Tab2.Text = "��������";
            this.toolTip1.SetToolTip(this.button��������������������Tab2, "�������� ������� ������.");
            this.button��������������������Tab2.Click += new System.EventHandler(this.button��������������������Tab2_Click_1);
            // 
            // textBox�������������������������������
            // 
            this.textBox�������������������������������.Dock = System.Windows.Forms.DockStyle.Top;
            this.textBox�������������������������������.Location = new System.Drawing.Point(0, 0);
            this.textBox�������������������������������.Name = "textBox�������������������������������";
            this.textBox�������������������������������.Size = new System.Drawing.Size(267, 21);
            this.textBox�������������������������������.TabIndex = 1;
            this.toolTip1.SetToolTip(this.textBox�������������������������������, "������� ����� ��� ������.");
            this.textBox�������������������������������.TextChanged += new System.EventHandler(this.textBox�������������������������������_TextChanged);
            // 
            // button���������������������������������������
            // 
            this.button���������������������������������������.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button���������������������������������������.Location = new System.Drawing.Point(200, 27);
            this.button���������������������������������������.Name = "button���������������������������������������";
            this.button���������������������������������������.Size = new System.Drawing.Size(64, 22);
            this.button���������������������������������������.TabIndex = 7;
            this.button���������������������������������������.Text = "��������";
            this.toolTip1.SetToolTip(this.button���������������������������������������, "�������� ������� ������.");
            this.button���������������������������������������.Click += new System.EventHandler(this.button���������������������������������������_Click);
            // 
            // panel4Tab1
            // 
            this.panel4Tab1.Controls.Add(this.dataGrid����������������);
            this.panel4Tab1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel4Tab1.Location = new System.Drawing.Point(0, 0);
            this.panel4Tab1.Name = "panel4Tab1";
            this.panel4Tab1.Size = new System.Drawing.Size(740, 150);
            this.panel4Tab1.TabIndex = 2;
            // 
            // dataGrid����������������
            // 
            this.dataGrid����������������.CaptionFont = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGrid����������������.CaptionText = "�������� ��������� ��������� ������������";
            this.dataGrid����������������.DataMember = "";
            this.dataGrid����������������.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGrid����������������.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGrid����������������.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid����������������.Location = new System.Drawing.Point(0, 0);
            this.dataGrid����������������.Name = "dataGrid����������������";
            this.dataGrid����������������.ReadOnly = true;
            this.dataGrid����������������.Size = new System.Drawing.Size(740, 150);
            this.dataGrid����������������.TabIndex = 0;
            this.dataGrid����������������.TableStyles.AddRange(new System.Windows.Forms.DataGridTableStyle[] {
            this.dataGridTableStyle����������������});
            this.dataGrid����������������.Resize += new System.EventHandler(this.dataGrid����������������_Resize);
            this.dataGrid����������������.DoubleClick += new System.EventHandler(this.dataGrid����������������_DoubleClick);
            this.dataGrid����������������.CurrentCellChanged += new System.EventHandler(this.dataGrid����������������_CurrentCellChanged);
            this.dataGrid����������������.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dataGrid����������������_MouseUp);
            this.dataGrid����������������.Leave += new System.EventHandler(this.dataGrid����������������_Leave);
            // 
            // dataGridTableStyle����������������
            // 
            this.dataGridTableStyle����������������.AlternatingBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.dataGridTableStyle����������������.DataGrid = this.dataGrid����������������;
            this.dataGridTableStyle����������������.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
            this.dataGridTextBoxColumn1,
            this.dataGridTextBoxColumn2,
            this.dataGridTextBoxColumn3,
            this.dataGridTextBoxColumn4,
            this.dataGridTextBoxColumn5,
            this.dataGridTextBoxColumn6,
            this.dataGridTextBoxColumn7,
            this.dataGridTextBoxColumn8});
            this.dataGridTableStyle����������������.HeaderFont = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridTableStyle����������������.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridTableStyle����������������.MappingName = "�������";
            this.dataGridTableStyle����������������.ReadOnly = true;
            this.dataGridTableStyle����������������.RowHeadersVisible = false;
            // 
            // dataGridTextBoxColumn1
            // 
            this.dataGridTextBoxColumn1.Format = "";
            this.dataGridTextBoxColumn1.FormatInfo = null;
            this.dataGridTextBoxColumn1.HeaderText = "��������";
            this.dataGridTextBoxColumn1.MappingName = "�����������������";
            this.dataGridTextBoxColumn1.NullText = "";
            this.dataGridTextBoxColumn1.Width = 75;
            // 
            // dataGridTextBoxColumn2
            // 
            this.dataGridTextBoxColumn2.Format = "";
            this.dataGridTextBoxColumn2.FormatInfo = null;
            this.dataGridTextBoxColumn2.HeaderText = "����-�";
            this.dataGridTextBoxColumn2.MappingName = "����������������������";
            this.dataGridTextBoxColumn2.NullText = "";
            this.dataGridTextBoxColumn2.Width = 75;
            // 
            // dataGridTextBoxColumn3
            // 
            this.dataGridTextBoxColumn3.Format = "";
            this.dataGridTextBoxColumn3.FormatInfo = null;
            this.dataGridTextBoxColumn3.HeaderText = "���� ����.";
            this.dataGridTextBoxColumn3.MappingName = "����������";
            this.dataGridTextBoxColumn3.NullText = "";
            this.dataGridTextBoxColumn3.Width = 65;
            // 
            // dataGridTextBoxColumn4
            // 
            this.dataGridTextBoxColumn4.Format = "";
            this.dataGridTextBoxColumn4.FormatInfo = null;
            this.dataGridTextBoxColumn4.HeaderText = "� �����.";
            this.dataGridTextBoxColumn4.MappingName = "����������";
            this.dataGridTextBoxColumn4.NullText = "";
            this.dataGridTextBoxColumn4.Width = 65;
            // 
            // dataGridTextBoxColumn5
            // 
            this.dataGridTextBoxColumn5.Format = "";
            this.dataGridTextBoxColumn5.FormatInfo = null;
            this.dataGridTextBoxColumn5.HeaderText = "���� ����.";
            this.dataGridTextBoxColumn5.MappingName = "����������";
            this.dataGridTextBoxColumn5.NullText = "";
            this.dataGridTextBoxColumn5.Width = 65;
            // 
            // dataGridTextBoxColumn6
            // 
            this.dataGridTextBoxColumn6.Format = "";
            this.dataGridTextBoxColumn6.FormatInfo = null;
            this.dataGridTextBoxColumn6.HeaderText = "� ����.";
            this.dataGridTextBoxColumn6.MappingName = "���������";
            this.dataGridTextBoxColumn6.NullText = "";
            this.dataGridTextBoxColumn6.Width = 75;
            // 
            // dataGridTextBoxColumn7
            // 
            this.dataGridTextBoxColumn7.Format = "";
            this.dataGridTextBoxColumn7.FormatInfo = null;
            this.dataGridTextBoxColumn7.HeaderText = "����������";
            this.dataGridTextBoxColumn7.MappingName = "�����������������";
            this.dataGridTextBoxColumn7.NullText = "";
            this.dataGridTextBoxColumn7.Width = 240;
            // 
            // dataGridTextBoxColumn8
            // 
            this.dataGridTextBoxColumn8.Format = "";
            this.dataGridTextBoxColumn8.FormatInfo = null;
            this.dataGridTextBoxColumn8.HeaderText = "��������";
            this.dataGridTextBoxColumn8.MappingName = "��������������";
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
            this.menuItem��������������});
            this.menuItem1.Text = "����";
            // 
            // menuItem8
            // 
            this.menuItem8.Index = 0;
            this.menuItem8.Text = "���";
            this.menuItem8.Click += new System.EventHandler(this.menuItem8_Click);
            // 
            // menuItem9
            // 
            this.menuItem9.Index = 1;
            this.menuItem9.Text = "-";
            // 
            // menuItem��������������
            // 
            this.menuItem��������������.Index = 2;
            this.menuItem��������������.Text = "�����";
            this.menuItem��������������.Click += new System.EventHandler(this.menuItem�������_Click);
            // 
            // menuItem2
            // 
            this.menuItem2.Index = 1;
            this.menuItem2.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem�������������������������,
            this.menuItem������������������������,
            this.menuItem���������������������,
            this.menuItem��������������������,
            this.menuItem11,
            this.menuItem12,
            this.menuItem31});
            this.menuItem2.Text = "�����������";
            // 
            // menuItem�������������������������
            // 
            this.menuItem�������������������������.Index = 0;
            this.menuItem�������������������������.Text = "��������";
            this.menuItem�������������������������.Click += new System.EventHandler(this.menuItem�������������������������_Click);
            // 
            // menuItem������������������������
            // 
            this.menuItem������������������������.Index = 1;
            this.menuItem������������������������.Text = "�������������";
            this.menuItem������������������������.Click += new System.EventHandler(this.menuItem������������������������_Click);
            // 
            // menuItem���������������������
            // 
            this.menuItem���������������������.Index = 2;
            this.menuItem���������������������.Text = "����������";
            this.menuItem���������������������.Click += new System.EventHandler(this.menuItem���������������������_Click);
            // 
            // menuItem��������������������
            // 
            this.menuItem��������������������.Index = 3;
            this.menuItem��������������������.Text = "���� ����������";
            this.menuItem��������������������.Click += new System.EventHandler(this.menuItem��������������������_Click);
            // 
            // menuItem11
            // 
            this.menuItem11.Index = 4;
            this.menuItem11.Text = "������������ ������";
            this.menuItem11.Click += new System.EventHandler(this.menuItem11_Click);
            // 
            // menuItem12
            // 
            this.menuItem12.Index = 5;
            this.menuItem12.Text = "���� ��������� ������������ ������";
            this.menuItem12.Click += new System.EventHandler(this.menuItem12_Click);
            // 
            // menuItem31
            // 
            this.menuItem31.Index = 6;
            this.menuItem31.Text = "���������� ������������ �����";
            this.menuItem31.Click += new System.EventHandler(this.menuItem31_Click);
            // 
            // menuItem20
            // 
            this.menuItem20.Index = 2;
            this.menuItem20.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem21,
            this.menuItem22,
            this.menuItem23});
            this.menuItem20.Text = "������";
            this.menuItem20.Click += new System.EventHandler(this.menuItem20_Click);
            // 
            // menuItem21
            // 
            this.menuItem21.Index = 0;
            this.menuItem21.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem24,
            this.menuItem25});
            this.menuItem21.Text = "����� ������";
            this.menuItem21.Click += new System.EventHandler(this.menuItem21_Click);
            // 
            // menuItem24
            // 
            this.menuItem24.Index = 0;
            this.menuItem24.Text = "����������� �����������";
            this.menuItem24.Click += new System.EventHandler(this.menuItem24_Click);
            // 
            // menuItem25
            // 
            this.menuItem25.Index = 1;
            this.menuItem25.Text = "��������� � �������� ������� ����������";
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
            this.menuItem22.Text = "������ �� �������� ���������������";
            this.menuItem22.Click += new System.EventHandler(this.menuItem22_Click);
            // 
            // menuItem26
            // 
            this.menuItem26.Index = 0;
            this.menuItem26.Text = "���������� �� �������� ���������������";
            this.menuItem26.Click += new System.EventHandler(this.menuItem26_Click);
            // 
            // menuItem27
            // 
            this.menuItem27.Index = 1;
            this.menuItem27.Text = "����� � �������� ����������";
            this.menuItem27.Click += new System.EventHandler(this.menuItem27_Click);
            // 
            // menuItem28
            // 
            this.menuItem28.Index = 2;
            this.menuItem28.Text = "������ ��������";
            this.menuItem28.Click += new System.EventHandler(this.menuItem28_Click);
            // 
            // menuItem23
            // 
            this.menuItem23.Index = 2;
            this.menuItem23.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem29,
            this.menuItem30,
            this.menuItem32});
            this.menuItem23.Text = "������ �� ��������� ���������������";
            this.menuItem23.Click += new System.EventHandler(this.menuItem23_Click);
            // 
            // menuItem29
            // 
            this.menuItem29.Index = 0;
            this.menuItem29.Text = "���������� �� ��������� ���������������";
            this.menuItem29.Click += new System.EventHandler(this.menuItem29_Click);
            // 
            // menuItem30
            // 
            this.menuItem30.Index = 1;
            this.menuItem30.Text = "����� �� ��������� ����������";
            this.menuItem30.Click += new System.EventHandler(this.menuItem30_Click);
            // 
            // menuItem32
            // 
            this.menuItem32.Index = 2;
            this.menuItem32.Text = "������ ����� �������� ������������ ������";
            this.menuItem32.Click += new System.EventHandler(this.menuItem32_Click);
            // 
            // menuItem17
            // 
            this.menuItem17.Index = 3;
            this.menuItem17.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem18,
            this.menuItem19});
            this.menuItem17.Text = "��������";
            this.menuItem17.Visible = false;
            // 
            // menuItem18
            // 
            this.menuItem18.Index = 0;
            this.menuItem18.Text = "��������";
            this.menuItem18.Click += new System.EventHandler(this.menuItem18_Click);
            // 
            // menuItem19
            // 
            this.menuItem19.Index = 1;
            this.menuItem19.Text = "���������";
            this.menuItem19.Click += new System.EventHandler(this.menuItem19_Click);
            // 
            // menuItem3
            // 
            this.menuItem3.Index = 4;
            this.menuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
            this.menuItem����������������,
            this.menuItem����������������������,
            this.menuItemContext��������������,
            this.menuItem5,
            this.menuItem4,
            this.menuItem7,
            this.menuItem6,
            this.menuItem10,
            this.menuItem13,
            this.menuItem14,
            this.menuItem15,
            this.menuItem16});
            this.menuItem3.Text = "������";
            this.menuItem3.Visible = false;
            // 
            // menuItem����������������
            // 
            this.menuItem����������������.Index = 0;
            this.menuItem����������������.Text = "��������� � ��������� ������� ����������";
            this.menuItem����������������.Click += new System.EventHandler(this.menuItem����������������_Click);
            // 
            // menuItem����������������������
            // 
            this.menuItem����������������������.Index = 1;
            this.menuItem����������������������.Text = "����������� �����������";
            this.menuItem����������������������.Click += new System.EventHandler(this.menuItem����������������������_Click);
            // 
            // menuItemContext��������������
            // 
            this.menuItemContext��������������.Index = 2;
            this.menuItemContext��������������.Text = "������ ��������";
            this.menuItemContext��������������.Click += new System.EventHandler(this.menuItemContext��������������_Click);
            // 
            // menuItem5
            // 
            this.menuItem5.Index = 3;
            this.menuItem5.Text = "���������� �� ������������";
            this.menuItem5.Click += new System.EventHandler(this.menuItem5_Click);
            // 
            // menuItem4
            // 
            this.menuItem4.Index = 4;
            this.menuItem4.Text = "���������� �� ���������������";
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
            this.menuItem6.Text = "���������� �� ��������� ���������������";
            this.menuItem6.Click += new System.EventHandler(this.menuItem6_Click);
            // 
            // menuItem10
            // 
            this.menuItem10.Index = 7;
            this.menuItem10.Text = "������ ����� �������� ������������ ������";
            this.menuItem10.Click += new System.EventHandler(this.menuItem10_Click);
            // 
            // menuItem13
            // 
            this.menuItem13.Index = 8;
            this.menuItem13.Text = "������ ��������";
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
            this.menuItem15.Text = "����� � ����������";
            this.menuItem15.Click += new System.EventHandler(this.menuItem15_Click);
            // 
            // menuItem16
            // 
            this.menuItem16.Index = 11;
            this.menuItem16.Text = "����� ���������� �� ������������";
            this.menuItem16.Click += new System.EventHandler(this.menuItem16_Click);
            // 
            // tabControl�����������������
            // 
            this.tabControl�����������������.Controls.Add(this.tabPage1);
            this.tabControl�����������������.Controls.Add(this.tabPage2);
            this.tabControl�����������������.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl�����������������.Location = new System.Drawing.Point(3, 3);
            this.tabControl�����������������.Name = "tabControl�����������������";
            this.tabControl�����������������.SelectedIndex = 0;
            this.tabControl�����������������.Size = new System.Drawing.Size(748, 286);
            this.tabControl�����������������.TabIndex = 3;
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
            this.tabPage1.Text = "������� ���������";
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
            this.panel1.Controls.Add(this.label����Tab1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(280, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(460, 110);
            this.panel1.TabIndex = 1;
            // 
            // label����Tab1
            // 
            this.label����Tab1.BackColor = System.Drawing.SystemColors.Window;
            this.label����Tab1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label����Tab1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label����Tab1.Location = new System.Drawing.Point(0, 0);
            this.label����Tab1.Name = "label����Tab1";
            this.label����Tab1.ReadOnly = true;
            this.label����Tab1.Size = new System.Drawing.Size(456, 106);
            this.label����Tab1.TabIndex = 0;
            this.label����Tab1.TabStop = false;
            this.label����Tab1.Text = "";
            this.label����Tab1.Leave += new System.EventHandler(this.label����Tab1_Leave);
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
            this.tabPage2.Text = "��������� \"� ����\"";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // panel1Tab2
            // 
            this.panel1Tab2.Controls.Add(this.dataGrid��������������);
            this.panel1Tab2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1Tab2.Location = new System.Drawing.Point(0, 0);
            this.panel1Tab2.Name = "panel1Tab2";
            this.panel1Tab2.Size = new System.Drawing.Size(740, 150);
            this.panel1Tab2.TabIndex = 3;
            // 
            // dataGrid��������������
            // 
            this.dataGrid��������������.CaptionFont = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGrid��������������.CaptionText = "�������� ��������� ��������� \"� ����\"";
            this.dataGrid��������������.DataMember = "";
            this.dataGrid��������������.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGrid��������������.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGrid��������������.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid��������������.Location = new System.Drawing.Point(0, 0);
            this.dataGrid��������������.Name = "dataGrid��������������";
            this.dataGrid��������������.ReadOnly = true;
            this.dataGrid��������������.Size = new System.Drawing.Size(740, 150);
            this.dataGrid��������������.TabIndex = 0;
            this.dataGrid��������������.TableStyles.AddRange(new System.Windows.Forms.DataGridTableStyle[] {
            this.dataGridTableStyle��������������});
            this.dataGrid��������������.Resize += new System.EventHandler(this.dataGrid��������������_Resize);
            this.dataGrid��������������.DoubleClick += new System.EventHandler(this.dataGrid��������������_DoubleClick);
            this.dataGrid��������������.CurrentCellChanged += new System.EventHandler(this.dataGrid��������������_CurrentCellChanged);
            this.dataGrid��������������.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dataGrid��������������_MouseUp);
            this.dataGrid��������������.Leave += new System.EventHandler(this.dataGrid��������������_Leave);
            // 
            // dataGridTableStyle��������������
            // 
            this.dataGridTableStyle��������������.AlternatingBackColor = System.Drawing.Color.LavenderBlush;
            this.dataGridTableStyle��������������.DataGrid = this.dataGrid��������������;
            this.dataGridTableStyle��������������.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
            this.dataGridTextBoxColumn9,
            this.dataGridTextBoxColumn10,
            this.dataGridTextBoxColumn11,
            this.dataGridTextBoxColumn12,
            this.dataGridTextBoxColumn13,
            this.dataGridTextBoxColumn14,
            this.dataGridTextBoxColumn15,
            this.dataGridTextBoxColumn16});
            this.dataGridTableStyle��������������.HeaderFont = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridTableStyle��������������.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridTableStyle��������������.MappingName = "�������";
            this.dataGridTableStyle��������������.RowHeadersVisible = false;
            // 
            // dataGridTextBoxColumn9
            // 
            this.dataGridTextBoxColumn9.Format = "";
            this.dataGridTextBoxColumn9.FormatInfo = null;
            this.dataGridTextBoxColumn9.HeaderText = "��������";
            this.dataGridTextBoxColumn9.MappingName = "�����������������";
            this.dataGridTextBoxColumn9.NullText = "";
            this.dataGridTextBoxColumn9.Width = 75;
            // 
            // dataGridTextBoxColumn10
            // 
            this.dataGridTextBoxColumn10.Format = "";
            this.dataGridTextBoxColumn10.FormatInfo = null;
            this.dataGridTextBoxColumn10.HeaderText = "����-�";
            this.dataGridTextBoxColumn10.MappingName = "����������������������";
            this.dataGridTextBoxColumn10.NullText = "";
            this.dataGridTextBoxColumn10.Width = 75;
            // 
            // dataGridTextBoxColumn11
            // 
            this.dataGridTextBoxColumn11.Format = "";
            this.dataGridTextBoxColumn11.FormatInfo = null;
            this.dataGridTextBoxColumn11.HeaderText = "���� ����.";
            this.dataGridTextBoxColumn11.MappingName = "����������";
            this.dataGridTextBoxColumn11.NullText = "";
            this.dataGridTextBoxColumn11.Width = 65;
            // 
            // dataGridTextBoxColumn12
            // 
            this.dataGridTextBoxColumn12.Format = "";
            this.dataGridTextBoxColumn12.FormatInfo = null;
            this.dataGridTextBoxColumn12.HeaderText = "� �����.";
            this.dataGridTextBoxColumn12.MappingName = "����������";
            this.dataGridTextBoxColumn12.NullText = "";
            this.dataGridTextBoxColumn12.Width = 65;
            // 
            // dataGridTextBoxColumn13
            // 
            this.dataGridTextBoxColumn13.Format = "";
            this.dataGridTextBoxColumn13.FormatInfo = null;
            this.dataGridTextBoxColumn13.HeaderText = "���� ������.";
            this.dataGridTextBoxColumn13.MappingName = "����������";
            this.dataGridTextBoxColumn13.NullText = "";
            this.dataGridTextBoxColumn13.Width = 65;
            // 
            // dataGridTextBoxColumn14
            // 
            this.dataGridTextBoxColumn14.Format = "";
            this.dataGridTextBoxColumn14.FormatInfo = null;
            this.dataGridTextBoxColumn14.HeaderText = "� ����.";
            this.dataGridTextBoxColumn14.MappingName = "���������";
            this.dataGridTextBoxColumn14.NullText = "";
            this.dataGridTextBoxColumn14.Width = 75;
            // 
            // dataGridTextBoxColumn15
            // 
            this.dataGridTextBoxColumn15.Format = "";
            this.dataGridTextBoxColumn15.FormatInfo = null;
            this.dataGridTextBoxColumn15.HeaderText = "����������";
            this.dataGridTextBoxColumn15.MappingName = "�����������������";
            this.dataGridTextBoxColumn15.NullText = "";
            this.dataGridTextBoxColumn15.Width = 180;
            // 
            // dataGridTextBoxColumn16
            // 
            this.dataGridTextBoxColumn16.Format = "";
            this.dataGridTextBoxColumn16.FormatInfo = null;
            this.dataGridTextBoxColumn16.HeaderText = "���������";
            this.dataGridTextBoxColumn16.MappingName = "���������";
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
            this.panel7.Controls.Add(this.label����Tab2);
            this.panel7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel7.Location = new System.Drawing.Point(280, 0);
            this.panel7.Name = "panel7";
            this.panel7.Size = new System.Drawing.Size(460, 110);
            this.panel7.TabIndex = 4;
            // 
            // label����Tab2
            // 
            this.label����Tab2.BackColor = System.Drawing.SystemColors.Window;
            this.label����Tab2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label����Tab2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label����Tab2.Location = new System.Drawing.Point(0, 0);
            this.label����Tab2.Name = "label����Tab2";
            this.label����Tab2.ReadOnly = true;
            this.label����Tab2.Size = new System.Drawing.Size(456, 106);
            this.label����Tab2.TabIndex = 0;
            this.label����Tab2.TabStop = false;
            this.label����Tab2.Text = "";
            this.label����Tab2.Leave += new System.EventHandler(this.label����Tab2_Leave);
            // 
            // panel4
            // 
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel4.Controls.Add(this.checkBox��������������);
            this.panel4.Controls.Add(this.comboBox��������������);
            this.panel4.Controls.Add(this.label�������������������������Tab2);
            this.panel4.Controls.Add(this.button��������������������Tab2);
            this.panel4.Controls.Add(this.panel6);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(280, 110);
            this.panel4.TabIndex = 3;
            // 
            // checkBox��������������
            // 
            this.checkBox��������������.AutoSize = true;
            this.checkBox��������������.Location = new System.Drawing.Point(4, 55);
            this.checkBox��������������.Name = "checkBox��������������";
            this.checkBox��������������.Size = new System.Drawing.Size(195, 17);
            this.checkBox��������������.TabIndex = 6;
            this.checkBox��������������.Text = "������ ������ ���������������";
            this.checkBox��������������.UseVisualStyleBackColor = true;
            this.checkBox��������������.CheckedChanged += new System.EventHandler(this.checkBox��������������_CheckedChanged);
            // 
            // comboBox��������������
            // 
            this.comboBox��������������.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBox��������������.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBox��������������.DisplayMember = "����������������������";
            this.comboBox��������������.DropDownHeight = 400;
            this.comboBox��������������.DropDownWidth = 400;
            this.comboBox��������������.FormattingEnabled = true;
            this.comboBox��������������.IntegralHeight = false;
            this.comboBox��������������.Location = new System.Drawing.Point(0, 27);
            this.comboBox��������������.Name = "comboBox��������������";
            this.comboBox��������������.Size = new System.Drawing.Size(269, 21);
            this.comboBox��������������.TabIndex = 5;
            this.comboBox��������������.ValueMember = "����������������������";
            // 
            // label�������������������������Tab2
            // 
            this.label�������������������������Tab2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.label�������������������������Tab2.Location = new System.Drawing.Point(0, 89);
            this.label�������������������������Tab2.Name = "label�������������������������Tab2";
            this.label�������������������������Tab2.Size = new System.Drawing.Size(276, 17);
            this.label�������������������������Tab2.TabIndex = 4;
            // 
            // panel6
            // 
            this.panel6.Controls.Add(this.textBox������������Tab2);
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
            this.dataGridTextBoxColumn��������,
            this.dataGridTextBoxColumn�������������,
            this.dataGridTextBoxColumn����������,
            this.dataGridTextBoxColumn����������,
            this.dataGridTextBoxColumn����������,
            this.dataGridTextBoxColumn���������,
            this.dataGridTextBoxColumn����������,
            this.dataGridTextBoxColumn��������,
            this.dataGridBoolColumn�����});
            this.dataGridTableStyle2.HeaderFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGridTableStyle2.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridTableStyle2.MappingName = "�������";
            this.dataGridTableStyle2.ReadOnly = true;
            this.dataGridTableStyle2.RowHeadersVisible = false;
            // 
            // dataGridTextBoxColumn��������
            // 
            this.dataGridTextBoxColumn��������.Format = "";
            this.dataGridTextBoxColumn��������.FormatInfo = null;
            this.dataGridTextBoxColumn��������.HeaderText = "��������";
            this.dataGridTextBoxColumn��������.MappingName = "�����������������";
            this.dataGridTextBoxColumn��������.NullText = "";
            this.dataGridTextBoxColumn��������.ReadOnly = true;
            this.dataGridTextBoxColumn��������.Width = 75;
            // 
            // dataGridTextBoxColumn�������������
            // 
            this.dataGridTextBoxColumn�������������.Format = "";
            this.dataGridTextBoxColumn�������������.FormatInfo = null;
            this.dataGridTextBoxColumn�������������.HeaderText = "����-�";
            this.dataGridTextBoxColumn�������������.MappingName = "����������������������";
            this.dataGridTextBoxColumn�������������.NullText = "";
            this.dataGridTextBoxColumn�������������.ReadOnly = true;
            this.dataGridTextBoxColumn�������������.Width = 75;
            // 
            // dataGridTextBoxColumn����������
            // 
            this.dataGridTextBoxColumn����������.Format = "";
            this.dataGridTextBoxColumn����������.FormatInfo = null;
            this.dataGridTextBoxColumn����������.HeaderText = "����������";
            this.dataGridTextBoxColumn����������.MappingName = "����������";
            this.dataGridTextBoxColumn����������.NullText = "";
            this.dataGridTextBoxColumn����������.ReadOnly = true;
            this.dataGridTextBoxColumn����������.Width = 67;
            // 
            // dataGridTextBoxColumn����������
            // 
            this.dataGridTextBoxColumn����������.Format = "";
            this.dataGridTextBoxColumn����������.FormatInfo = null;
            this.dataGridTextBoxColumn����������.HeaderText = "���������";
            this.dataGridTextBoxColumn����������.MappingName = "����������";
            this.dataGridTextBoxColumn����������.NullText = "";
            this.dataGridTextBoxColumn����������.ReadOnly = true;
            this.dataGridTextBoxColumn����������.Width = 67;
            // 
            // dataGridTextBoxColumn����������
            // 
            this.dataGridTextBoxColumn����������.Format = "";
            this.dataGridTextBoxColumn����������.FormatInfo = null;
            this.dataGridTextBoxColumn����������.HeaderText = "������.";
            this.dataGridTextBoxColumn����������.MappingName = "����������";
            this.dataGridTextBoxColumn����������.NullText = "";
            this.dataGridTextBoxColumn����������.ReadOnly = true;
            this.dataGridTextBoxColumn����������.Width = 65;
            // 
            // dataGridTextBoxColumn���������
            // 
            this.dataGridTextBoxColumn���������.Format = "";
            this.dataGridTextBoxColumn���������.FormatInfo = null;
            this.dataGridTextBoxColumn���������.HeaderText = "�����.";
            this.dataGridTextBoxColumn���������.MappingName = "���������";
            this.dataGridTextBoxColumn���������.NullText = "";
            this.dataGridTextBoxColumn���������.ReadOnly = true;
            this.dataGridTextBoxColumn���������.Width = 65;
            // 
            // dataGridTextBoxColumn����������
            // 
            this.dataGridTextBoxColumn����������.Format = "";
            this.dataGridTextBoxColumn����������.FormatInfo = null;
            this.dataGridTextBoxColumn����������.HeaderText = "����������";
            this.dataGridTextBoxColumn����������.MappingName = "�����������������";
            this.dataGridTextBoxColumn����������.NullText = "";
            this.dataGridTextBoxColumn����������.ReadOnly = true;
            this.dataGridTextBoxColumn����������.Width = 250;
            // 
            // dataGridTextBoxColumn��������
            // 
            this.dataGridTextBoxColumn��������.Format = "";
            this.dataGridTextBoxColumn��������.FormatInfo = null;
            this.dataGridTextBoxColumn��������.HeaderText = "��������";
            this.dataGridTextBoxColumn��������.MappingName = "��������������";
            this.dataGridTextBoxColumn��������.NullText = "";
            this.dataGridTextBoxColumn��������.ReadOnly = true;
            this.dataGridTextBoxColumn��������.Width = 65;
            // 
            // dataGridBoolColumn�����
            // 
            this.dataGridBoolColumn�����.HeaderText = "� ����";
            this.dataGridBoolColumn�����.MappingName = "�����";
            this.dataGridBoolColumn�����.NullText = "";
            this.dataGridBoolColumn�����.Width = 45;
            // 
            // tabControl��������������
            // 
            this.tabControl��������������.Alignment = System.Windows.Forms.TabAlignment.Left;
            this.tabControl��������������.Controls.Add(this.tabPage3);
            this.tabControl��������������.Controls.Add(this.tabPage4);
            this.tabControl��������������.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl��������������.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabControl��������������.ItemSize = new System.Drawing.Size(40, 30);
            this.tabControl��������������.Location = new System.Drawing.Point(0, 0);
            this.tabControl��������������.Multiline = true;
            this.tabControl��������������.Name = "tabControl��������������";
            this.tabControl��������������.SelectedIndex = 0;
            this.tabControl��������������.Size = new System.Drawing.Size(792, 300);
            this.tabControl��������������.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
            this.tabControl��������������.TabIndex = 4;
            this.tabControl��������������.SelectedIndexChanged += new System.EventHandler(this.tabControl��������������_SelectedIndexChanged);
            // 
            // tabPage3
            // 
            this.tabPage3.BackColor = System.Drawing.Color.Transparent;
            this.tabPage3.Controls.Add(this.tabControl�����������������);
            this.tabPage3.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tabPage3.Location = new System.Drawing.Point(34, 4);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(754, 292);
            this.tabPage3.TabIndex = 0;
            this.tabPage3.Text = "��������";
            this.tabPage3.ToolTipText = "��������� ��������";
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
            this.tabPage4.Text = "���������";
            this.tabPage4.ToolTipText = "��������� ���������";
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
            this.splitContainer3.Panel1.Controls.Add(this.dataGrid������������������);
            this.splitContainer3.Panel2MinSize = 0;
            this.splitContainer3.Size = new System.Drawing.Size(744, 198);
            this.splitContainer3.SplitterDistance = 737;
            this.splitContainer3.TabIndex = 1;
            // 
            // dataGrid������������������
            // 
            this.dataGrid������������������.CaptionFont = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGrid������������������.CaptionText = "��������� ���������";
            this.dataGrid������������������.DataMember = "";
            this.dataGrid������������������.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGrid������������������.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dataGrid������������������.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGrid������������������.Location = new System.Drawing.Point(0, 0);
            this.dataGrid������������������.Name = "dataGrid������������������";
            this.dataGrid������������������.ReadOnly = true;
            this.dataGrid������������������.Size = new System.Drawing.Size(737, 198);
            this.dataGrid������������������.TabIndex = 0;
            this.dataGrid������������������.TableStyles.AddRange(new System.Windows.Forms.DataGridTableStyle[] {
            this.dataGridTableStyle������������������});
            this.dataGrid������������������.Resize += new System.EventHandler(this.dataGrid������������������_Resize);
            this.dataGrid������������������.DoubleClick += new System.EventHandler(this.dataGrid������������������_DoubleClick);
            this.dataGrid������������������.CurrentCellChanged += new System.EventHandler(this.dataGrid������������������_CurrentCellChanged);
            this.dataGrid������������������.MouseUp += new System.Windows.Forms.MouseEventHandler(this.dataGrid������������������_MouseUp);
            this.dataGrid������������������.Leave += new System.EventHandler(this.dataGrid������������������_Leave);
            // 
            // dataGridTableStyle������������������
            // 
            this.dataGridTableStyle������������������.AlternatingBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.dataGridTableStyle������������������.DataGrid = this.dataGrid������������������;
            this.dataGridTableStyle������������������.GridColumnStyles.AddRange(new System.Windows.Forms.DataGridColumnStyle[] {
            this.dataGridTextBoxColumn����������������,
            this.dataGridTextBoxColumn�����������,
            this.dataGridTextBoxColumn����������������������,
            this.dataGridTextBoxColumn����������������,
            this.dataGridTextBoxColumn��������������������});
            this.dataGridTableStyle������������������.HeaderForeColor = System.Drawing.SystemColors.ControlText;
            this.dataGridTableStyle������������������.MappingName = "��������������������������";
            this.dataGridTableStyle������������������.RowHeadersVisible = false;
            // 
            // dataGridTextBoxColumn����������������
            // 
            this.dataGridTextBoxColumn����������������.Format = "";
            this.dataGridTextBoxColumn����������������.FormatInfo = null;
            this.dataGridTextBoxColumn����������������.HeaderText = "���� ����.";
            this.dataGridTextBoxColumn����������������.MappingName = "����";
            this.dataGridTextBoxColumn����������������.NullText = "";
            this.dataGridTextBoxColumn����������������.ReadOnly = true;
            this.dataGridTextBoxColumn����������������.Width = 75;
            // 
            // dataGridTextBoxColumn�����������
            // 
            this.dataGridTextBoxColumn�����������.Format = "";
            this.dataGridTextBoxColumn�����������.FormatInfo = null;
            this.dataGridTextBoxColumn�����������.HeaderText = "����� ���������";
            this.dataGridTextBoxColumn�����������.MappingName = "��������������";
            this.dataGridTextBoxColumn�����������.NullText = "";
            this.dataGridTextBoxColumn�����������.ReadOnly = true;
            this.dataGridTextBoxColumn�����������.Width = 110;
            // 
            // dataGridTextBoxColumn����������������������
            // 
            this.dataGridTextBoxColumn����������������������.Format = "";
            this.dataGridTextBoxColumn����������������������.FormatInfo = null;
            this.dataGridTextBoxColumn����������������������.HeaderText = "�������";
            this.dataGridTextBoxColumn����������������������.MappingName = "����������������";
            this.dataGridTextBoxColumn����������������������.NullText = "";
            this.dataGridTextBoxColumn����������������������.ReadOnly = true;
            this.dataGridTextBoxColumn����������������������.Width = 150;
            // 
            // dataGridTextBoxColumn����������������
            // 
            this.dataGridTextBoxColumn����������������.Format = "";
            this.dataGridTextBoxColumn����������������.FormatInfo = null;
            this.dataGridTextBoxColumn����������������.HeaderText = "����������";
            this.dataGridTextBoxColumn����������������.MappingName = "����������";
            this.dataGridTextBoxColumn����������������.NullText = "";
            this.dataGridTextBoxColumn����������������.ReadOnly = true;
            this.dataGridTextBoxColumn����������������.Width = 300;
            // 
            // dataGridTextBoxColumn��������������������
            // 
            this.dataGridTextBoxColumn��������������������.Format = "";
            this.dataGridTextBoxColumn��������������������.FormatInfo = null;
            this.dataGridTextBoxColumn��������������������.HeaderText = "�������� ��������";
            this.dataGridTextBoxColumn��������������������.MappingName = "���������������������������";
            this.dataGridTextBoxColumn��������������������.NullText = "";
            this.dataGridTextBoxColumn��������������������.ReadOnly = true;
            this.dataGridTextBoxColumn��������������������.Width = 110;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.comboBox��������������);
            this.splitContainer2.Panel1.Controls.Add(this.button���������������������������������������);
            this.splitContainer2.Panel1.Controls.Add(this.label��������������������������������������������);
            this.splitContainer2.Panel1.Controls.Add(this.textBox�������������������������������);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.label����Tab3);
            this.splitContainer2.Size = new System.Drawing.Size(744, 76);
            this.splitContainer2.SplitterDistance = 267;
            this.splitContainer2.TabIndex = 0;
            // 
            // comboBox��������������
            // 
            this.comboBox��������������.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox��������������.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBox��������������.FormattingEnabled = true;
            this.comboBox��������������.Items.AddRange(new object[] {
            "���� ���",
            "������",
            "�������",
            "����",
            "������",
            "���",
            "����",
            "����",
            "������",
            "��������",
            "�������",
            "������",
            "�������"});
            this.comboBox��������������.Location = new System.Drawing.Point(1, 27);
            this.comboBox��������������.Name = "comboBox��������������";
            this.comboBox��������������.Size = new System.Drawing.Size(189, 24);
            this.comboBox��������������.TabIndex = 8;
            this.comboBox��������������.SelectedIndexChanged += new System.EventHandler(this.comboBox��������������_SelectedIndexChanged);
            // 
            // label��������������������������������������������
            // 
            this.label��������������������������������������������.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.label��������������������������������������������.Location = new System.Drawing.Point(0, 56);
            this.label��������������������������������������������.Name = "label��������������������������������������������";
            this.label��������������������������������������������.Size = new System.Drawing.Size(267, 20);
            this.label��������������������������������������������.TabIndex = 6;
            // 
            // label����Tab3
            // 
            this.label����Tab3.BackColor = System.Drawing.SystemColors.Window;
            this.label����Tab3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label����Tab3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label����Tab3.Location = new System.Drawing.Point(0, 0);
            this.label����Tab3.Name = "label����Tab3";
            this.label����Tab3.ReadOnly = true;
            this.label����Tab3.Size = new System.Drawing.Size(473, 76);
            this.label����Tab3.TabIndex = 1;
            this.label����Tab3.TabStop = false;
            this.label����Tab3.Text = "";
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
            this.checkBox1.Text = "�������";
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
            this.checkBox2.Text = "���� ���";
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
            this.label2.Text = "������ �� ����";
            // 
            // ds11
            // 
            this.ds11.DataSetName = "DS1";
            this.ds11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // menuItem33
            // 
            this.menuItem33.Index = 3;
            this.menuItem33.Text = "������ ����� �������� ������������ ������";
            this.menuItem33.Click += new System.EventHandler(this.menuItem33_Click_1);
            // 
            // Form�������
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(792, 300);
            this.Controls.Add(this.tabControl��������������);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Menu = this.mainMenu1;
            this.Name = "Form�������";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "����������� ���������������";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form�������_FormClosing);
            this.Load += new System.EventHandler(this.Form�������_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataView�������������������)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4Tab1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid����������������)).EndInit();
            this.tabControl�����������������.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panel1Tab1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.panel1Tab2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid��������������)).EndInit();
            this.panel5.ResumeLayout(false);
            this.panel7.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panel6.ResumeLayout(false);
            this.panel6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataView���������������������)).EndInit();
            this.tabControl��������������.ResumeLayout(false);
            this.tabPage3.ResumeLayout(false);
            this.tabPage4.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer3.Panel1.ResumeLayout(false);
            this.splitContainer3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid������������������)).EndInit();
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel1.PerformLayout();
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataView������������������)).EndInit();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion


        /// <summary>
        /// ����� ����� � ����������
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (��������������������())
            {
                MessageBox.Show(null, "��������� ��� ��������", "����������� ���������������");
                return;
            }

            Application.Run(new Form�������());
        }

        #region ������

        /// <summary>
        /// ������������ � ��. 
        /// ������� ������� � ��������. 
        /// ��������� ������� ������� �� ��.
        /// ������� �������� � ���������� ��� � �������� ��������� ������ � ��������
        /// </summary>
        private void ��������������������������()
        {
            ������������� = new System.Threading.Thread(new System.Threading.ThreadStart(����������������������));
            �������������.Start();
            try
            {
                //=============������� ds11
                //��������
                this.ds11.��������.Clear();
                this.ds11.�������.Clear();
                this.ds11.���������.Clear();

                this.ds11.��������������.Clear();
                this.ds11.����������.Clear();


                //���������
                this.ds11.��������������������������.Clear();
                this.ds11.�����������������.Clear();

                this.ds11.���������������������.Clear();
                this.ds11.������������������������������������.Clear();

                //�������� ds11 ������ ������� =====================
                //��������
                DS1TableAdapters.���������TableAdapter ���������TableAdapter = new RegKor.DS1TableAdapters.���������TableAdapter();
                ���������TableAdapter.Fill(ds11.���������);

                //��������� ���� � ���������������� ����� ���� ������������ = "������" ����� 2012 ��� ��� ������
                //���� ������������ = "�����" ����� 2011 ��� ��� ������
                Classess.������������������� ��� = new RegKor.Classess.�������������������();
                bool flag = ���.�������������������������();

                ////if (flag == false)
                ////{
                ////    //���� ��� 2011 ��� ����� ������ �� ��������� �� ��� ���������
                DS1TableAdapters.��������������TableAdapter ��������������TableAdapter = new RegKor.DS1TableAdapters.��������������TableAdapter();
                ��������������TableAdapter.Fill(ds11.��������������);
                ////}

                //////���� ���� = true ������ ��� � app.config 2012 ��� ����� �������

                ////if (flag == true)
                ////{
                ////    Classess.�������������� �������������� = new RegKor.Classess.��������������();
                ////    DataSet ds�������������� = ��������������.�����������������������_DataSet();

                ////    foreach (DataRow row�������������� in ds��������������.Tables[0].Rows)
                ////    {
                ////        DataRow row1 = ds11.��������������.NewRow();
                ////        row1[0] = row��������������[0];
                ////        row1[1] = row��������������[1];
                ////        ds11.��������������.Rows.Add(row1);
                ////    }
                ////}
                ////==========================================

                DS1TableAdapters.����������TableAdapter ����������TableAdapter = new RegKor.DS1TableAdapters.����������TableAdapter();
                ����������TableAdapter.Fill(ds11.����������);

                //�������� � ������ ����������
                ������������ �� = new ������������();
                FillDataSet fillDataSet = new FillDataSet();

                using (SqlConnection con = new SqlConnection(��.�����������������()))
                {
                    con.Open();
                    SqlTransaction transaction = con.BeginTransaction("transactLoad");

                    //�������� ������ �� �� �� 1 ������� ����������� ���� (1.01.2012 ���� ��������� ��������� �� 2013 ���)
                    //DS1TableAdapters.��������TableAdapter ��������TableAdapter = new RegKor.DS1TableAdapters.��������TableAdapter();
                    //��������TableAdapter.Fill(ds11.��������);
                    string query�������� = "select * from dbo.�������� where ���������� >= '" + ������������� + "' and ���������� <= '" + ������������� + "' ";
                    fillDataSet.FillTable(query��������, ds11, "��������", con, transaction);

                    DataTable tabTest = ds11.��������;

                    //�������� ������� �������
                    //DS1TableAdapters.�������TableAdapter �������TableAdapter = new RegKor.DS1TableAdapters.�������TableAdapter();
                    //�������TableAdapter.Fill(ds11.�������);
                    string query������� = "select * from ������� where ���������� >= '" + ������������� + "'  and ���������� <= '" + ������������� + "' ";
                    fillDataSet.FillTable(query�������, ds11, "�������", con, transaction);

                    //�������� �����������������
                    string query����������������� = "select * from ����������������� where ���� >= '" + ������������� + "' ";
                    fillDataSet.FillTable(query�����������������, ds11, "�����������������", con, transaction);

                    //�������� �������� ��������������������������
                    string query�������������������������� = "select * from �������������������������� where ���� >= '" + ������������� + "' ";
                    fillDataSet.FillTable(query��������������������������, ds11, "��������������������������", con, transaction);

                }
                

                //DS1TableAdapters.������������������������������������TableAdapter ����������������TableAdapter = new RegKor.DS1TableAdapters.������������������������������������TableAdapter();
                //����������������TableAdapter.Fill(ds11.������������������������������������);
                


                // ���������
                DS1TableAdapters.���������������������TableAdapter ���������������������TableAdapter = new RegKor.DS1TableAdapters.���������������������TableAdapter();
                ���������������������TableAdapter.Fill(ds11.���������������������);

                //������ ������� ��������� ����
                //DS1TableAdapters.�����������������TableAdapter �����������������TableAdapter = new RegKor.DS1TableAdapters.�����������������TableAdapter();
                //�����������������TableAdapter.Fill(ds11.�����������������);

                //DS1TableAdapters.��������������������������TableAdapter ��������������������������TableAdapter = new RegKor.DS1TableAdapters.��������������������������TableAdapter();
                //��������������������������TableAdapter.Fill(ds11.��������������������������);

                this.Refresh();

                //��������
                dataView�������������������.Table = ds11.�������;
                dataView�������������������.RowFilter = "�����=False AND ���������� >='01.12." + ������������ + "'";
                dataGrid����������������.DataSource = dataView�������������������;

                dataView���������������������.Table = ds11.�������;
                dataView���������������������.RowFilter = "�����=True AND ���������� >='01.12." + ������������ + "'";
                dataGrid��������������.DataSource = dataView���������������������;
                //���������
                dataView������������������.Table = ds11.��������������������������;
                dataGrid������������������.DataSource = dataView������������������;

                //�������� ComboBox��������������
                this.comboBox��������������.DataSource = ds11.��������������.Select("","����������������������");
                

                this.Refresh();

                ����������();
                this.Refresh();

            }
            //catch (Exception exc)
            //{
            //    �������������.Abort();
            //    MessageBox.Show(this, exc.Message + "\n" + exc.Source, "����� \"��������������������������()\"");
            //    this.Enabled = false;
            //    this.menuItem2.Enabled = false;
            //    this.menuItem3.Enabled = false;
            //    string str = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Environment.CurrentDirectory + "\\RegKor.exe").FileVersion;
            //    this.Text = "����������� ���������������. ������: " + str + ". SQL Server: ����������� �� �����������";
            //}
            finally
            {
                �������������.Abort();
            }
        }

        /// <summary>
        /// ������ ������ ������ � ����, ������� ������� � ��������
        /// � ��������� �� ������ �� ����
        /// </summary>
        private void ��������������()
        {
            ������������� = new System.Threading.Thread(new System.Threading.ThreadStart(����������������������));
            �������������.Start();
            this.Refresh();
            try
            {
                //��������
                DS1TableAdapters.���������TableAdapter ���������TableAdapter = new RegKor.DS1TableAdapters.���������TableAdapter();
                ���������TableAdapter.Update(ds11.���������);

                DS1TableAdapters.��������������TableAdapter ��������������TableAdapter = new RegKor.DS1TableAdapters.��������������TableAdapter();
                ��������������TableAdapter.Update(ds11.��������������);

                DS1TableAdapters.����������TableAdapter ����������TableAdapter = new RegKor.DS1TableAdapters.����������TableAdapter();
                ����������TableAdapter.Update(ds11.����������);

                //DS1TableAdapters.��������TableAdapter ��������TableAdapter = new RegKor.DS1TableAdapters.��������TableAdapter();
                //��������TableAdapter.Update(ds11.��������);
                //���������
                DS1TableAdapters.���������������������TableAdapter ���������������������TableAdapter = new RegKor.DS1TableAdapters.���������������������TableAdapter();
                ���������������������TableAdapter.Update(ds11.���������������������);

                //DS1TableAdapters.�����������������TableAdapter �����������������TableAdapter = new RegKor.DS1TableAdapters.�����������������TableAdapter();
                //�����������������TableAdapter.Update(ds11.�����������������);
                this.Refresh();

                //��������
                this.ds11.��������.Clear();
                this.ds11.�������.Clear();
                this.ds11.���������.Clear();
                this.ds11.��������������.Clear();
                this.ds11.����������.Clear();
                //���������
                this.ds11.��������������������������.Clear();
                this.ds11.�����������������.Clear();
                this.ds11.���������������������.Clear();

                //��������
                ���������TableAdapter = new RegKor.DS1TableAdapters.���������TableAdapter();
                ���������TableAdapter.Fill(ds11.���������);
                ��������������TableAdapter = new RegKor.DS1TableAdapters.��������������TableAdapter();
                ��������������TableAdapter.Fill(ds11.��������������);
                ����������TableAdapter = new RegKor.DS1TableAdapters.����������TableAdapter();
                ����������TableAdapter.Fill(ds11.����������);
                //���������
                ���������������������TableAdapter = new RegKor.DS1TableAdapters.���������������������TableAdapter();
                ���������������������TableAdapter.Fill(ds11.���������������������);

                ������������ �� = new ������������();
                FillDataSet fillDataSet = new FillDataSet();

                using (SqlConnection con = new SqlConnection(��.�����������������()))
                {
                    StringBuilder builder = new StringBuilder();

                    con.Open();
                    SqlTransaction transaction = con.BeginTransaction("updateTransaction");
                    //�������� ��������
                    //��������TableAdapter = new RegKor.DS1TableAdapters.��������TableAdapter();
                    //��������TableAdapter.Fill(ds11.��������);
                    string query�������� = "select * from dbo.�������� where ���������� >= '" + ������������� + "' and ���������� <= '" + ������������� + "'  ";
                    fillDataSet.FillTable(query��������, ds11, "��������", con, transaction);

                    //�������� ������� �������
                    //DS1TableAdapters.�������TableAdapter �������TableAdapter = new RegKor.DS1TableAdapters.�������TableAdapter();
                    //�������TableAdapter.Fill(ds11.�������);
                    string query������� = "select * from ������� where ���������� >= '" + ������������� + "'";
                    fillDataSet.FillTable(query�������, ds11, "�������", con, transaction);

                    //���������

                    //�������� �������� ���������
                    //�����������������TableAdapter = new RegKor.DS1TableAdapters.�����������������TableAdapter();
                    //�����������������TableAdapter.Fill(ds11.�����������������);
                    string query����������������� = "select * from ����������������� where ���� >= '" + ������������� + "'  and ���� <= '" + ������������� + "' ";
                    fillDataSet.FillTable(query�����������������, ds11, "�����������������", con, transaction);

                    //�������� ������� ��������������������������
                    //DS1TableAdapters.��������������������������TableAdapter ��������������������������TableAdapter = new RegKor.DS1TableAdapters.��������������������������TableAdapter();
                    //��������������������������TableAdapter.Fill(ds11.��������������������������);
                    string query�������������������������� = "select * from �������������������������� where ���� >= '" + ������������� + "'";
                    fillDataSet.FillTable(query��������������������������, ds11, "��������������������������", con, transaction);

                    // ���� ���������� ���� �� �����, �� ���� ������� ����� � ���� ������� ���������, ��� �� ����� �� ������� ������ �� ���������� ������ (�� �� ������ ����� �� �����).
                    //// ��������� ��������� ������ �� �������� ������.
                    //builder.Append(query��������);
                    //builder.Append(query�������);
                    //builder.Append(query�����������������);
                    //builder.Append(query��������������������������);

                    //SqlCommand comDel = new SqlCommand(builder.ToString().Trim(), con);
                    //comDel.Transaction = transaction;

                    //comDel.ExecuteNonQuery();

                }

                dataView������������������� = new DataView(ds11.�������);
                dataView��������������������� = new DataView(ds11.�������);
                dataView������������������ = new DataView(ds11.��������������������������);
                this.Refresh();

                dataGrid����������������.DataSource = null;
                dataGrid��������������.DataSource = null;
                dataGrid������������������.DataSource = null;

                dataView�������������������.Table = ds11.�������;
                dataView�������������������.RowFilter = "�����=False AND ���������� >='01.12." + ������������ + "'";
                dataGrid����������������.DataSource = dataView�������������������;

                dataView���������������������.Table = ds11.�������;
                dataView���������������������.RowFilter = "�����=True AND ���������� >='01.12." + ������������ + "'";
                dataGrid��������������.DataSource = dataView���������������������;

                dataView������������������.Table = ds11.��������������������������;
                dataView������������������.RowFilter = "���� >='01.12." + ������������ + "'";
                dataGrid������������������.DataSource = dataView������������������;
                this.Refresh();

                ����������();
                this.Refresh();
            }
            catch (Exception exc)
            {
                MessageBox.Show("" + exc.InnerException + "\n" + exc.Message + "\n" + exc.Source);
                Dispose(true);
            }
            finally
            {
                �������������.Abort();
                this.Refresh();
            }

        }

        /// <summary>
        /// ������������ �������������� ������ �� ���� 
        /// � ����������� ����������� �� �������������� ������
        /// </summary>
        private void ����������()
        {
            // ��������
            //int ���������� = 0;
            string ���������� = "0";
            string ����������������������� = "0";
            string ��������������������������� = "0";
            string ���������������������������������� = "0";
            string ������������������������ = "0";

            int ����� = 0;
            int ��������������������� = 0;
            int ���������� = 0;

            ������������ sConnect = new ������������();
            string sConn = sConnect.�����������������();

            using (SqlConnection con = new SqlConnection(sConn))
            {
                con.Open();

                // ����� ����������.
                Statistic statistic = new Statistic(selectedYear);
                DataTable tab = statistic.���������������(con);

                ���������� = tab.Rows[0][0].ToString();

                // ����� �������� ����������.
                ����������������������� = statistic.�����������������������(con).Rows[0][0].ToString();

                // ����� ���������� ������������ �� ��������.
                ��������������������������� = statistic.�������������������������������������(con).Rows[0][0].ToString();

                // ����� ���������� ����������� ����������.
                ���������������������������������� = statistic.������������������������������������������������(con).Rows[0][0].ToString();

                // ����� ��������� ����������.
                ������������������������ = statistic.������������������������(con).Rows[0][0].ToString(); ;

            }
            

            //������� ���� ������ ���������� ����� ������������.


            //DataRow[] rows = ds11.�������.Select("���������� >='01.12." + ������������ + "'");
            //���������� = rows.Length;

            //rows = ds11.�������.Select("�����=True AND ���������� >='01.12." + ������������ + "'");
            //����� = rows.Length;

            //rows = ds11.�������.Select("�����=False AND ���������� >='01.12." + ������������ + "'");
            //��������������������� = rows.Length;

            //rows = ds11.�������.Select("����������=True AND ���������� >='01.12." + ������������ + "'");
            //���������� = rows.Length;

            //string ���� = "����� ���������� ���������� � ����: " + ���������� + "\n" +
            //                     "���������� ��������� � ����: " + ����� + "\n" +
            //                     "���������� ��������� ������������ : " + ��������������������� + "\n" +
            //                     "���������� ������� �� �������� : " + ����������;

            string ���� = "����� ���������� � ����: " + ���������� + "\n" +
                                 "����� �������� ����������: " + ����������������������� + "\n" +
                                 "����� ���������� ������������ �� �������� : " + ��������������������������� + "\n" +
                                 "����� ��������� ���������� ������������ �� �������� : " + ����������������������������������;


            label����Tab1.Text = ����;
            label����Tab2.Text = ����;

            // ���������
            //���������� = 0;
            //rows = ds11.��������������������������.Select("���� >='01.12." + ������������ + "'");
            //���������� = rows.Length;
            //label����Tab3.Text = "��������� ���������� � ����: " + ���������� + "\n";
            label����Tab3.Text = "��������� ���������� � ����: " + ������������������������ + "\n";
        }

        /// <summary>
        /// �������� ���� ���-�������� � ��������� ParameterFields ��� ������ Crystal Reports
        /// </summary>
        /// <param name="paramName">��� ���������</param>
        /// <param name="paramValue">string �������� ���������</param>
        /// <param name="paramFields">string ��������� ����������</param>
        public static void ������������������(string paramName,
            string paramValue,
            ParameterFields paramFields)
        {
            ParameterField paramField = new ParameterField();// ��������
            ParameterDiscreteValue paramDiscreteValue = new ParameterDiscreteValue();
            ParameterValues paramValues = new ParameterValues();
            // ������������� ��� ���������
            paramField.ParameterFieldName = paramName;// ��� ���������
            // ������������� �������� ���������
            paramDiscreteValue.Value = paramValue;
            paramValues.Add(paramDiscreteValue);
            paramField.CurrentValues = paramValues;
            // ��������� �������� � ���������� ���������
            paramFields.Add(paramField);
        }

        /// <summary>
        /// ���������� ����� � ������� "������ ����������� ��������"
        /// </summary>
        private void ��������������()
        {
            ������������� = new System.Threading.Thread(new System.Threading.ThreadStart(����������������������));
            �������������.Start();

            DataGrid datagrid = new DataGrid();

            if (dataGrid����������������.CanSelect)
            {
                datagrid = dataGrid����������������;
            }
            if (dataGrid��������������.CanSelect)
            {
                datagrid = dataGrid��������������;
            }

            if (datagrid.CurrentCell.RowNumber == -1)
            {
                return;
            }

            int id��������������� = this.ID���������������;

            FormView ����������� = new FormView();
            // ������� ����� �� �������:
            this.Enabled = false;

            try
            {
                // ReportDocument �������� �������� � ������ ��� �������� ������:
                ReportDocument rptDoc = new ReportDocument();
                // ��������� ���� ������:
                string fileName = @"..\report\Card.rpt";
                // ���� ������:
                rptDoc.Load(fileName);
                // �������� ������:
                rptDoc.SetDataSource(ds11);
                // ������������ ������� �������� ������ � ��������� � ����:
                �����������.reportViewer.ReportSource = rptDoc;
                // �������� ��������� � �����:
                ������������������("id_card", Convert.ToString(id���������������), �����������.reportViewer.ParameterFieldInfo);
                // ���������� �����:
                �������������.Abort();
                �����������.ShowDialog(this);
            }
            catch (System.IndexOutOfRangeException exc)
            {
                MessageBox.Show("��� ������� ��� ������. \n" + exc.StackTrace);
                return;
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, "��������� ������ ��� �������� ����� ������ \"������ �������� ���������\".\n" + exc.Message + "\n" + exc.InnerException, "������ �������� ����� ������");
                return;
            }
            finally
            {
                �������������.Abort();
                this.Enabled = true;
            }
        }

        /// <summary>
        /// ���������� ����� � ������� "��������� � �������� ������ ����������"
        /// </summary>
        private void ����������������������������()
        {
            FormView ����������� = new FormView();
            this.Enabled = false;
            try
            {
                // ReportDocument �������� �������� � ������ ��� �������� ������:
                ReportDocument rptDoc = new ReportDocument();
                // ��������� ���� ������:
                string fileName = @"..\report\ExpiredDoc.rpt";
                // ���� ������:
                rptDoc.Load(fileName);
                // �������� ������:
                rptDoc.SetDataSource(ds11);
                // ������������ ������� �������� ������ � ��������� � ����:
                �����������.reportViewer.ReportSource = rptDoc;
                // ���� ����������� ����������� �� ���� �����:
                �����������.WindowState = FormWindowState.Maximized;
                // ��������� ���� ��������
                �������������.Abort();
                // ���������� �����:
                �����������.ShowDialog(this);
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, "��������� ������ ��� �������� ����� ������ \"��������� � ��������� ������� ����������\".\n" + exc.Message, "������ �������� ����� ������");
                return;
            }
            finally
            {
                �������������.Abort();
                this.Enabled = true;
            }
        }

        /// <summary>
        /// ���������� � Word �������� ������������ ����������.
        /// </summary>
        /// <param name="tab">�������</param>
        /// <param name="tabP">�������������</param>
        //private void ����������������������������(DataTable tab)//
        private void ����������������������������(List<���������������������> list)
        {
            //������ ����� Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //app.Documents.Add(("�������������.doc");

            string filName = Environment.CurrentDirectory + @"\������\��������� � ��������� ������� ���������� ��.doc";

            //��������� ��������
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

            ////���� ������ ������.
            object wdrepl = Word.WdReplace.wdReplaceAll;
            //object searchtxt = "GreetingLine";
            object searchtxt = "date";
            object newtxt = (object)DateTime.Today.ToShortDateString();
            //object frwd = true;
            object frwd = false;
            doc.Content.Find.Execute(ref searchtxt, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd, ref missing, ref missing, ref newtxt, ref wdrepl, ref missing, ref missing,
            ref missing, ref missing);

            //�������� �������
            object bookNaziv = "�������";
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

            table.Borders.Enable = 1; // ����� - �������� �����
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 9;

            //������� ����� �������.
            table.Cell(1, 1).Range.Text = "� �/�";
            table.Cell(1, 2).Range.Text = "������������� �����������";
            table.Cell(1, 3).Range.Text = "���� �����������";
            table.Cell(1, 4).Range.Text = "����� ��������";
            table.Cell(1, 5).Range.Text = "���� ����������";

            Object beforeRow1 = Type.Missing;
            table.Rows.Add(ref beforeRow1);

            int count = 1;

            //�������� ������� �������.
            //foreach (DataRow row in tab.Rows)
            foreach(��������������������� item in list)
            {
                table.Cell(count + 1, 1).Range.Text = item.�������.Trim(); // count.ToString().Trim();
                table.Cell(count + 1, 2).Range.Text = item.������������������������.Trim(); // row["���������"].ToString().Trim();
                table.Cell(count + 1, 3).Range.Text = item.���������������.Trim(); //Convert.ToDateTime(row["����������"]).ToShortDateString();
                table.Cell(count + 1, 4).Range.Text = item.�������������.Trim();  //row["���������"].ToString().Trim();
                table.Cell(count + 1, 5).Range.Text = item.��������������.Trim(); //Convert.ToDateTime(row["��������������"]).ToShortDateString();

                Object beforeRow2 = Type.Missing;
                table.Rows.Add(ref beforeRow2);

                count++;
            }


            ////�������� ������� �������.
            //foreach (DataRow row in tabP.Rows)
            //{
            //    table.Cell(count + 1, 1).Range.Text = count.ToString().Trim();
            //    table.Cell(count + 1, 2).Range.Text = row["���������"].ToString().Trim();
            //    table.Cell(count + 1, 3).Range.Text = Convert.ToDateTime(row["����������"]).ToShortDateString();
            //    table.Cell(count + 1, 4).Range.Text = row["���������"].ToString().Trim();
            //    table.Cell(count + 1, 5).Range.Text = Convert.ToDateTime(row["��������������"]).ToShortDateString();

            //    Object beforeRow2 = Type.Missing;
            //    table.Rows.Add(ref beforeRow2);

            //    count++;
            //}

            //������ ��������� ������
            table.Rows[count + 1].Delete();

            // ���� ������.
            object wdrepl2 = Word.WdReplace.wdReplaceAll;
            //object searchtxt = "GreetingLine";
            object searchtxt2 = "countdoc";
            object newtxt2 = (object)list.Count;
            //object frwd = true;
            object frwd2 = false;
            doc.Content.Find.Execute(ref searchtxt2, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd2, ref missing, ref missing, ref newtxt2, ref wdrepl2, ref missing, ref missing,
            ref missing, ref missing);

           // ���������� �������� � ������� ����.
           app.Visible = true;

           //doc = Application.Documents["�������������.doc"] as Word._Document;
           //doc.Close(ref doNotSaveChanges, ref missing, ref missing);
        }

        /// <summary>
        /// ���������� ����� � ������� "����������� �����������"
        /// </summary>
        private void ����������������������������()
        {
            //������������� = new System.Threading.Thread(new System.Threading.ThreadStart(����������������������));
            //�������������.Start();

            // ������� ��� ����������� �����������:
            DS���������������������� ds����������� = new DS����������������������();

            // ���������� ����:
            DateTime ������ = DateTime.Now.AddDays(1);

            // ������ ����������.
            // �������� ���� ������������ ������� ��������� � ���������� ������ ���������� ������ ��� ������
            //DataRow[] ����������������� = ds11.�������.Select("��������������<='" + ������.Date + "' AND ���������� >='01.12.2011' AND ����������=True AND �����=False");

            ������� ������� = new �������();
            //DataRow[] ����������������� = �������.���������������������������������().Select("��������������<='" + ������.Date + "' AND ���������� >='01.12.2011' AND ����������=True AND �����=False");
            DataRow[] ����������������� = �������.�����������������������();//.Select("���������� >='01.12.2017' AND ����������=True AND �����=False");
           

            // ������ ��� ����� �����������
            System.Collections.ArrayList ����������������� = new ArrayList();

            // ��������� ������ �������
            foreach (DataRow row in �����������������)
            {
                //string ���������� = (string)row["���������"];
                string ���������� = (string)row["������������������"];
                string[] ���������� = ����������.Split(',');
                foreach (string ��� in ����������)
                {
                    if (!�����������������.Contains(���.Trim()))
                    {
                        �����������������.Add(���.Trim());
                    }
                }
            }

            int CountDocumentControl = 0;

            StatisticControlNotific statistic = new StatisticControlNotific();

            // ������� ����� ���������� �� ��������.
            statistic.������������������������ = �������.���������������������������������().Select("����������=True AND �����=False").Length;

            // ������� ����� ���������� �� �������� � �������� ������.
            
            statistic.�������������������������������� = �������.���������������������������������().Select("��������������<'" + DateTime.Now.Date + "' AND ����������=True AND �����=False").Length;

           
            foreach (Object ��� in �����������������)
            {
                // ������� ��������������� ���������������������.
                PersonDocument pd = new PersonDocument();

                pd.FioPerson = ���.ToString().Trim();

                // ���������� ���������� �� ��������.
                pd.������������������������ = �������.���������������������������(���.ToString().Trim()).Tables[0].Rows.Count;

                // ���������� ������������ ����������.
                pd.��������������������� = �������.����������������������������(���.ToString().Trim()).Tables[0].Select();

                // ���������� ������������ ����������.
                pd.�������������������������������� = pd.���������������������.Length;

                // ���������� �� ������������ ����������.
                pd.����������������������� = �������.������������������������������(���.ToString().Trim()).Tables[0].Select();

                pd.���������������������������������� = pd.�����������������������.Length;

                // ������� � ������������� �����������.
                //DataRow[] dtOverDoc = �������.���������������������������������().Rows;//���.ToString().Trim());//.Select("��������������<'" + DateTime.Now.Date + "' AND ������������������ LIKE '%" + ��� + "%' AND ����������=True AND �����=False");

                //// ������� ������ ������������ ����������.
                //pd.��������������������� = dtOverDoc;

                //// ����� ���������� �� ��������.
                //pd.������������������������ = �������.���������������������������������().Select("������������������ LIKE '%" + ��� + "%' AND ����������=True AND �����=False").Length;

                //// ������ ���������� �� ��������.
                //pd.������������������� = �������.���������������������������������().Select("������������������ LIKE '%" + ��� + "%' AND ����������=True AND �����=False");

                //// ���������� ������������ ���������� ��� �������� ������������.
                //pd.�������������������������������� = dtOverDoc.Length;

               

                //// ������� � �� ������������� �����������.
                ////DataRow[] dtNotOverDoc = �������.���������������������������������().Select("�������������� >='" + ������.Date + "' and ��������������<'" + DateTime.Now.Date + "' AND ������������������ LIKE '%" + ��� + "%' AND ����������=True AND �����=False");

                //DataRow[] dtNotOverDoc = �������.���������������������������������().Select("�������������� >'" + DateTime.Now.Date + "' AND ������������������ LIKE '%" + ��� + "%' AND ����������=True AND �����=False");

                //// ������� �� ������������ ���������.
                //pd.����������������������� = dtNotOverDoc;

                //// ������� ���������� �� ������������ ����������.
                //pd.���������������������������������� = dtNotOverDoc.Length;

                statistic.������������������.Add(pd);
            }

            string iTest = "";

            /*
            foreach (Object ��� in �����������������)
            {

                // ��������� �� �������� ��� �������� ����������:
                //DataRow[] ����� = ds11.�������.Select("��������� LIKE '%" + ��� + "%' AND ����������=True AND �����=False");
                DataRow[] ����� = �������.���������������������������������().Select("������������������ LIKE '%" + ��� + "%' AND ����������=True AND �����=False");

                CountDocumentControl += �����.Length;

                // ������������ ��������� ��� �������� ����������:
                //DataRow[] ������������ = ds11.�������.Select("��������������<'" + DateTime.Now.Date + "' AND ��������� LIKE '%" + ��� + "%' AND ����������=True AND �����=False");
                DataRow[] ������������ = �������.���������������������������������().Select("��������������<'" + DateTime.Now.Date + "' AND ������������������ LIKE '%" + ��� + "%' AND ����������=True AND �����=False");
                
                
                // ��������� � 1 ���� ��� �������� ����������:
                //DataRow[] �1��� = ds11.�������.Select("��������������='" + ������.Date + "' AND ��������� LIKE '%" + ��� + "%' AND ����������=True AND �����=False");
                DataRow[] �1��� = �������.���������������������������������().Select("��������������='" + ������.Date + "' AND ������������������ LIKE '%" + ��� + "%' AND ����������=True AND �����=False");

                // ��������� ������ � ������� "����������":
                ds�����������.����������.Add����������Row(���.ToString(), �����.Length, �1���.Length, ������������.Length);

                // ������ �� �������� ����������:
                DataRow[] ���������� = ds�����������.����������.Select("������������������='" + ���.ToString() + "'");
                int id���������� = (int)����������[0]["id_����������"];

                // ���������� ����� ��������� ��� �������� ����������:
                int ������� = 1;

                // ��� ���������:
                int ��� = 0;// �1���� = 0, ������������ = 1

                //��������� ��������� � ����� ���� � ������� "���������"
                foreach (DataRow �������� in �1���)
                {
                    DateTime ��������������� = (DateTime)��������["����������"];
                    string ������������� = (string)��������["���������"];
                    DateTime ������������ = (DateTime)��������["��������������"];
                    ds�����������.���������.Add���������Row(
                                                            id����������,
                                                            ���������������,
                                                            �������������,
                                                            ������������,
                                                            �������,
                                                            ���
                                                            );
                    �������++;
                }

                //��������� ��������� � ����� ���� � ������� "���������"
                ������� = 1;
                ��� = 1;
                foreach (DataRow �������� in ������������)
                {
                    DateTime ��������������� = (DateTime)��������["����������"];
                    string ������������� = (string)��������["���������"];
                    DateTime ������������ = (DateTime)��������["��������������"];
                    ds�����������.���������.Add���������Row(id����������,
                                                            ���������������,
                                                            �������������,
                                                            ������������,
                                                            �������,
                                                            ���);
                    �������++;
                }
            }


            */

            FormPrint���������������������� formPrint = new FormPrint����������������������();
            
            // ��������� ��������� ������.
            //formPrint.DataSetForm = ds�����������;
            formPrint.DataStatistic = statistic;

            // ��������� ���������� ���������� �� ��������.
            formPrint.CountDocControl = CountDocumentControl;
            formPrint.Show();


            //string sTest = "";

            //Form���������������������� ����������� = new Form����������������������(ds�����������);
            //// ��������� ��. �����:
            //this.Enabled = false;

            //try
            //{
            //    // ReportDocument �������� �������� � ������ ��� �������� ������:
            //    ReportDocument rptDoc = new ReportDocument();
            //    // ��������� ���� ������:
            //    string fileName = @"..\report\KontrolMessage.rpt";
            //    // ���� ������:
            //    rptDoc.Load(fileName);
            //    // �������� ������:
            //    rptDoc.SetDataSource(ds�����������);
            //    // ������������ ������� �������� ������ � ��������� � ����:
            //    �����������.reportViewer.ReportSource = rptDoc;
            //    // ��������� ���� ��������:
            //    �������������.Abort();
            //    // ���������� �����:
            //    �����������.ShowDialog(this);
            //}
            //catch (System.Exception exc)
            //{
            //    �������������.Abort();
            //    MessageBox.Show("������ ������ \"����������� �����������\". \n" + exc.Message + "\n" + exc.StackTrace);
            //    return;
            //}
            //finally
            //{
            //    �������������.Abort();
            //    this.Enabled = true;
            //}
        }

        private void ����������������������()
        {
            Form�������� form = new Form��������();
            form.Left = (this.Left) + this.Width / 2 - (form.Width / 2);
            form.Top = (this.Top) + this.Height / 2 - (form.Height / 2);
            form.TopMost = true;
            form.ShowDialog();
        }

        /// <summary>
        /// ���������, �������� ����� ��������� ��� ���
        /// </summary>
        /// <returns>true ���� ��������� ��������, ����� false</returns>
        static bool ��������������������()
        {
            bool createdNew;
            mutex = new System.Threading.Mutex(false, "RegKorMutex", out createdNew);
            return !createdNew;
        }


        #endregion

        #region ��������
        private int ID���������������
        {
            get
            {

                DataGrid datagrid = new DataGrid();

                if (dataGrid����������������.CanSelect)
                {
                    datagrid = dataGrid����������������;
                }
                if (dataGrid��������������.CanSelect)
                {
                    datagrid = dataGrid��������������;
                }
                if (dataGrid������������������.CanSelect)
                {
                    datagrid = dataGrid������������������;
                }
                // �������� ������ ������������ � ���������� ������:
                BindingManagerBase bmb = this.BindingContext[datagrid.DataSource, datagrid.DataMember];
                bmb.Position = datagrid.CurrentCell.RowNumber;
                datagrid.Select(datagrid.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                return (int)drv["id_��������"];
            }
        }

        #endregion

        #region �������

        /// <summary>
        /// ������� ������ � ������ ������� "����������������"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid����������������_DoubleClick(object sender, EventArgs e)
        {
            //// �������� ������� � ������� �������� ������:
            //DataGrid.HitTestInfo myHitTest = dataGrid����������������.HitTest(����.X, ����.Y);

            //if (myHitTest.Type == DataGrid.HitTestType.Cell)// ���� �������� � ������, ��� ������� ����� �������
            //{
            //    Form�������� form = new Form��������(ds11, ID���������������, ������������);
            //    form.ShowDialog(this);
            //    if (form.DialogResult == DialogResult.OK)
            //    {
            //        DS1TableAdapters.��������TableAdapter ������� = new RegKor.DS1TableAdapters.��������TableAdapter();
            //        �������.Update(form.��������������);
            //        ��������������();
            //    }
            //}

        }

        /// <summary>
        /// ������� ������ � ������� "��������������"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid��������������_DoubleClick(object sender, EventArgs e)
        {
            //// �������� ������� � ������� �������� ������:
            //DataGrid.HitTestInfo myHitTest = dataGrid��������������.HitTest(����.X, ����.Y);

            //if (myHitTest.Type == DataGrid.HitTestType.Cell)// ���� �������� � ������, ��� ������� ����� �������
            //{
            //    Form�������� form = new Form��������(ds11, ID���������������, ������������);
            //    form.ShowDialog(this);
            //    if (form.DialogResult == DialogResult.OK)
            //    {
            //        DS1TableAdapters.��������TableAdapter ������� = new RegKor.DS1TableAdapters.��������TableAdapter();
            //        �������.Update(form.��������������);
            //        ��������������();
            //    }
            //}

        }

        /// <summary>
        /// ������� ������ � ������� "������������������"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid������������������_DoubleClick(object sender, EventArgs e)
        {
            /*
            // �������� ������� � ������� �������� ������:
            DataGrid.HitTestInfo myHitTest = dataGrid������������������.HitTest(����.X, ����.Y);

            if (myHitTest.Type == DataGrid.HitTestType.Cell)// ���� �������� � ������, ��� ������� ����� �������
            {
                ������������� = new System.Threading.Thread(new System.Threading.ThreadStart(����������������������));
                �������������.Start();
                // �������� ������ ������������ � ���������� ������:
                BindingManagerBase bmb = this.BindingContext[dataGrid������������������.DataSource, dataGrid������������������.DataMember];
                bmb.Position = dataGrid������������������.CurrentCell.RowNumber;
                dataGrid������������������.Select(dataGrid������������������.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                DataRow[] row = ds11.�����������������.Select("id_��������=" + (int)drv["id_��������"]);
                DS1.�����������������Row ������������������ = (DS1.�����������������Row)row[0];
                Form����������������� form = new Form�����������������(ds11, ������������������, ������������);
                �������������.Abort();
                form.ShowDialog(this);
                if (form.DialogResult == DialogResult.OK)
                {
                    string sTest = "asd";

                    DS1.�����������������Row rowTest = form.�����������������������;


                    DS1TableAdapters.�����������������TableAdapter ������� = new RegKor.DS1TableAdapters.�����������������TableAdapter();
                    �������.Update(form.�����������������������);
                    ��������������();
                }
            }*/

        }

        /// <summary>
        /// ������ ���� � ������� ����������������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid����������������_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            ����.X = e.X;
            ����.Y = e.Y;

            // �������� ������� � ������� �������� ������:
            DataGrid.HitTestInfo myHitTest = dataGrid����������������.HitTest(e.X, e.Y);

            if (e.Button == System.Windows.Forms.MouseButtons.Right)// ���� �������� ������ �������
            {
                // ������� ��� ������ ������������ ����
                contextMenu1.MenuItems.Clear();

                // ������� ����� ���� �������� ����������:
                MenuItem menuItemContext�������������� = new System.Windows.Forms.MenuItem("�������� ������");
                menuItemContext��������������.Click += new EventHandler(menuItemContext��������������_Click);
                this.contextMenu1.MenuItems.Add(0, menuItemContext��������������);

                if (myHitTest.Type == DataGrid.HitTestType.Cell)// ���� �������� � ������
                {
                    // ������� ������� ���������:
                    dataGrid����������������.UnSelect(dataGrid����������������.CurrentRowIndex);
                    // ������ ������� �������:
                    dataGrid����������������.CurrentCell = new DataGridCell(myHitTest.Row, myHitTest.Column);
                    // �������� ��� ������:
                    dataGrid����������������.Select(myHitTest.Row);

                    MenuItem menuItemContext�������������� = new System.Windows.Forms.MenuItem("�������� ������");
                    menuItemContext��������������.Click += new EventHandler(menuItemContext��������������_Click);
                    this.contextMenu1.MenuItems.Add(1, menuItemContext��������������);

                    MenuItem menuItemContext������������� = new System.Windows.Forms.MenuItem("������� ������");
                    menuItemContext�������������.Click += new EventHandler(menuItemContext�������������_Click);
                    this.contextMenu1.MenuItems.Add(2, menuItemContext�������������);

                    MenuItem menuItemContext�������������� = new System.Windows.Forms.MenuItem("������ ��������");
                    menuItemContext��������������.Click += new EventHandler(menuItemContext��������������_Click);
                    this.contextMenu1.MenuItems.Add(3, menuItemContext��������������);

                }

                // ���������� ��������� ����������� ����:
                contextMenu1.Show(dataGrid����������������, new System.Drawing.Point(e.X, e.Y));
            }

            if (myHitTest.Type == DataGrid.HitTestType.Cell)// ���� �������� � ������, ��� ������� ����� �������
            {
                // �������� ������ ������������ � ���������� ������:
                BindingManagerBase bmb = this.BindingContext[dataGrid����������������.DataSource, dataGrid����������������.DataMember];
                bmb.Position = dataGrid����������������.CurrentCell.RowNumber;
                dataGrid����������������.Select(dataGrid����������������.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                // ������� ���������� ������ �� �������������� �����:�������������������
                label����Tab1.Text = "��������: " + drv["�����������������"].ToString() + Environment.NewLine +
                    "����-�: " + drv["����������������������"].ToString() + Environment.NewLine +
                    "���� ����.: " + Convert.ToDateTime(drv["����������"]).ToShortDateString() +
                    "  ������.: " + drv["����������"].ToString() + Environment.NewLine +
                    "���� ����.: " + Convert.ToDateTime(drv["����������"]).ToShortDateString() +
                    "  �����.: " + drv["���������"].ToString() + Environment.NewLine +
                    "����������: " + drv["�����������������"].ToString() + Environment.NewLine +
                    "���� ��������: " + drv["���������"].ToString();
            }
            else if (dataGrid����������������.CurrentRowIndex > -1)
            {
                // ������� ������� ���������:
                dataGrid����������������.UnSelect(dataGrid����������������.CurrentRowIndex);
                ����������();
            }
        }

        /// <summary>
        /// ������ ���� � ������� "�������� ������� ���������"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid��������������_MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            ����.X = e.X;
            ����.Y = e.Y;

            // �������� ������� � ������� �������� ������:
            DataGrid.HitTestInfo myHitTest = dataGrid��������������.HitTest(e.X, e.Y);

            if (e.Button == System.Windows.Forms.MouseButtons.Right)// ���� �������� ������ �������
            {
                // ������� ��� ������ ������������ ����
                contextMenu1.MenuItems.Clear();

                if (myHitTest.Type == DataGrid.HitTestType.Cell)// ���� �������� � ������
                {
                    // ������� ������� ���������:
                    dataGrid��������������.UnSelect(dataGrid��������������.CurrentRowIndex);
                    // ������ ������� �������:
                    dataGrid��������������.CurrentCell = new DataGridCell(myHitTest.Row, myHitTest.Column);
                    // �������� ��� ������:
                    dataGrid��������������.Select(myHitTest.Row);

                    MenuItem menuItemContext�������������� = new System.Windows.Forms.MenuItem("�������� ������");
                    menuItemContext��������������.Click += new EventHandler(menuItemContext��������������_Click2);
                    this.contextMenu1.MenuItems.Add(0, menuItemContext��������������);

                    MenuItem menuItemContext������������� = new System.Windows.Forms.MenuItem("������� ������");
                    menuItemContext�������������.Click += new EventHandler(menuItemContext�������������_Click2);
                    this.contextMenu1.MenuItems.Add(1, menuItemContext�������������);

                    MenuItem menuItemContext�������������� = new System.Windows.Forms.MenuItem("������ ��������");
                    menuItemContext��������������.Click += new EventHandler(menuItemContext��������������_Click);
                    this.contextMenu1.MenuItems.Add(2, menuItemContext��������������);

                    MenuItem menuItemContext������ = new MenuItem("������������� �� �������");
                    menuItemContext������.Click += new EventHandler(menuItemContext������_Click);
                    this.contextMenu1.MenuItems.Add(3, menuItemContext������);

                    //MenuItem menuItemContext�������������������������� = new MenuItem("������������� �� �������");
                    //menuItemContext��������������������������.Click += new EventHandler(menuItemContext��������������������������_Click);
                    //this.contextMenu1.MenuItems.Add(3, menuItemContext��������������������������);


                }

                // ���������� ��������� ����������� ����:
                contextMenu1.Show(dataGrid��������������, new System.Drawing.Point(e.X, e.Y));
            }

            if (myHitTest.Type == DataGrid.HitTestType.Cell)// ���� �������� � ������, ��� ������� ����� �������
            {
                // �������� ������ ������������ � ���������� ������:
                BindingManagerBase bmb = this.BindingContext[dataGrid��������������.DataSource, dataGrid��������������.DataMember];
                bmb.Position = dataGrid��������������.CurrentCell.RowNumber;
                dataGrid��������������.Select(dataGrid��������������.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                // ������� ���������� ������ �� �������������� �����:�������������������
                label����Tab2.Text = "��������: " + drv["�����������������"].ToString() + Environment.NewLine +
                                    "����-�: " + drv["����������������������"].ToString() + Environment.NewLine +
                                    "���� ����.: " + Convert.ToDateTime(drv["����������"]).ToShortDateString() +
                                    "  ������.: " + drv["����������"].ToString() + Environment.NewLine +
                                    "���� ����.: " + Convert.ToDateTime(drv["����������"]).ToShortDateString() +
                                    "  �����.: " + drv["���������"].ToString() + Environment.NewLine +
                                    "����������: " + drv["�����������������"].ToString() + Environment.NewLine +
                                    "���� ��������: " + drv["���������"].ToString() + Environment.NewLine +
                                    "��������� ����������: " + drv["�������������������"].ToString();
            }
            else if (dataGrid��������������.CurrentRowIndex > -1)
            {
                // ������� ������� ���������:
                dataGrid��������������.UnSelect(dataGrid��������������.CurrentRowIndex);
                ����������();
            }
        }

        void menuItemContext��������������������������_Click(object sender, EventArgs e)
        {
           //// ������� id ������� ��������.
           //int idCard = this.DataGr
        }

        void menuItemContext������_Click(object sender, EventArgs e)
        {
            FormSelectMonth fsm = new FormSelectMonth();
            fsm.������������ = ������������;
            fsm.ShowDialog();

            if (fsm.DialogResult == DialogResult.OK)
            {
                // ������� ������ � ��������� ���� ������.
                string ���������� = fsm.Get����������;
                string ��������� = fsm.Get�����������;

                if (���������� != null)
                {

                    dataGrid��������������.DataSource = null;
                    //string ������ = "�����=True AND ���������� >='01.12.2011' AND (����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    //    " OR ����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    //    " OR ���������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    //    " OR ������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    //    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    //    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    //    " OR ���������� LIKE '%" + textBox������������Tab2.Text + "%')";

                    //========================

                    //������� ���������� - ������
                    string ������;
                    if (this.comboBox��������������.Visible == true)
                    {
                        ������ = "�����=True AND (���������� >='01.12.2011' AND ���������� >='" + ���������� + "' AND ���������� <='" + ��������� + "') AND (����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ���������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ���������� LIKE '%" + textBox������������Tab2.Text + "%')" +
                            " AND ���������������������� = '" + this.comboBox��������������.Text + "'";
                        //" AND ���������������������� = '��� �� \"���� �.��������\"'";
                    }
                    else
                    {
                        ������ = "�����=True AND (���������� >='01.12.2011' AND ���������� >='" + ���������� + "' AND ���������� <='" + ��������� + "') AND (����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ���������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ���������� LIKE '%" + textBox������������Tab2.Text + "%')";
                    }

                    if (this.textBox������������Tab2.Text == "")
                    {
                        ������ = "�����=True AND (���������� >='01.12.2011' AND ���������� >='" + ���������� + "' AND ���������� <='" + ��������� + "') AND (����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ���������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                            " OR ���������� LIKE '%" + textBox������������Tab2.Text + "%')";
                    }
                    //=====================
                    dataView���������������������.RowFilter = ������;
                    dataGrid��������������.DataSource = dataView���������������������;
                    textBox������������Tab2.Focus();
                    if (textBox������������Tab2.Text != "")
                    {
                        label�������������������������Tab2.Text = "�������� ����������: " + dataView���������������������.Count;
                    }
                    else
                    {
                        label�������������������������Tab2.Text = "";
                    }

                }

            }

        }

        /// <summary>
        /// ������ ���� � ������� "��������� ���������"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid������������������_MouseUp(object sender, MouseEventArgs e)
        {
            ����.X = e.X;
            ����.Y = e.Y;

            // �������� ������� � ������� �������� ������:
            DataGrid.HitTestInfo myHitTest = dataGrid������������������.HitTest(e.X, e.Y);

            if (e.Button == System.Windows.Forms.MouseButtons.Right)// ���� �������� ������ �������
            {
                // ������� ��� ������ ������������ ����
                contextMenu1.MenuItems.Clear();

                // ������� ����� ���� �������� ����������:
                MenuItem menuItemContext�������������� = new System.Windows.Forms.MenuItem("�������� ������");
                menuItemContext��������������.Click += new EventHandler(menuItemContext�����������������������_Click);
                this.contextMenu1.MenuItems.Add(0, menuItemContext��������������);

                if (myHitTest.Type == DataGrid.HitTestType.Cell)// ���� �������� � ������
                {

                    // ������� ������� ���������:
                    dataGrid������������������.UnSelect(dataGrid������������������.CurrentRowIndex);
                    // ������ ������� �������:
                    dataGrid������������������.CurrentCell = new DataGridCell(myHitTest.Row, myHitTest.Column);
                    // �������� ��� ������:
                    dataGrid������������������.Select(myHitTest.Row);

                    MenuItem menuItemContext�������������� = new System.Windows.Forms.MenuItem("�������� ������");
                    menuItemContext��������������.Click += new EventHandler(menuItemContext�����������������������_Click);
                    this.contextMenu1.MenuItems.Add(1, menuItemContext��������������);

                    MenuItem menuItemContext������������� = new System.Windows.Forms.MenuItem("������� ������");
                    menuItemContext�������������.Click += new EventHandler(menuItemContext����������������������_Click);
                    this.contextMenu1.MenuItems.Add(2, menuItemContext�������������);

                }

                // ���������� ��������� ����������� ����:
                contextMenu1.Show(dataGrid������������������, new System.Drawing.Point(e.X, e.Y));
            }

            if (myHitTest.Type == DataGrid.HitTestType.Cell)// ���� �������� � ������, ��� ������� ����� �������
            {
                // ������� ������� ���������:
                dataGrid������������������.UnSelect(dataGrid������������������.CurrentRowIndex);
                // ������ ������� �������:
                dataGrid������������������.CurrentCell = new DataGridCell(myHitTest.Row, myHitTest.Column);
                // �������� ��� ������:
                dataGrid������������������.Select(myHitTest.Row);

                // �������� ������ ������������ � ���������� ������:
                BindingManagerBase bmb = this.BindingContext[dataGrid������������������.DataSource, dataGrid����������������.DataMember];
                bmb.Position = dataGrid������������������.CurrentCell.RowNumber;
                dataGrid������������������.Select(dataGrid������������������.CurrentCell.RowNumber);
                DataRowView drv = (DataRowView)bmb.Current;
                // ������� ���������� ������ �� �������������� �����:�������������������
                label����Tab3.Text = "����: " + Convert.ToDateTime(drv["����"]).ToShortDateString() + Environment.NewLine +
                    "�����: " + drv["��������������"].ToString() + Environment.NewLine +
                    "�������: " + drv["����������������"].ToString() + Environment.NewLine +
                    "�����������: " + drv["���������������������"].ToString() + Environment.NewLine +
                    "����������: " + drv["����������"].ToString() + Environment.NewLine +
                    "����� �� ��������: " + drv["���������������������������"].ToString() + Environment.NewLine;
            }
            else if (dataGrid������������������.CurrentRowIndex > -1)
            {
                // ������� ������� ���������:
                dataGrid������������������.UnSelect(dataGrid������������������.CurrentRowIndex);
                ����������();
            }
        }

        /// <summary>
        /// ��������� ��������� � ������� ������� ���������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid����������������_CurrentCellChanged(object sender, System.EventArgs e)
        {
            dataGrid����������������.Select(dataGrid����������������.CurrentCell.RowNumber);
            // �������� ������ ������������ � ���������� ������:
            BindingManagerBase bmb = this.BindingContext[dataGrid����������������.DataSource, dataGrid����������������.DataMember];
            bmb.Position = dataGrid����������������.CurrentCell.RowNumber;
            dataGrid����������������.Select(dataGrid����������������.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            // ������� ���������� ������ �� �������������� �����:�������������������
            label����Tab1.Text = "��������: " + drv["�����������������"].ToString() + Environment.NewLine +
                "����-�: " + drv["����������������������"].ToString() + Environment.NewLine +
                "���� ����.: " + Convert.ToDateTime(drv["����������"]).ToShortDateString() +
                "  ������.: " + drv["����������"].ToString() + Environment.NewLine +
                "���� ����.: " + Convert.ToDateTime(drv["����������"]).ToShortDateString() +
                "  �����.: " + drv["���������"].ToString() + Environment.NewLine +
                "����������: " + drv["�����������������"].ToString() + Environment.NewLine +
                "���� ��������: " + drv["���������"].ToString();
        }

        /// <summary>
        /// ��������� ��������� � ������� ��������� "� ����"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid��������������_CurrentCellChanged(object sender, System.EventArgs e)
        {
            dataGrid��������������.Select(dataGrid��������������.CurrentCell.RowNumber);
            // �������� ������ ������������ � ���������� ������:
            BindingManagerBase bmb = this.BindingContext[dataGrid��������������.DataSource, dataGrid��������������.DataMember];
            bmb.Position = dataGrid��������������.CurrentCell.RowNumber;
            dataGrid��������������.Select(dataGrid��������������.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            // ������� ���������� ������ �� �������������� �����:�������������������
            label����Tab2.Text = "��������: " + drv["�����������������"].ToString() + Environment.NewLine +
                "����-�: " + drv["����������������������"].ToString() + Environment.NewLine +
                "���� ����.: " + Convert.ToDateTime(drv["����������"]).ToShortDateString() +
                "  ������.: " + drv["����������"].ToString() + Environment.NewLine +
                "���� ����.: " + Convert.ToDateTime(drv["����������"]).ToShortDateString() +
                "  �����.: " + drv["���������"].ToString() + Environment.NewLine +
                "����������: " + drv["�����������������"].ToString() + Environment.NewLine +
                "���� ��������: " + drv["���������"].ToString() + Environment.NewLine +
                "��������� ����������: " + drv["�������������������"].ToString();
        }

        /// <summary>
        /// ��������� ��������� � ������� ��������� ���������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dataGrid������������������_CurrentCellChanged(object sender, EventArgs e)
        {
            dataGrid������������������.Select(dataGrid������������������.CurrentCell.RowNumber);

            // �������� ������ ������������ � ���������� ������:
            BindingManagerBase bmb = this.BindingContext[dataGrid������������������.DataSource, dataGrid����������������.DataMember];
            bmb.Position = dataGrid������������������.CurrentCell.RowNumber;
            dataGrid������������������.Select(dataGrid������������������.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            // ������� ���������� ������ �� �������������� �����:�������������������
            label����Tab3.Text = "����: " + Convert.ToDateTime(drv["����"]).ToShortDateString() + Environment.NewLine +
                "�����: " + drv["��������������"].ToString() + Environment.NewLine +
                "�������: " + drv["����������������"].ToString() + Environment.NewLine +
                "�����������: " + drv["���������������������"].ToString() + Environment.NewLine +
                "����������: " + drv["����������"].ToString() + Environment.NewLine +
                "����� �� ��������: " + drv["���������������������������"].ToString() + Environment.NewLine;
        }

        private void dataGrid����������������_Leave(object sender, System.EventArgs e)
        {
            if (label����Tab1.Focused)
            {
                return;
            }
            ����������();
        }

        private void dataGrid��������������_Leave(object sender, System.EventArgs e)
        {
            if (label����Tab2.Focused)
            {
                return;
            }
            ����������();
        }


        private void dataGrid������������������_Leave(object sender, EventArgs e)
        {
            if (label����Tab3.Focused)
            {
                return;
            }
            ����������();
        }

        private void checkBoxKontrolFilter_CheckedChanged(object sender, System.EventArgs e)
        {
            if (checkBoxKontrolFilter.Checked)
            {
                textBox������������Tab1.Text = "";
                textBox������������Tab1.Enabled = false;
                dataView�������������������.RowFilter = "�����=False AND ����������=True AND ���������� >='01.12.2011'";
                dataGrid����������������.DataSource = dataView�������������������;
            }
            else
            {
                textBox������������Tab1.Text = "";
                textBox������������Tab1.Enabled = true;
                dataView�������������������.RowFilter = "�����=False AND ���������� >='01.12.2011'";
                dataGrid����������������.DataSource = dataView�������������������;
            }
            ����������();

        }

        /// <summary>
        /// ��������� ������ � textBox������������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox������������_TextChanged(object sender, System.EventArgs e)
        {
            dataGrid����������������.DataSource = null;
            string ������ = "�����=False AND ���������� >='01.12.2011' AND (����������������� LIKE '%" + textBox������������Tab1.Text + "%'" +
                            " OR ����������������� LIKE '%" + textBox������������Tab1.Text + "%'" +
                            " OR ���������������������� LIKE '%" + textBox������������Tab1.Text + "%'" +
                            " OR ��������� LIKE '%" + textBox������������Tab1.Text + "%'" +
                            " OR ��������� LIKE '%" + textBox������������Tab1.Text + "%'" +
                            " OR ���������� LIKE '%" + textBox������������Tab1.Text + "%')";
            dataView�������������������.RowFilter = ������;
            dataGrid����������������.DataSource = dataView�������������������;
            textBox������������Tab1.Focus();
            if (textBox������������Tab1.Text != "")
            {
                label�������������������������Tab1.Text = "�������� ����������: " + dataView�������������������.Count;
            }
            else
            {
                label�������������������������Tab1.Text = "";
            }
        }

        /// <summary>
        /// ��������� ������ � ������ ������ �� ���������� "� ����"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox������������Tab2_TextChanged(object sender, System.EventArgs e)
        {
            dataGrid��������������.DataSource = null;
            //string ������ = "�����=True AND ���������� >='01.12.2011' AND (����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
            //    " OR ����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
            //    " OR ���������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
            //    " OR ������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
            //    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
            //    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
            //    " OR ���������� LIKE '%" + textBox������������Tab2.Text + "%')";

            //========================

            //������� ���������� - ������
            string ������;
            if (this.comboBox��������������.Visible == true)
            {
                ������ = "�����=True AND ���������� >='01.12.2011' AND (����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ���������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ���������� LIKE '%" + textBox������������Tab2.Text + "%')" +
                    " AND ���������������������� = '" + this.comboBox��������������.Text + "'";
                //" AND ���������������������� = '��� �� \"���� �.��������\"'";
            }
            else
            {
                ������ = "�����=True AND ���������� >='01.12.2011' AND (����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ���������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ���������� LIKE '%" + textBox������������Tab2.Text + "%')";
            }
            
            if(this.textBox������������Tab2.Text == "")
            {
                ������ = "�����=True AND ���������� >='01.12.2011' AND (����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ����������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ���������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ������������������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ��������� LIKE '%" + textBox������������Tab2.Text + "%'" +
                    " OR ���������� LIKE '%" + textBox������������Tab2.Text + "%')";
            }
            //=====================
            dataView���������������������.RowFilter = ������;
            dataGrid��������������.DataSource = dataView���������������������;
            textBox������������Tab2.Focus();
            if (textBox������������Tab2.Text != "")
            {
                label�������������������������Tab2.Text = "�������� ����������: " + dataView���������������������.Count;
            }
            else
            {
                label�������������������������Tab2.Text = "";
            }
        }

        private void textBox�������������������������������_TextChanged(object sender, EventArgs e)
        {
            DataView view = (DataView)dataGrid������������������.DataSource;
            view.RowFilter = ��������;
            textBox�������������������������������.Focus();
            label��������������������������������������������.Text = "�������� ����������: " + view.Count;
        }


        /// <summary>
        /// ������� ���������� ������ �� ������� ����������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button��������������������Tab1_Click(object sender, System.EventArgs e)
        {
            checkBoxKontrolFilter.Checked = false;
            textBox������������Tab1.Text = "";
            ����������();
        }

        /// <summary>
        /// ���������� ������� ������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button��������������������_Click(object sender, System.EventArgs e)
        {
            textBox������������Tab1.Text = "";
        }

        /// <summary>
        /// �������� ����� �� ���������� "� ����"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button��������������������Tab2_Click(object sender, System.EventArgs e)
        {
            textBox������������Tab2.Text = "";
        }

        private void button��������������������Tab2_Click_1(object sender, System.EventArgs e)
        {
            textBox������������Tab2.Text = "";
        }

        /// <summary>
        /// ������� ������ ������ ��������� ����������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button���������������������������������������_Click(object sender, EventArgs e)
        {
            textBox�������������������������������.Text = "";
            comboBox��������������.SelectedItem = "���� ���";
        }

        private void label����Tab1_Leave(object sender, System.EventArgs e)
        {
            ����������();
        }

        private void label����Tab2_Leave(object sender, System.EventArgs e)
        {
            ����������();
        }

        /// <summary>
        /// ����������� ���� �������� ������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContext��������������_Click(object sender, EventArgs e)
        {

               string iTest = ������������;

            // ���������� ��� �������� id ��������� �������� ������������.
            int idPersonDate = 0;

            int seletYear = Convert.ToInt16(this.������������) + 1;

            string filePatchLog = Application.StartupPath + @"\fileLog.txt";

            if (File.Exists(filePatchLog) == true)
            {
                File.Delete(filePatchLog);
                Log.WriteLine(filePatchLog, "�������� ���");
            }
            else
            {
                Log.WriteLine(filePatchLog, "������� ��� ����");
            }
            

            // ���������� ��� �������� ���������� ���.
            int inc = 0;
            StringBuilder builder = new StringBuilder();

            // ������ ��� �������� ������� � �� ��� ��������� id �����������.
            StringBuilder build��� = new StringBuilder();

            //Form�������� form = new Form��������(ds11, ������������,false);
            Form�������� form = new Form��������(ds11, seletYear.ToString(), false);

            // ������� � �������� ������� ���.
            form.CurrentYear = this.selectedYear;

            form.ShowDialog(this);

            // �������� �����.-
            �������������� docNumNext = form.�����������������������;

            if (form.DialogResult == DialogResult.OK)
            {

                DS1.��������Row row = form.��������������;

                // ����������� ���� ��� ������������� ���������.
                Guid guidCard = Guid.NewGuid();

                inc = form.IncrementDate;

                string patchToServer = string.Empty;

                // ���������� ��� �������� ����� ����� �� �������.
                string namFileServer = string.Empty;

                // ������� ��������� ������ ����������� ��������� ������ ������������.
                Item�������������������������� �������������������������� = form.�����������������;

                // �������� ������ ����������� ���������� � ������� ������� ������� ��������.
                this.ListPerson = form.ListPerson;

                // ���������� ����.

                //���� ���������� ���� ���������� ���������� ��������� �� �������.
                if (form.SaveDocServer == true)
                {
                    if (form.���������������� == true)
                    {
                        // ������� ���� � �����.
                        string filePatch = form.PathFileServer;

                        // ��� ��������� ����������.
                        //string archiver = @"C:\Program Files\7-Zip\7z.exe";

                        // ������� ��� ����� ������� ����� ��������������.
                        string archive = form.FileName;// +@"\*.*";

                        // GUID ������������ �������� �����.
                        string file = form.PathFileServer;

                        // �������� ��� ��� ����� ������ ����������� ������������ �����.
                        string namFileS = docNumNext.�����.ToString() + "-" + docNumNext.������� + "_" + file;
                        string namFile = docNumNext.�����.ToString() + "-" + docNumNext.�������;

                        // ���� � ���������� ���������� ����� � �������.
                        string patch = Application.StartupPath + @"\Archive\" + namFile + ".7z";

                        fileName = patch;

                        namFileServer = namFile;// +".7z";

                        // ��������� ���� ����� ������������ ����.
                        string patchDir = Application.StartupPath + @"\Archive\";

                        // ���������� �����. (������ ����������)
                        //Archiver.AddToArchive(archiver, archive, patch,patchDir);

                        Log.WriteLine(filePatchLog, "������������� ����� ������");

                        // ���� � 7z.dll.
                        string sevenZipDll = Application.StartupPath + @"\7z.dll";
                        if (archive.Length > 0)
                        {
                            // �������, ��� �������� ������� ��� ������ �� ������.
                            flagInsertCopyDoc = true;

                            Archiver.AddToArchive(sevenZipDll, archive, patch, patchDir);
                        }
                        else
                        {
                            // �������, ��� �������� ��� ������ �� ������ �� ��������.
                            flagInsertCopyDoc = false;

                            MessageBox.Show("�� �� ������� ����� ����� � ����������� �������� �� ������","��������",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                        }


                        Log.WriteLine(filePatchLog, "������������� ����� �����");

                        // ���� ���� ����� ������������ �����.
                        patchToServer = patchServerFile + @"\" + namFileS.Trim();

                        fileNameCopy = patchToServer;
                    }
                    else
                    {
                        return;
                    }
                }

                // ===Begin========������� ������� ���� �������� �������� � ���� ������.
                // �������� ������ �� �������. (������ ,)
                string[] s����s = row["���������"].ToString().Split(',');

                int id_�������� = Convert.ToInt32(row["id_��������"]);

                // �������� ����� ������.
                DateTime todoy = DateTime.Now;

                // ������� ������.
                int iCount = 1;

                //// ���������� ������ 
                //foreach (string str in s����s)
                //{
                //    string insert = "declare @id_" + iCount + "  int " +
                //                    "SELECT @id_" + iCount + " = id_���������� " +
                //                    "FROM [����������] " +
                //                    "where [������������������] = '" + str.Trim() + "' " +
                //                    "INSERT INTO [�����������������������������] " +
                //                               "([id����������] " +
                //                               ",[���������������] " +
                //                               ",[����������������] " +
                //                               ",[�����������������] " +
                //                               ",[id��������] " +
                //                               ",[�������������������]) " +
                //                         "VALUES " +
                //                               "(@id_" + iCount + " " +
                //                               ",'" + todoy + "' " +
                //                               ",NULL " +
                //                               ",NULL " +
                //                               ","+ id_�������� +" " +
                //                               ",NULL) ";

                //    // ������� � ������.
                //    builder.Append(insert);

                //    iCount++;
                //}
                //=============End=====================

                // ���� ������������ �� ������ �������� ������� ����� ������������ � ���������� �� ������
                // ����� ��������� ���� ������ ��������� �� ������ ��� �������� ������ ���������.
                if (flagInsertCopyDoc == false)
                {
                    form.���������������� = false;
                }
               
                // �� ���������, ��������� ����� �������� ������ �� ������������.
                if (form.FlagRecordRepeet == false)
                {
                    string queryInsert = string.Empty;

                    // ������ ��� �������� ��� ������ �����.
                    string md5 = string.Empty;


                    if (form.���������������� == true)
                    {
                        if (form.FlagAddDoc == true)
                        {

                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          " declare @id_�������� int " +
                                          " declare @������� int  " +
                                        " select top 1 @������� = ������� from �������� " +
                " where ���������� <= '" + seletYear.ToString().Trim() + "1231' and ���������� >= '" + ������������ + "1231' and FlagAuto is null " +
                  "order by id_�������� desc " +
                                         //" where ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                         // " and id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                                // " where FlagAuto is null) " + 
                                                //"order by ������� desc " +
                                                 "INSERT INTO �������� " +
                                                 "([id_���������] " +
                                                 ",[id_��������������] " +
                                                 ",[�����] " +
                                                ",[����������] " +
                                                ",[����������] " +
                                                ",[�����������������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[�������������������] " +
                                                ",[��������������] " +
                                                ",[�������] " +
                                                ",[����������������������] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                ",MD5 " +
                                                ",id�����������������������  " +
                                                ",DataWriterServerDoc " +
                                                ",NameFileDocumentVipNetEmailTitlePage " +
                                                //",FileDate " +
                                                //",FileDateTitlePage " +
                                                ",FlagAuto " +
                                                ",��� ) " +
                                                "VALUES " +
                                                "( " + row["id_���������"] + " " +
                                                "," + row["id_��������������"] + " " +
                                                ",'" + row["�����"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["�����������������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["�������������������"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["�������"] + " " +
                                                //"," + docNumNext.����� + " " +
                                                ", @������� + 1 " +
                                                ",'" + row["����������������������"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",'" + namFileServer + "'  " +
                                                ",'" + form.PathFileServer + "' " +
                                                ",'md5' " +
                                                "," + ��������������������������.Id + "  " +
                                                ", NULL " +
                                                ", NULL " +
                                                //", NULL " +
                                                //", NULL " +
                                                ", NULL " +
                                                ", '"+ form.FlagDsp +"' ) " +
                                                "SELECT @id_�������� = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                        else
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          " declare @id_�������� int " +
                                         " declare @������� int  " +
                                          " select top 1 @������� = ������� from �������� " +
                                    " where ���������� <= '" + seletYear.ToString().Trim() + "1231' and ���������� >= '" + ������������ + "1231' and FlagAuto is null " +
                                      "order by id_�������� desc " +
                                          ////" select top 1 @������� = ������� from �������� " +
                                          ////" where FlagAuto is null and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          ////"order by id_�������� desc " +
                                         // " select top 1 @������� = ������� from �������� " +
                                         // " where ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                         //// " where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                         // " and id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                                         // " where FlagAuto is null) " +
                                         //        "order by ������� desc " +
                                                 "INSERT INTO �������� " +
                                                 "([id_���������] " +
                                                 ",[id_��������������] " +
                                                 ",[�����] " +
                                                ",[����������] " +
                                                ",[����������] " +
                                                ",[�����������������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[�������������������] " +
                                                ",[��������������] " +
                                                ",[�������] " +
                                                ",[����������������������] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                 ",MD5 " +
                                                ",id�����������������������  " +
                                                 ",DataWriterServerDoc " +
                                                ",NameFileDocumentVipNetEmailTitlePage " +
                                                //",FileDate " +
                                                //",FileDateTitlePage " +
                                                ",FlagAuto " +
                                                ",��� ) " +
                                                "VALUES " +
                                                "( " + row["id_���������"] + " " +
                                                "," + row["id_��������������"] + " " +
                                                ",'" + row["�����"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["�����������������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["�������������������"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["�������"] + " " +
                                 //"," + docNumNext.����� + " " +
                                                ", @������� + 1 " +
                                                ",'" + row["����������������������"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",'" + namFileServer + "'  " +
                                                ",'" + form.PathFileServer + "' " +
                                                ",NULL " + 
                                                 "," + ��������������������������.Id + "  " +
                                                 ", NULL " +
                                                ", NULL " +
                                                //", NULL " +
                                                //", NULL " +
                                                ", NULL " +
                                                ", '" + form.FlagDsp + "' ) " +
                                                "SELECT @id_�������� = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                    }
                    else
                    {

                        if (form.FlagAddDoc == true)
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                            " declare @id_�������� int " +
                                         " declare @������� int  " +
                                          " select top 1 @������� = ������� from �������� " +
                                        " where ���������� <= '" + seletYear.ToString().Trim() + "1231' and ���������� >= '" + ������������ + "1231' and FlagAuto is null " +
                                          "order by id_�������� desc " +
                                                                  ////" select top 1 @������� = ������� from �������� " +
                                          ////" where FlagAuto is null and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          ////"order by id_�������� desc " +
                                          //" select top 1 @������� = ������� from �������� " +
                                          ////" where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          //" where ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          //" and id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                                          //" where FlagAuto is null) " +
                                          //       "order by ������� desc " +
                                                "INSERT INTO �������� " +
                                                 "([id_���������] " +
                                                 ",[id_��������������] " +
                                                 ",[�����] " +
                                                ",[����������] " +
                                                ",[����������] " +
                                                ",[�����������������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[�������������������] " +
                                                ",[��������������] " +
                                                ",[�������] " +
                                                ",[����������������������] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                ",MD5 " +
                                                ",id�����������������������  " +
                                                  ",DataWriterServerDoc " +
                                                ",NameFileDocumentVipNetEmailTitlePage " +
                                                //",FileDate " +
                                                //",FileDateTitlePage " +
                                                ",FlagAuto " +
                                                ",��� ) " +
                                                "VALUES " +
                                                "( " + row["id_���������"] + " " +
                                                "," + row["id_��������������"] + " " +
                                                ",'" + row["�����"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["�����������������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["�������������������"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["�������"] + " " +
                                 //"," + docNumNext.����� + " " +
                                                ", @������� + 1 " +
                                                ",'" + row["����������������������"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",NULL  " +
                                                ",'" + guidCard + "' " +
                                                ",'md5' " +
                                                "," + ��������������������������.Id + "  " +
                                                ", NULL " +
                                                ", NULL " +
                                                //", NULL " +
                                                //", NULL " +
                                                ", NULL " +
                                                ", '" + form.FlagDsp + "' ) " +
                                                "SELECT @id_�������� = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                        else
                        {

                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                            " declare @id_�������� int " +
                                         " declare @������� int  " +
                                         " select top 1 @������� = ������� from �������� " +
                                        " where ���������� <= '" + seletYear.ToString().Trim() + "1231' and ���������� >= '" + ������������ + "1231' and FlagAuto is null " +
                                          "order by id_�������� desc " +
                                          ////" select top 1 @������� = ������� from �������� " +
                                          ////" where FlagAuto is null and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          ////"order by id_�������� desc " +
                                          //" select top 1 @������� = ������� from �������� " +
                                          ////" where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          //" where ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          //" and id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                                          //" where FlagAuto is null) " +
                                          //       "order by ������� desc " +
                                                "INSERT INTO �������� " +
                                                 "([id_���������] " +
                                                 ",[id_��������������] " +
                                                 ",[�����] " +
                                                ",[����������] " +
                                                ",[����������] " +
                                                ",[�����������������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[�������������������] " +
                                                ",[��������������] " +
                                                ",[�������] " +
                                                ",[����������������������] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                               ",MD5 " +
                                                ",id�����������������������  " +
                                                ",DataWriterServerDoc " +
                                                ",NameFileDocumentVipNetEmailTitlePage " +
                                                //",FileDate " +
                                                //",FileDateTitlePage " +
                                                ",FlagAuto " +
                                                ",��� ) " +
                                                "VALUES " +
                                                "( " + row["id_���������"] + " " +
                                                "," + row["id_��������������"] + " " +
                                                ",'" + row["�����"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["�����������������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["�������������������"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["�������"] + " " +
                                                //"," + docNumNext.����� + " " +
                                                ", @������� + 1 " +
                                                ",'" + row["����������������������"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",NULL  " +
                                                ",'" + guidCard + "' " +
                                                ",NULL " +
                                                  "," + ��������������������������.Id + "  " +
                                                   ", NULL " +
                                                ", NULL " +
                                                //", NULL " +
                                                //", NULL " +
                                                ", NULL " +
                                                ", '" + form.FlagDsp + "' ) " +
                                                "SELECT @id_�������� = @@IDENTITY  ";

                                               

                            builder.Append(queryInsert);
                        }
                    }

                    // ���������� ������ ����������� ����� �������� � ������������� ���� ������� ��������.
                    foreach (string str in s����s)
                    {
                        string insert = "declare @id_" + iCount + "  int " +
                                        "SELECT @id_" + iCount + " = id_���������� " +
                                        "FROM [����������] " +
                                        "where [������������������] = '" + str.Trim() + "' " +
                                        "INSERT INTO [�����������������������������] " +
                                                   "([id����������] " +
                                                   ",[���������������] " +
                                                   ",[����������������] " +
                                                   ",[�����������������] " +
                                                   ",[id��������] " +
                                                   ",[�������������������]) " +
                                             "VALUES " +
                                                   "(@id_" + iCount + " " +
                                                   //",'" + ����SQL.����(todoy.ToShortDateString()) + "' " +
                                                   ",GETDATE() " +
                                                   ",NULL " +
                                                   ",NULL " +
                                                   ",@id_�������� " +
                                                   ",NULL) ";

                        // ������� � ������.
                        builder.Append(insert);

                        iCount++;
                    }

                    // ���������� ������ � ��������� ������� ���������, ���� ��������� ��������� � ������������ ������� � ���������� ������� ������� ������� ��������.
                    foreach (PersonRecepient person in this.ListPerson)
                    {
                        string insert = "INSERT INTO [������������������������������������] " +
                                        "([id_person] " +
                                       ",[id_�����������������] " +
                                       ",[id_��������]) " +
                                       "VALUES " +
                                       "(" + person.ID + " " +
                                       "," + ��������������������������.Id + " " +
                                       ",@id_�������� ) ";

                        // ������� � ������.
                        builder.Append(insert);
                    }

                    // ������� ��������� ��� �������� ������������ ������.
                    builder.Append(form.QueryPersonDateForCardInput);
                                                
                    //builder.Append(queryInsert + "COMMIT TRANSACTION ");
                    builder.Append("COMMIT TRANSACTION ");

                    string sTest = builder.ToString().Trim();
                }

                // ��������� ����� ������ ������������.
                if (form.FlagRecordRepeet == true)
                {
                    string queryInsert = string.Empty;
                    if (form.���������������� == true)
                    {
                        queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                      "begin transaction  " +
                                        " declare @id_�������� int " +
                                      " declare @������� int " +
                                       " select top 1 @������� = ������� from �������� " +
                " where ���������� <= '" + seletYear.ToString().Trim() + "1231' and ���������� >= '" + ������������ + "1231' and FlagAuto is null " +
                  "order by id_�������� desc " +
                            ////" select top 1 @������� = ������� from �������� " +
                            ////   " where FlagAuto is null and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                            ////   "order by id_�������� desc " +
                            //               "select top 1 @������� = ������� from �������� " +
                            //              //"where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                            //              " where ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                            //              " and id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                            //              " where FlagAuto is null) " +
                            ////"select top 1 @������� = [�������] from �������� " +
                            ////"where ���������� >= '" + seletYear.ToString().Trim() + "0101' " +  and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                            //                     "order by ������� desc " +
                                            " INSERT INTO �������� " +
                                             "([id_���������] " +
                                             ",[id_��������������] " +
                                             ",[�����] " +
                                            ",[����������] " +
                                            ",[����������] " +
                                            ",[�����������������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[�������������������] " +
                                            ",[��������������] " +
                                            ",[�������] " +
                                            ",[����������������������] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                             ",[FlagCardRepeet] " +
                                            ",NameFileDocument ) " +
                                            "VALUES " +
                                            "( " + row["id_���������"] + " " +
                                            "," + row["id_��������������"] + " " +
                                            ",'" + row["�����"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["�����������������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["�������������������"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["�������"] + " " +
                            //"," + docNumNext.����� + " " +
                                            ", @������� + 1 " +
                                            ",'" + row["����������������������"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                             ",'" + namFileServer + "'  " +
                                            ",'" + form.PathFileServer + "' ) " +
                                           "INSERT INTO �������������� " +
                                             "([id_���������] " +
                                             ",[id_��������������] " +
                                             ",[�����] " +
                                            ",[����������] " +
                                            ",[����������] " +
                                            ",[�����������������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[�������������������] " +
                                            ",[��������������] " +
                                            ",[�������] " +
                                            ",[����������������������] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                            ",id_����������������  " +
                                            ",�������������� " +
                                            ",FlagControl)" +
                                            "VALUES " +
                                            "( " + row["id_���������"] + " " +
                                            "," + row["id_��������������"] + " " +
                                            ",'" + row["�����"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["�����������������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["�������������������"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["�������"] + " " +
                                            "," + docNumNext.����� + " " +
                                            ",'" + row["����������������������"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                            ",@@IDENTITY " +
                                            "," + inc + " " +
                                            ",'False') ";                     }
                    else
                    {
                        queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                      "begin transaction  " +
                                        " declare @id_�������� int " +
                                      "declare @������� int " +
                                       " select top 1 @������� = ������� from �������� " +
                " where ���������� <= '" + seletYear.ToString().Trim() + "1231' and ���������� >= '" + ������������ + "1231' and FlagAuto is null " +
                  "order by id_�������� desc " +
                            ////" select top 1 @������� = ������� from �������� " +
                            ////   " where FlagAuto is null and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                            ////   "order by id_�������� desc " +
                            //              "select top 1 @������� = ������� from �������� " +
                            //              //"where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                            //              " where ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                            //              " and id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                            //              " where FlagAuto is null) " +
                            ////"select top 1 @������� = [�������] from �������� " +
                            ////"where ���������� >= '" + seletYear.ToString().Trim() + "0101' " +  and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                            //                     "order by ������� desc " +
                                      "INSERT INTO �������� " +
                                             "([id_���������] " +
                                             ",[id_��������������] " +
                                             ",[�����] " +
                                            ",[����������] " +
                                            ",[����������] " +
                                            ",[�����������������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[�������������������] " +
                                            ",[��������������] " +
                                            ",[�������] " +
                                            ",[����������������������] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                             ",[FlagCardRepeet] " +
                                            ",NameFileDocument ) " +
                                            "VALUES " +
                                            "( " + row["id_���������"] + " " +
                                            "," + row["id_��������������"] + " " +
                                            ",'" + row["�����"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["�����������������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["�������������������"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["�������"] + " " +
                            //"," + docNumNext.����� + " " +
                                            ", @������� + 1 " +
                                            ",'" + row["����������������������"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                             ",NULL  " +
                                             ",NULL ) " +
                                           "INSERT INTO �������������� " +
                                             "([id_���������] " +
                                             ",[id_��������������] " +
                                             ",[�����] " +
                                            ",[����������] " +
                                            ",[����������] " +
                                            ",[�����������������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[�������������������] " +
                                            ",[��������������] " +
                                            ",[�������] " +
                                            ",[����������������������] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                            ",id_����������������  " +
                                            ",�������������� " +
                                            ",FlagControl)" +
                                            "VALUES " +
                                            "( " + row["id_���������"] + " " +
                                            "," + row["id_��������������"] + " " +
                                            ",'" + row["�����"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["�����������������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["�������������������"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["�������"] + " " +
                                            "," + docNumNext.����� + " " +
                                            ",'" + row["����������������������"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                            ",@@IDENTITY " +
                                            "," + inc + " " +
                                            ",'False') ";
                                           
                    }

                    // ������� ������ ������� �� Insert.
                    builder.Append(queryInsert);

                    // ������� ��������� ��� �������� ������������ ������.
                    builder.Append(form.QueryPersonDateForCardInput);

                    // ������� ����������.
                    builder.Append("COMMIT TRANSACTION");

                    string iTest2 = "";
                }

                //// �������� ������.
                ������������ connectBD = new ������������();
                string sCon = connectBD.�����������������();

                // ���� �������� �������� ����� �����.
                bool flagCopyServer = false;

                string sTest2 = builder.ToString().Trim();

                // �������� ������ �� ������� (� ��������� �� � ������ ����������.
                using (SqlConnection con = new SqlConnection(sCon))
                {

                    //Log.WriteLine(filePatchLog, "����� ��������� ���� �� ������");

                    //if (form.SaveDocServer == true)
                    //{
                    //    try
                    //    {
                    //        if (form.���������������� == true)
                    //        {

                    //            Log.WriteLine(filePatchLog, "�������� ���� �� ������");

                    //            // �������� ������� �� �������� ��� ������ �� ������.
                    //            if (flagInsertCopyDoc == true)
                    //            {
                    //                //��������� ���� �� ������ �������� ����������.
                    //                //File.Copy(fileName, fileNameCopy, true);
                    //            }
                    //            Log.WriteLine(filePatchLog, "�������� ���������� ���� �� ������");

                    //            //���� ���� ������������ ������� ������� ���� � true.
                    //            flagCopyServer = true;
                    //        }
                    //    }
                    //    catch(Exception exp)
                    //    {
                    //        Log.WriteLine(filePatchLog, "������ ��� ����������� - ");
                    //        Log.WriteLine(filePatchLog, exp.Message);
                    //        MessageBox.Show("������ ��� ����������� �����");

                    //        flagCopyServer = false;

                    //        return;
                    //    }

                    //    string fileTest = fileNameCopy;
                    //    if (File.Exists(fileNameCopy) == true)
                    //    {
                    //        Log.WriteLine(filePatchLog, "�������� ������ �� ������");

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
                            // ���� ���� ������������ ������� ������� ���� � true.
                            flagCopyServer = true;

                            con.Open();
                            SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                            com.ExecuteNonQuery();
                    //}
                }


                //ds11.��������.Add��������Row(row);
                ��������������();

                // ����� �� ������ �������� � ������� ����� ���������.
                string queryNumDoc = "select id_��������,�������,��������� from [��������] " +
                                     "where GuidName = '" + guidCard + "' ";

                string �������� = string.Empty;

                DataTable tabNum;

                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();

                    SqlDataAdapter da = new SqlDataAdapter(queryNumDoc, con);

                    DataSet ds = new DataSet();

                    da.Fill(ds, "numDoc");

                    tabNum = ds.Tables["numDoc"];
                }

                �������� = tabNum.Rows[0]["�������"].ToString().Trim() + "/" + tabNum.Rows[0]["���������"].ToString().Trim();

                string ����� = ��������;

                // ������� ����� id ��������.
                string idCard = tabNum.Rows[0]["id_��������"].ToString().Trim();
                
                // ������� ����� ������������������� ���������.
                FormMessage frmMessage = new FormMessage(�����);
                frmMessage.NumCardDoc = idCard.Trim();
                frmMessage.�������������� = ��������;
                frmMessage.�������������������������� = ��������������������������;
                frmMessage.TopMost = true;
                frmMessage.ShowDialog();

            }
        }

        /// <summary>
        /// ����������� ���� �������� ��������� ������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContext�����������������������_Click(object sender, EventArgs e)
        {
            string iTest = ������������;

            int seletYear = Convert.ToInt16(this.������������) + 1;

            
            Form����������������� form = new Form�����������������(ds11, ������������, false);

            // ��������� ���� � false.
            form.Flag����������� = false;

            // ��������� �������.
            form.������� = "";
            
            DialogResult result = form.ShowDialog(this);
            if (result == DialogResult.OK)
            {

                // ������� ��������� ������ ����������� ��������� ������ ������������.
                Item�������������������������� �������������������������� = form.�����������������;


                DS1.�����������������Row row = form.�����������������������;

                �������������� doc = new ��������������();
                doc.����� = Convert.ToInt16(form.�����Doc.�����);

                // ������� ����������.
                numberPrefix = string.Empty;

                // ������� ������� ������ ���������.
                numberPrefix = form.���������������������;

                //ds11.�����������������.Add�����������������Row(row); 

                List<int> listId������� = form.ListID��������;
                //List<�����������������> listOP = form.List�����������������;

                // ���� �� ����� ������������ ������.
                if (form.Flag����������� == false)
                {

                  // ������ ��� �������� SQL ����������, ��� ���������� � ����� ����������.
                    StringBuilder buildInsert = new StringBuilder();

                    // ���������� ��� �������� ������
                    string numDirect = string.Empty;
                    
                    // �������� �������� � ��� ��� ���.
                    string query = string.Empty;

                    if (Convert.ToBoolean(form.FlagDsp) == false)
                    {
                        query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                       "begin transaction  " +
                                       "declare @numDoc int " +
                                        "select top 1 @numDoc = ��������������� from ����������������� " +
                                       " where ���� >= '" + ������������ + "1201' and ���� <= '" + (Convert.ToInt32(������������) + 1).ToString().Trim() + "1231' " +
                            "order by id_�������� desc " +
                            ////           "select top 1 @numDoc = ��������������� from ����������������� " +
                            //////"where ���� >= '" + seletYear.ToString().Trim() + "0101' and ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                            ////"where ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                            ////" and " +
                            ////" id_�������� in (SELECT MAX(id_��������) FROM [�����������������] " +
                            ////" where FlagAutho is null) " +
                            ////" order by id_�������� desc " +
                                       "declare @key int " +
                                       "INSERT INTO ����������������� " +
                                       "([����] " +
                                       ",[�������������] " +
                                       ",[id_�������������] " +
                                       ",[�������������������] " +
                                       ",[���������������] " +
                                       ",[id_��������] " +
                                       ",[����������] " +
                                       ",[id_������������������] " +
                                       ",[����������������������] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ",id����������������������� " +
                                       ",FlagAutho " +
                                       ",��� ) " +
                                       "VALUES " +
                                       "('" + ����SQL.����(Convert.ToDateTime(row["����"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row["�������������"] + "' " +
                                       "," + row["id_�������������"] + " " +
                                       ",'" + row["�������������������"] + "' " +
                            //"," + row["���������������"] + " " +
                            //", "+ doc.����� + " " +
                                       ", @numDoc + 1  " +
                                       "," + row["id_��������"] + " " +
                                       ",'" + row["����������"] + "' " +
                            //","+ row["id_������������������"]+" " +
                                       ",NULL " +
                            //",'"+ form.�������.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       // ",NULL " +
                                       //",NULL " +
                                       ", " + ��������������������������.Id + " " +
                                       ", NULL " +
                                       ",'" + form.FlagDsp + "' ) " +
                                       "set @key = @@IDENTITY ";
                    }
                    else
                    {
                        query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                       "begin transaction  " +
                                       "declare @numDoc int " +
                                        "select top 1 @numDoc = ��������������� from ����������������� " +
                                       " where ���� >= '" + ������������ + "1201' and ���� <= '" + (Convert.ToInt32(������������) + 1).ToString().Trim() + "1231' " +
                            "order by id_�������� desc " + 
                            //  "select top 1 @numDoc = ��������������� from ����������������� " +
                            ////"where ���� >= '" + seletYear.ToString().Trim() + "0101' and ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                            //"where ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                            //" and " +
                            //" id_�������� in (SELECT MAX(id_��������) FROM [�����������������] " +
                            //" where FlagAutho is null) " +
                            //" order by id_�������� desc " +

                                       "declare @key int " +
                                       "INSERT INTO ����������������� " +
                                       "([����] " +
                                       ",[�������������] " +
                                       ",[id_�������������] " +
                                       ",[�������������������] " +
                                       ",[���������������] " +
                                       ",[id_��������] " +
                                       ",[����������] " +
                                       ",[id_������������������] " +
                                       ",[����������������������] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ",id����������������������� " +
                                       ",FlagAutho " +
                                       ",���  " +
                                       ",���Desc ) " +
                                       "VALUES " +
                                       "('" + ����SQL.����(Convert.ToDateTime(row["����"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row["�������������"] + "' " +
                                       "," + row["id_�������������"] + " " +
                                       ",'" + row["�������������������"] + "' " +
                            //"," + row["���������������"] + " " +
                            //", "+ doc.����� + " " +
                                       ", @numDoc + 1  " +
                                       "," + row["id_��������"] + " " +
                                       ",'" + row["����������"] + "' " +
                            //","+ row["id_������������������"]+" " +
                                       ",NULL " +
                            //",'"+ form.�������.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       // ",NULL " +
                                       //",NULL " +
                                       ", " + ��������������������������.Id + " " +
                                       ", NULL " +
                                       ",'" + form.FlagDsp + "'  " +
                                       ",'���' )" +
                                       "set @key = @@IDENTITY ";
                    }

                    buildInsert.Append(query);

                    // ��������� � ������ ������� �� ������� SQL ���������� �� ������� � ������� [���������������������������������������].

                    string sInsert = string.Empty;
                    sInsert = String.Format(form.QueryInsert.Trim(), "@key");


                    buildInsert.Append(sInsert.Trim());

                    // �������� ����������.
                    buildInsert.Append("COMMIT TRANSACTION ");

                    string sTestInsertCardInput = buildInsert.ToString();

                    string sTest = "";

                    // ������� ������ ��� �������� ��������� �������� ����� ��������������.
                    //form.List�����������������.Clear();

                    ������������ connBD = new ������������();
                    string sCon = connBD.�����������������();

                    SqlConnection con = new SqlConnection(sCon);
                    con.Open();
                    SqlCommand com = new SqlCommand(buildInsert.ToString(), con);
                    //com.ExecuteNonQuery();
                    con.Close();
                }
                else
                {

                   
                    // ���������.
                    DS1.�����������������Row row2 = form.�����������������������;

                    int i = row2.id_������������������;

                    // ������ ��� �������� �������.
                    System.Text.StringBuilder builder = new System.Text.StringBuilder();

                    // ��������� ���� � FALSE.
                    string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                    "begin transaction  " +
                                    " declare @numCard  int " +
                                    //" select @numCard = MAX(���������������) from ����������������� " +
                                    ////" where ���� >= '20170101' and FlagAutho is null " +
                                    //"where ���� >= '" + seletYear.ToString().Trim() + "0101' and FlagAutho is null " +
                         "select top 1 @numCard = ���������������  from ����������������� " +
                        //"where ���� >= '" + seletYear.ToString().Trim() + "0101' and ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                         "where ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                          " and " +
                        " id_�������� in (SELECT MAX(id_��������) FROM [�����������������] " +
                        " where FlagAutho is null) " +
                         "order by id_�������� desc " +
                                        //////" select top 1 @������� = ������� from �������� " +
                                        //////  " where FlagAuto is null and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                        //////  "order by id_�������� desc " +
                                           "declare @key int " +
                                   "INSERT INTO ����������������� " +
                                   "([����] " +
                                   ",[�������������] " +
                                   ",[id_�������������] " +
                                   ",[�������������������] " +
                                   ",[���������������] " +
                                   ",[id_��������] " +
                                   ",[����������] " +
                                   ",[id_������������������] " +
                                   ",[����������������������] " +
                                   ",[FlagPersonData] " +
                                   ",[GUID] " +
                                   //",FileData " +
                                   //",FileDateTitlePage " +
                                   ", id�����������������������)" +
                                   "VALUES " +
                                   "('" + ����SQL.����(Convert.ToDateTime(row2["����"]).ToShortDateString().Trim()) + "' " +
                                   ",'" + row2["�������������"] + "' " +
                                   "," + row2["id_�������������"] + " " +
                                   ",'" + row2["�������������������"] + "' " +
                        //"," + row2["���������������"] + " " +
                                    //", " + doc.����� + " " +
                                    ", @numCard + 1  " +
                                   "," + row2["id_��������"] + " " +
                                   ",'" + row2["����������"] + "' " +
                                   "," + row2["id_������������������"] + " " +
                        //",NULL " +
                        //",'"+ form.�������.Trim() +"' " +
                                   ",NULL" +
                                   ",'" + row2["FlagPersonData"] + "' " +
                                   ",'" + form.StrGuid.Trim() + "'  " +
                                   //",NULL " +
                                   //",NULL " +
                                   ", " + ��������������������������.Id + " )" +
                                   " declare @idCard int " +
                                   "select top 1 @idCard = id_��������  from ����������������� " +
                                   "order by id_�������� desc ";
                    

                    builder.Append(query);

                    string ������������������ = string.Empty;

                    DataRow[] rowsSelect = ds11.���������������������.Select("id_�������������= "+ Convert.ToInt32(row2["id_�������������"]) +" ");
                    foreach (DataRow item in rowsSelect)
                    {
                        ������������������ = item["������������������"].ToString().Trim();
                    }

                    string ������������������� = "��� �����. � ���. ��������� " + row2["�������������"].ToString().Trim() + "-" + row2["�������������������"].ToString().Trim() + "-" + ������������������ + "/" + doc.�����.ToString().Trim();// row2["���������������"].ToString().Trim();
                    //string ������������������� = "��� �����. � ���. ��������� " + row2["�������������"].ToString().Trim() + "-" + row2["�������������������"].ToString().Trim() + "-" + ������������������ + "/CAST(@numCard + 1 AS nvarchar) ";// +row2["���������������"].ToString().Trim();

                    // ��������� ���� � TRUE.
                    string queryUpdate = "UPDATE [��������] " +
                                         "SET ������������������� = '" + ������������������� + "' " + //' + CAST(@numCard + 1 AS nvarchar) " +
                                         //"FlagPersonData = '" + row["FlagPersonData"] + "' " +
                                         ",����� = 'True' " +
                                         "where id_�������� = " + row["id_������������������"] + " ";
                    // ������ ������ ��������� �� ���������� ������ � �� �������������� � ������ ������, ����� ��������� �� � ����� ����������.
                    builder.Append(queryUpdate);

                    string sTestNum = builder.ToString().Trim();

                    // ������ �� ������� id � ��������� ������� �����������������������.
                    foreach (����������������� itm in form.List�����������������)
                    {
                        string queryIns = "INSERT INTO [���������������������������������������] " +
                                       "([id_��������] " +
                                       ",[id_�����������������]) " +
                                       "VALUES " +
                                       //"('" + row.id_�������� + "' " +
                                       "( @idCard " +
                                       ",'" + itm.Id_����������������� + "' ) ";

                        builder.Append(queryIns);
                    }


                    // �������� ��������� ������� �������� �����������������.
                    foreach (int id�� in listId�������)
                    {

                        string queryId�� = "INSERT INTO [����������������������������������] " +
                                           "([id_����������������] " +
                                           ",[id_�����������������]) " +
                                           "VALUES " +
                                           "(" + id�� + " " +
                                           //"," + row.id_�������� + " ) " +
                                            ",@idCard ) " + 
                                           "update �������� " +
                                           "set ������������������� = '" + form.��������������.Trim() + "' " + " + CAST(@numCard + 1 AS nvarchar) " +
                                           "where id_�������� = "+ id�� +" ";
                        
                        builder.Append(queryId��);
                    }

                    // ��������, ��� �������� �� ������� �� �������� ����� � ������� ��������� �������.
                    �������������� card = new ��������������(Convert.ToInt32(row["id_������������������"]));
                    bool flagStatusRepeet = card.������������������������();

                    // ���� ������ = true ������ �� ����� ���� � ���������� �� ������� ������������ ���������� ������ �����.
                    if (flagStatusRepeet == true)
                    {
                        // ������ ��� ���������� ������� � ����� ����������.
                        //StringBuilder querTransact = new StringBuilder();

                        /*
                         * �������� �������� �� �� ���� �������� ������� ��� ���.
                         * ��� ����� ������ �������� � ���� ����� � ������� ��������, ���� ����������� �������� False 
                         * ����� �� �������� �� �������� ������� � ��������� ������ ���.
                        */
                        ������������ bdConnect = new ������������();
                        using (SqlConnection conn = new SqlConnection(bdConnect.�����������������().Trim()))
                        {
                            conn.Open();
                            �������������� card2 = new ��������������(Convert.ToInt32(row["id_������������������"]));
                            bool flagVD = card2.Get��������������(conn);

                            // ���� �� �������� �������� �������� �������.
                            if (flagVD == false)
                            {
                                // ��������� �������� ���� ����� ������� �������� � True, � ��� �� �������� ���� � �������������� �����������, ��� ��� ���� ��� �� ������ ������ ��� �����.
                                //string queryUp = " update �������� " +
                                //               "set ����� = 'True' " +
                                //               "where id_�������� = " + Convert.ToInt32(row["id_������������������"]) + " " + 
                                //               "update �������������� " +
                                //               "set FlagControl = 'True' " + 
                                //               "where id_���������������� = "+ Convert.ToInt32(row["id_������������������"]) +" ";


                                string queryUp = " update �������� " +
                                              "set ����� = 'True' " +
                                              "where id_�������� = " + Convert.ToInt32(row["id_������������������"]) + " " +
                                              " declare @date datetime " +
                                              "declare @day int " +
                                              "declare @SetDate datetime " +
                                              "select @date = ��������������,@day = �������������� from �������������� " +
                                              "where id_���������������� = " + Convert.ToInt32(row["id_������������������"]) + " " +
                                              "SELECT @SetDate = DATEADD(day, @day, @date); " +
                                              "update �������������� " +
                                              "set FlagControl = 'True' " +
                                              ",�������������� = @SetDate " +
                                              "where id_���������������� = " + Convert.ToInt32(row["id_������������������"]) + " ";

                                //string queryUp = " update �������� " +
                                //              "set ����� = 'True' " +
                                //              "where id_�������� = " + Convert.ToInt32(row["id_������������������"]) + " ";
                             
                                
                                //querTransact.Append(query);
                                builder.Append(queryUp);
                            }

                            // ���� ����� ���������.
                            if (flagVD == true)
                            {
                                // �������� �������� � ���� �������������� � ������� �������������� �� ���������� ���� ��������� � ���� ��������������.
                                string queryUpdatDate = " declare @date datetime " +
                                                        "declare @day int " +
                                                        "declare @SetDate datetime " +
                                                        "select @date = ��������������,@day = �������������� from �������������� " +
                                                        "where id_���������������� = "+ Convert.ToInt32(row["id_������������������"]) +" " +
                                                        "SELECT @SetDate = DATEADD(day, @day, @date); " +
                                                        "update �������������� " +
                                                        "set �������������� = @SetDate " +
                                                        "where id_���������������� = "+ Convert.ToInt32(row["id_������������������"]) +" ";

                                //querTransact.Append(queryUpdatDate);
                                builder.Append(queryUpdatDate);
                            }
                            

                        }

                    }

                    // �������� ����������.
                    builder.Append("COMMIT TRANSACTION ");

           

                    string queryTest = builder.ToString().Trim();

                    // �������� ������.
                    ������������ strConnectBD = new ������������();
                    string strConn = strConnectBD.�����������������();

                    // ������� ���������� � �������� ������.
                    SqlConnection con = new SqlConnection(strConn);
                    con.Open();
                    SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    com.ExecuteNonQuery();

                    // ������� ����������.
                    con.Close();

                }

                ��������������();

                string iTest2 = "test";

                // ������� ����� ���������.
                NumOutputCardVipNet numDoc = GetNumDocOutVipNet(form.StrGuid);

                //string numberDocument = form.���������������������.Trim() + "/" + numDoc.Trim();

                string numberDocument = numberPrefix + "/" + numDoc.���������������.Trim();

                // ������� ��������� � ����� �������.
                FormMessage message = new FormMessage(numberDocument.Trim());
                message.TopMost = true;
                message.�������������������������� = ��������������������������;
                message.NumCardDoc = numDoc.Id.ToString().Trim();
                message.�������������� = numberDocument;
                message.ShowDialog();

            }
        }

        /// <summary>
        /// ����������� ���� �������� ������ � ������� ������� ���������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContext��������������_Click(object sender, EventArgs e)
        {
            // ���������� ��� �������� ���������� ���.
            int inc = 0;

            // ���������� ��� �������� ������ �����.
            string namFile = string.Empty;

            string namFileServer = string.Empty;
            string patchToServer = string.Empty;

            // ������ ��� �������� �������.
            StringBuilder builderUpdate = new StringBuilder();

            // ��������� ������ �������� ����������.
            builderUpdate.Append("SET TRANSACTION ISOLATION LEVEL serializable begin transaction ");

            // �������� ������ ������������ � ���������� ������:
            BindingManagerBase bmb = this.BindingContext[dataGrid����������������.DataSource, dataGrid����������������.DataMember];
            bmb.Position = dataGrid����������������.CurrentCell.RowNumber;
            dataGrid����������������.Select(dataGrid����������������.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            int id�������� = (int)drv["id_��������"];
            Form�������� form = new Form��������(ds11, id��������, ������������);

            string testDsp = form.FlagDsp;

            form.ShowDialog(this);

            �������������� docNum = new ��������������();
            docNum = form.�����������������������;

            //doc.����� = Convert.ToInt16(form.�����������������������.�����);

            string strMd5 = string.Empty;

            if(form.FlagAddDoc == true)
            {
                strMd5 = "md5";
            }
            else
            {
                strMd5 = "0";
            }
                       

            // ������� ����� ��������� ������� �� �����������.
            �������������� docNumNext = form.�����������������������;

            // ��� ���������.
            string ������������ = form.������������;

            if (form.DialogResult == DialogResult.OK)
            {
                inc = form.IncrementDate;

                string sTest������� = form.FlagDsp;

                // ������� ������ ����������� ���������.
                Item�������������������������� �������������������� = form.�����������������;

                //���� ���������� ���� ���������� ���������� ��������� �� �������.
                if (form.SaveDocServer == true)
                {
                    if (form.���������������� == true)
                    {
                        // ������� ���� � �����.
                        string filePatch = form.PathFileServer;

                        // ��� ��������� ����������.
                        string archiver = @"C:\Program Files\7-Zip\7z.exe";

                        // ������� ��� ����� ������� ����� ��������������.
                        string archive = form.FileName;// +@"\*.*";

                        // GUID ������������ �������� �����.
                        string file = form.PathFileServer;

                        //string namFileS = docNumNext.�����.ToString() + "-" + docNumNext.������� + "_" + file;

                        string namFileS = docNumNext.�������.Trim().Replace("/","-") + "_" + file;

                        //string namFile = docNumNext.�����.ToString() + "-" + docNumNext.�������;

                        // �������� ��� ��� ����� ������ ����������� ������������ �����.
                        //string namFile = Guid.NewGuid().ToString();
                        
                        namFile = docNumNext.�������;

                        // ���� � ���������� ���������� ����� � �������.
                        string patch = Application.StartupPath + @"\Archive\" + namFileS + ".7z";

                        fileName = patch;

                        namFileServer = namFile;// +".7z";

                        // ���� � ����� ��� ���������� �������� ������
                        string patchDir = Application.StartupPath + @"\Archive\";

                        // ���������� �����. (������ ����������)
                        //Archiver.AddToArchive(archiver, archive, patch,patchDir);

                        // ���� � 7z.dll.
                        string sevenZipDll = Application.StartupPath + @"\7z.dll";
                        //Archiver.AddToArchive(sevenZipDll, archive, patch, patchDir);


                        // ���������� ����� ����� ����������.


                        // ���� ���� ����� ������������ �����.
                        //patchToServer = patchServerFile + @"\" + namFile.Trim(); //
                        patchToServer = patchServerFile + @"\" + namFileS.Trim();

                        // ��� ����� �� �������.
                        fileNameCopy = patchToServer;
                    }
                }
                else
                {
                    fileNameCopy = ������������;
                    namFileServer = form.������������;
                }

                // ===Begin========������� ������� ���� �������� �������� � ���� ������.
                DS1.��������Row row = form.��������������;

                // �������� ������ �� �������. (������ ,)
                string[] s����s = row["���������"].ToString().Split(',');

                int id_�������� = id��������;

                string �������� = "delete ����������������������������� " +
                                  "where id�������� = " + id_�������� + " ";

                builderUpdate.Append(��������);

                // �������� ����� ������.
                DateTime todoy = DateTime.Now;

                // ������� ������.
                int iCount = 1;

                // ���������� ������ 
                foreach (string str in s����s)
                {
                    string insert = "declare @id_" + iCount + "  int " +
                                    "SELECT @id_" + iCount + " = id_���������� " +
                                    "FROM [����������] " +
                                    "where [������������������] = '" + str.Trim() + "' " +
                                    "INSERT INTO [�����������������������������] " +
                                               "([id����������] " +
                                               ",[���������������] " +
                                               ",[����������������] " +
                                               ",[�����������������] " +
                                               ",[id��������] " +
                                               ",[�������������������]) " +
                                         "VALUES " +
                                               "(@id_" + iCount + " " +
                                               ",'" + ����SQL.����(todoy.ToShortDateString()) + "' " +
                                               ",NULL " +
                                               ",NULL " +
                                               "," + id_�������� + " " +
                                               ",NULL) ";

                    // ������� � ������.
                    builderUpdate.Append(insert);

                    iCount++;
                }



                DS1TableAdapters.��������TableAdapter ������� = new RegKor.DS1TableAdapters.��������TableAdapter();

                if (form.��������������.FlagCardRepeet == false)
                {

                    int Test = docNumNext.�����;

                    ControlFlagRepeet cfr = new ControlFlagRepeet(form.��������������.id_��������, form.��������������.FlagCardRepeet);
                    bool flag = cfr.CompareRepet();

                    if (flag == true && form.��������������.FlagCardRepeet == false)
                    {
                        �������������� doc = new ��������������();

                        string queryUpdate = "UPDATE [��������] " +
                                    "SET [id_���������] = " + form.��������������.id_��������� + " " +
                                    ",[id_��������������] = " + form.��������������.id_�������������� + " " +
                                    ",[�����] = '" + form.��������������.����� + "' " +
                                    ",[����������] = '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ",[����������] = '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ",[�����������������] = '" + form.��������������.�����������������.Trim() + "' " +
                                    ",[����������] = '" + form.��������������.���������� + "' " +
                                    //",[���������] = '" + form.��������������.���������.Trim() + "' " +
                                     ",[���������] = '" + docNumNext.������� + "' " +
                                    ",[����������] = '" + form.��������������.���������� + "' " +
                                    ",[���������] = '" + form.��������������.���������.Trim() + "' " +
                                    ",[�������������������] = '" + form.��������������.�������������������.Trim() + "'  " +
                                    ",[��������������] = '" + ����SQL.����(form.��������������.��������������.ToShortDateString()) + "' " +
                                    //",[�������] = " + form.��������������.������� + " " +
                                    ",[�������] = " + docNumNext.����� + " " +
                                    //",[�������] = " + docNum.����� + " " +
                                    ",[����������������������] = '' " +
                                    ",[FlagPersonData] = '" + form.��������������.FlagPersonData + "' " +
                                    ",[FlagCardRepeet] = '" + form.��������������.FlagCardRepeet + "' " +
                                    ",[NameFileDocument] = '" + fileNameCopy + "' " +
                                    ",GuidName = '" + form.PathFileServer.Trim() + "' " +
                                     ",md5 = '" + strMd5.Trim() + "' " +
                                     ",id����������������������� = " + ��������������������.Id + " " +
                                     ",��� = '" + form.FlagDsp + "' " +
                                    "WHERE id_�������� = " + form.��������������.id_�������� + " " +
                                    "DELETE FROM [��������������] " +
                                    "WHERE id_���������������� = " + form.��������������.id_�������� + " ";


                        builderUpdate.Append(queryUpdate);

                        //ExecuteQuery exe = new ExecuteQuery(builderUpdate.ToString().Trim());
                        //exe.Excecute();
                    }
                    if (flag == false && form.��������������.FlagCardRepeet == false)
                    {
                        �������.Update(form.��������������);

                        string queryUpdate = "UPDATE [��������] " +
                                             "set [NameFileDocument] = '" + namFileServer + "' " +
                                             ",GuidName = '"+ form.PathFileServer.Trim() +"' " +
                                             ",md5 = '" + strMd5.Trim() + "' " +
                                             ",id����������������������� = " + ��������������������.Id + " " +
                                             ",��� = '" + form.FlagDsp + "' " +
                                             "WHERE id_�������� = " + form.��������������.id_�������� + " ";

                        builderUpdate.Append(queryUpdate);

                        //ExecuteQuery exe = new ExecuteQuery(builderUpdate.ToString().Trim());
                        //exe.Excecute();
                                              
                    }

                    // ������� ��������� ������ ����������� ��������� ������ ������������.
                    Item�������������������������� �������������������������� = form.�����������������;


                    // ������� ������ ����������� ������� � ���������� ������� ������� ��������.
                    string queryDelete = "delete dbo.������������������������������������ " +
                                         "where id_�������� = " + form.��������������.id_�������� + " ";

                    this.ListPerson = form.ListPerson;

                    builderUpdate.Append(queryDelete);

                    if (this.ListPerson != null)
                    {
                        // ���������� ������ � ��������� ������� ���������, ���� ��������� ��������� � ������������ ������� � ���������� ������� ������� ������� ��������.
                        foreach (PersonRecepient person in this.ListPerson)
                        {
                            string insert = "INSERT INTO [������������������������������������] " +
                                            "([id_person] " +
                                           ",[id_�����������������] " +
                                           ",[id_��������]) " +
                                           "VALUES " +
                                           "(" + person.ID + " " +
                                           "," + ��������������������������.Id + " " +
                                           "," + form.��������������.id_�������� + " ) ";

                            // ������� � ������.
                            builderUpdate.Append(insert);
                        }
                    }

                    //�������.Update(form.��������������);
                }

                if (form.��������������.FlagCardRepeet == true)
                {
                    ControlFlagRepeet cfr = new ControlFlagRepeet(form.��������������.id_��������, form.��������������.FlagCardRepeet);
                    bool flag = cfr.CompareRepet();

                    string queryUpdate = string.Empty;

                    string stest = form.��������������.��������������.ToShortDateString();

                    // ������ ��������� ������ ���� �� ������������, � ����� ������������.
                    if (flag == false && form.��������������.FlagCardRepeet == true)
                    {

                        queryUpdate = "UPDATE [��������] " +
                                    "SET [id_���������] = " + form.��������������.id_��������� + " " +
                                    ",[id_��������������] = " + form.��������������.id_�������������� + " " +
                                    ",[�����] = '" + form.��������������.����� + "' " +
                                    ",[����������] = '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ",[����������] = '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ",[�����������������] = '" + form.��������������.�����������������.Trim() + "' " +
                                    ",[����������] = '" + form.��������������.���������� + "' " +
                                    //",[���������] = '" + form.��������������.���������.Trim() + "' " +
                                     ",[���������] = '" + docNumNext.������� + "' " +
                                    ",[����������] = '" + form.��������������.���������� + "' " +
                                    ",[���������] = '" + form.��������������.���������.Trim() + "' " +
                                    ",[�������������������] = '" + form.��������������.�������������������.Trim() + "'  " +
                                    ",[��������������] = '" + ����SQL.����(form.��������������.��������������.ToShortDateString()) + "' " +
                                    //",[�������] = " + form.��������������.������� + " " +
                                    ",[�������] = " + docNumNext.����� + " " +
                                     //",[�������] = " + docNum.����� + " " +
                                    ",[����������������������] = '' " +
                                    ",[FlagPersonData] = '" + form.��������������.FlagPersonData + "' " +
                                    ",[FlagCardRepeet] = '" + form.��������������.FlagCardRepeet + "' " +
                                     ",[NameFileDocument] = '" + fileNameCopy + "' " +
                                     ",GuidName = '" + form.PathFileServer.Trim() + "' " +
                                      ",md5 = '" + strMd5.Trim() + "' " +
                                    "WHERE id_�������� = " + form.��������������.id_�������� + " " +
                                    " INSERT INTO [��������������] " +
                                    "([id_���������] " +
                                    ",[id_��������������] " +
                                    ",[�����] " +
                                    ",[����������] " +
                                    ",[����������] " +
                                    ",[�����������������] " +
                                    ",[����������] " +
                                    ",[���������] " +
                                    ",[����������] " +
                                    ",[���������] " +
                                    ",[�������������������] " +
                                    ",[��������������] " +
                                    ",[�������] " +
                                    ",[����������������������] " +
                                    ",[FlagPersonData] " +
                                    ",[FlagCardRepeet] " +
                                    ",[id_����������������] " +
                                    ",�������������� )" +
                                    "VALUES " +
                                    "(" + form.��������������.id_��������� + " " +
                                    ", " + form.��������������.id_�������������� + " " +
                                    ",'" + form.��������������.����� + "' " +
                                    ", '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ", '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ",'" + form.��������������.�����������������.Trim() + "' " +
                                    ",'" + form.��������������.���������� + "' " +
                                    ",'" + form.��������������.���������.Trim() + "' " +
                                    ",'" + form.��������������.���������� + "' " +
                                    ",'" + form.��������������.���������.Trim() + "' " +
                                    ",'" + form.��������������.�������������������.Trim() + "'  " +
                                    ",'" + ����SQL.����(form.��������������.��������������.ToShortDateString()) + "' " +
                                    //"," + form.��������������.������� + " " +
                                     ",[�������] = " + docNum.����� + " " +
                                    ",'' " +
                                    ",'" + form.��������������.FlagPersonData + "' " +
                                    ", '" + form.��������������.FlagCardRepeet + "' " +
                                    "," + form.��������������.id_�������� + " " +
                                    ","+ inc +") ";

                        builderUpdate.Append(queryUpdate);

                        //ExecuteQuery exe = new ExecuteQuery(builderUpdate.ToString().Trim());
                        //exe.Excecute();

                        
                    }

                    // ������ ��������� � ������������ ������.
                    if (flag == true && form.��������������.FlagCardRepeet == true)
                    {
                        queryUpdate = "UPDATE [��������] " +
                                    "SET [id_���������] = " + form.��������������.id_��������� + " " +
                                    ",[id_��������������] = " + form.��������������.id_�������������� + " " +
                                    ",[�����] = '" + form.��������������.����� + "' " +
                                    ",[����������] = '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ",[����������] = '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ",[�����������������] = '" + form.��������������.�����������������.Trim() + "' " +
                                    ",[����������] = '" + form.��������������.���������� + "' " +
                                    //",[���������] = '" + form.��������������.���������.Trim() + "' " +
                                     ",[���������] = '" + docNumNext.������� + "' " +
                                    ",[����������] = '" + form.��������������.���������� + "' " +
                                    ",[���������] = '" + form.��������������.���������.Trim() + "' " +
                                    ",[�������������������] = '" + form.��������������.�������������������.Trim() + "'  " +
                                    ",[��������������] = '" + ����SQL.����(form.��������������.��������������.ToShortDateString()) + "' " +
                                    //",[�������] = " + form.��������������.������� + " " +
                            //",[�������] = " + docNumNext.����� + " " +
                                     ",[�������] = " + docNum.����� + " " +
                                    ",[����������������������] = '' " +
                                    ",[FlagPersonData] = '" + form.��������������.FlagPersonData + "' " +
                                    ",[FlagCardRepeet] = '" + form.��������������.FlagCardRepeet + "' " +
                                    ",[NameFileDocument] = '" + fileNameCopy + "' " +
                                    ",GuidName = '" + form.PathFileServer.Trim() + "' " +
                                     ",md5 = '" + strMd5.Trim() + "' " +
                                    "WHERE id_�������� = " + form.��������������.id_�������� + " " +
                                    " DELETE FROM [��������������] " +
                                    "WHERE id_���������������� = " + form.��������������.id_�������� + " " +
                                    " INSERT INTO [��������������] " +
                                    "([id_���������] " +
                                    ",[id_��������������] " +
                                    ",[�����] " +
                                    ",[����������] " +
                                    ",[����������] " +
                                    ",[�����������������] " +
                                    ",[����������] " +
                                    ",[���������] " +
                                    ",[����������] " +
                                    ",[���������] " +
                                    ",[�������������������] " +
                                    ",[��������������] " +
                                    ",[�������] " +
                                    ",[����������������������] " +
                                    ",[FlagPersonData] " +
                                    ",[FlagCardRepeet] " +
                                    ",[id_����������������] " +
                                    ",��������������)" +
                                    "VALUES " +
                                    "(" + form.��������������.id_��������� + " " +
                                    ", " + form.��������������.id_�������������� + " " +
                                    ",'" + form.��������������.����� + "' " +
                                    ", '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ", '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                    ",'" + form.��������������.�����������������.Trim() + "' " +
                                    ",'" + form.��������������.���������� + "' " +
                                    ",'" + form.��������������.���������.Trim() + "' " +
                                    ",'" + form.��������������.���������� + "' " +
                                    ",'" + form.��������������.���������.Trim() + "' " +
                                    ",'" + form.��������������.�������������������.Trim() + "'  " +
                                    ",'" + ����SQL.����(form.��������������.��������������.ToShortDateString()) + "' " +
                                    //"," + form.��������������.������� + " " +
                                     ",[�������] = " + docNum.����� + " " +
                                    ",'' " +
                                    ",'" + form.��������������.FlagPersonData + "' " +
                                    ", '" + form.��������������.FlagCardRepeet + "' " +
                                    "," + form.��������������.id_�������� + " " + 
                                    ","+ inc +") ";

                        builderUpdate.Append(queryUpdate);

                        //ExecuteQuery exe = new ExecuteQuery(builderUpdate.ToString().Trim());
                        //exe.Excecute();

                    }

                 }

                // ������ �� ���������� ��������� �������� �������� ���������� ���.

                 builderUpdate.Append(form.QueryPersonDateForCardInput);

                 // �������� ����������.
                 builderUpdate.Append(" COMMIT TRANSACTION  ");

                 string sQueryTest = builderUpdate.ToString();


                 //// �������� ������.
                 ������������ connectBD = new ������������();
                 string sCon = connectBD.�����������������();

                // ���� ��������� ������� �� ���������� ����.
                 bool flagCopyServer = false;

                 using (SqlConnection con = new SqlConnection(sCon))
                 {
                     if (form.SaveDocServer == true)
                     {
                         try
                         {
                             if (File.Exists(fileNameCopy) == true)
                             {
                                 if (form.���������������� == true)
                                 {
                                     //string asd = patchServerFile + @"\Move\" + namFile;

                                     //File.Move(fileNameCopy, asd);

                                     File.Delete(fileNameCopy);

                                     FileInfo file = new FileInfo(fileName);
                                     file.CopyTo(fileNameCopy, true);


                                     string patchDir = Application.StartupPath + @"\Archive\";

                                     // ������ ��� ����� �� ����������
                                     DirectoryInfo dirInfo = new DirectoryInfo(patchDir);

                                     foreach (FileInfo fil in dirInfo.GetFiles())
                                     {
                                         fil.Delete();
                                     }
                                 }


                             }
                             else
                             {
                                 if (form.���������������� == true)
                                 {
                                     // ��������� ���� �� ������ �������� ����������.
                                     File.Copy(fileName, fileNameCopy, true);
                                 }
                             }
                            
                             // ���� ���� ������������ ������� ������� ���� � true.
                             flagCopyServer = true;
                         }
                         catch
                         {
                             MessageBox.Show("������ ��� ����������� �����");

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
                         // ���� ���� ������������ ������� ������� ���� � true.
                         flagCopyServer = true;

                         con.Open();
                         SqlCommand com = new SqlCommand(builderUpdate.ToString().Trim(), con);
                          com.ExecuteNonQuery();
                     }

                 }

                ��������������();
            }
        }

        /// <summary>
        /// ����������� ���� �������� ������ � ������� ��������� ���������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContext�����������������������_Click(object sender, EventArgs e)
        {
            // �������� ������ ������������ � ���������� ������:
            BindingManagerBase bmb = this.BindingContext[dataGrid������������������.DataSource, dataGrid������������������.DataMember];
            bmb.Position = dataGrid������������������.CurrentCell.RowNumber;
            dataGrid������������������.Select(dataGrid������������������.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            DataRow[] row = ds11.�����������������.Select("id_��������=" + (int)drv["id_��������"]);
            DS1.�����������������Row ������������������ = (DS1.�����������������Row)row[0];
            
            Form����������������� form = new Form�����������������(ds11, ������������������, ������������);

            // ������ ��� ����������� ��������.
            form.FlagEdit = true;
            
            // ��������� � ����� id �������� ���������.
            form.Id���������������� = ������������������.id_��������;

            form.ShowDialog(this);
            if (form.DialogResult == DialogResult.OK)
            {
                DS1TableAdapters.�����������������TableAdapter ������� = new RegKor.DS1TableAdapters.�����������������TableAdapter();
                //�������.Update(form.�����������������������);

                DataRow row2 = form.�����������������������;

                // ������� �������� ��������� id ��� ��������� �� �������� ��� ������� � �������� ����������.
                List<int> listId������� = form.ListID��������;
                List<�����������������> listOP = form.List�����������������;

                List<int> ListID���������������������������������� = form.ListID����������������������������������;
                List<int>ListID��������������������������������������� = form.ListID���������������������������������������;

                StringBuilder builder = new StringBuilder();

                int id_������������������;
                if (row2["id_������������������"] == DBNull.Value)
                {
                    string query = string.Empty;

                    // �������� �������� � ��� ��� ���.
                    if (Convert.ToBoolean(form.FlagDsp) == false)
                    {
                        // ������� ��������� ���������.
                        query = "UPDATE [�����������������] " +
                                       "SET [����] = '" + ����SQL.����(Convert.ToDateTime(row2["����"]).ToShortDateString().Trim()) + "' " +
                                       ",[�������������] = '" + row2["�������������"] + "' " +
                                       ",[id_�������������] = " + row2["id_�������������"] + " " +
                                       ",[�������������������] = '" + row2["�������������������"] + "' " +
                                       ",[���������������] = " + row2["���������������"] + " " +
                                       ",[id_��������] = " + row2["id_��������"] + " " +
                                       ",[����������] = '" + row2["����������"] + "' " +
                            //",[id_������������������] = " + id_������������������ + " " +
                                       ",[����������������������] = '" + row2["����������������������"] + "' " +
                                       ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                                       ",��� = '" + form.FlagDsp + "' " +
                                       "where id_�������� = " + Convert.ToInt32(row2["id_��������"]) + " ";
                    }
                    else
                    {
                        // ������� ��������� ���������.
                        query = "UPDATE [�����������������] " +
                                       "SET [����] = '" + ����SQL.����(Convert.ToDateTime(row2["����"]).ToShortDateString().Trim()) + "' " +
                                       ",[�������������] = '" + row2["�������������"] + "' " +
                                       ",[id_�������������] = " + row2["id_�������������"] + " " +
                                       ",[�������������������] = '" + row2["�������������������"] + "' " +
                                       ",[���������������] = " + row2["���������������"] + " " +
                                       ",[id_��������] = " + row2["id_��������"] + " " +
                                       ",[����������] = '" + row2["����������"] + "' " +
                            //",[id_������������������] = " + id_������������������ + " " +
                                       ",[����������������������] = '" + row2["����������������������"] + "' " +
                                       ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                                       ",��� = '" + form.FlagDsp + "' " +
                                       ", ���Desc = '���' " +
                                       "where id_�������� = " + Convert.ToInt32(row2["id_��������"]) + " ";
                    }

                    builder.Append(query);

                    // ������� ��� �����.
                    //int iCountlistOP = 0;
                    string queryDelete = "DELETE FROM [���������������������������������������] " +
                                        "WHERE id_�������� = " + Convert.ToInt32(row2["id_��������"]) + " ";

                    builder.Append(queryDelete);

                    // ������� ������ �� ��������� ��������.
                    // ������ �� ������� id � ��������� ������� �����������������������.
                    string sUpdate = string.Empty;
                    sUpdate = String.Format(form.QueryInsert.Trim(), " "+ Convert.ToInt32(row2["id_��������"]) + " ");

                    builder.Append(sUpdate);

                  

                    // ������ ������ �� ����������� �������.
                    string delete = "DELETE FROM ���������������������������������� " +
                                    "WHERE id_����������������� = " + Convert.ToInt32(row2["id_��������"]) + " ";

                    builder.Append(delete);

                    // �������� ��������� ������� �������� �����������������.
                    foreach (int id�� in listId�������)
                    {
                        string queryId�� = "INSERT INTO [����������������������������������] " +
                                           "([id_����������������] " +
                                           ",[id_�����������������]) " +
                                           "VALUES " +
                                           "(" + id�� + " " +
                                           "," + Convert.ToInt32(row2["id_��������"]) + " ) " +
                                           "update �������� " +
                                           "set ������������������� = '" + form.��������������.Trim() + "' " +
                                           "where id_�������� = " + id�� + " ";

                        //iiCount++;

                        builder.Append(queryId��);
                    }

                }
                else
                {
                    id_������������������ = Convert.ToInt32(row2["id_������������������"]);

                    string query = string.Empty;

                    //// ������� ��������� ���������.
                    //string query = "UPDATE [�����������������] " +
                    //                "SET [����] = '" + ����SQL.����(Convert.ToDateTime(row2["����"]).ToShortDateString().Trim()) + "' " +
                    //                ",[�������������] = '" + row2["�������������"] + "' " +
                    //                ",[id_�������������] = " + row2["id_�������������"] + " " +
                    //                ",[�������������������] = '" + row2["�������������������"] + "' " +
                    //                ",[���������������] = " + row2["���������������"] + " " +
                    //                ",[id_��������] = " + row2["id_��������"] + " " +
                    //                ",[����������] = '" + row2["����������"] + "' " +
                    //                ",[id_������������������] = " + id_������������������ + " " +
                    //                ",[����������������������] = '" + row2["����������������������"] + "' " +
                    //                ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                    //                 ",��� = '" + form.FlagDsp + "' " +
                    //                "where id_�������� = " + Convert.ToInt32(row2["id_��������"]) + " ";

                    // �������� �������� � ��� ��� ���.
                    if (Convert.ToBoolean(form.FlagDsp) == false)
                    {
                        // ������� ��������� ���������.
                        query = "UPDATE [�����������������] " +
                                       "SET [����] = '" + ����SQL.����(Convert.ToDateTime(row2["����"]).ToShortDateString().Trim()) + "' " +
                                       ",[�������������] = '" + row2["�������������"] + "' " +
                                       ",[id_�������������] = " + row2["id_�������������"] + " " +
                                       ",[�������������������] = '" + row2["�������������������"] + "' " +
                                       ",[���������������] = " + row2["���������������"] + " " +
                                       ",[id_��������] = " + row2["id_��������"] + " " +
                                       ",[����������] = '" + row2["����������"] + "' " +
                            ",[id_������������������] = " + id_������������������ + " " +
                                       ",[����������������������] = '" + row2["����������������������"] + "' " +
                                       ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                                       ",��� = '" + form.FlagDsp + "' " +
                                       "where id_�������� = " + Convert.ToInt32(row2["id_��������"]) + " ";
                    }
                    else
                    {
                        // ������� ��������� ���������.
                        query = "UPDATE [�����������������] " +
                                       "SET [����] = '" + ����SQL.����(Convert.ToDateTime(row2["����"]).ToShortDateString().Trim()) + "' " +
                                       ",[�������������] = '" + row2["�������������"] + "' " +
                                       ",[id_�������������] = " + row2["id_�������������"] + " " +
                                       ",[�������������������] = '" + row2["�������������������"] + "' " +
                                       ",[���������������] = " + row2["���������������"] + " " +
                                       ",[id_��������] = " + row2["id_��������"] + " " +
                                       ",[����������] = '" + row2["����������"] + "' " +
                            ",[id_������������������] = " + id_������������������ + " " +
                                       ",[����������������������] = '" + row2["����������������������"] + "' " +
                                       ",[FlagPersonData] = '" + row2["FlagPersonData"] + "' " +
                                       ",��� = '" + form.FlagDsp + "' " +
                                       ", ���Desc = '���' " +
                                       "where id_�������� = " + Convert.ToInt32(row2["id_��������"]) + " ";
                    }

                    builder.Append(query);

                    string updateQuery = "UPDATE [��������] " +
                                         "SET [�����] = 'True' " +
                                         "WHERE id_�������� = "+ id_������������������ +" ";

                    builder.Append(updateQuery);

                    // ������� ��� �����.
                    //int iCountlistOP = 0;
                    string queryDelete = "DELETE FROM [���������������������������������������] " +
                                       "WHERE id_�������� = " + Convert.ToInt32(row2["id_��������"]) + " ";

                    builder.Append(queryDelete);

                    // ������� ������ �� ��������� ��������.
                    // ������ �� ������� id � ��������� ������� �����������������������.
                    foreach (����������������� itm in listOP)
                    {

                        string queryIns = "INSERT INTO [���������������������������������������] " +
                                       "([id_��������] " +
                                       ",[id_�����������������]) " +
                                       "VALUES " +
                                       "('" + Convert.ToInt32(row2["id_��������"]) + "' " +
                                       ",'" + itm.Id_����������������� + "' ) ";


                        //iCountlistOP++;

                        builder.Append(queryIns);
                    }

                    //int iiCount = 0;

                    // ������ ������ �� ����������� �������.
                    string delete = "DELETE FROM ���������������������������������� " +
                                    "WHERE id_����������������� = " + Convert.ToInt32(row2["id_��������"]) + " ";
                    
                    builder.Append(delete);

                    // �������� ��������� ������� �������� �����������������.
                    foreach (int id�� in listId�������)
                    {
                        string queryId�� = "INSERT INTO [����������������������������������] " +
                                           "([id_����������������] " +
                                           ",[id_�����������������]) " +
                                           "VALUES " +
                                           "(" + id�� + " " +
                                           "," + Convert.ToInt32(row2["id_��������"]) + " ) ";

                        //iiCount++;

                        builder.Append(queryId��);
                    }
                }

                // ���� ������ ��� �������� ��� � ��������� ������������ ������.
                if (Convert.ToBoolean(row2["FlagPersonData"]) == true)
                {

                    string ������������������ = string.Empty;

                    DataRow[] rowsSelect = ds11.���������������������.Select("id_�������������= " + Convert.ToInt32(row2["id_�������������"]) + " ");
                    foreach (DataRow item in rowsSelect)
                    {
                        ������������������ = item["������������������"].ToString().Trim();
                    }

                    string ������������������� = "��� �����. � ���. ��������� " + row2["�������������"].ToString().Trim() + "-" + row2["�������������������"].ToString().Trim() + "-" + ������������������ + "/" + row2["���������������"].ToString().Trim();

                    if (row2["id_������������������"] != DBNull.Value)
                    {
                        // ��������� ���� � TRUE.
                        string queryUpdate = " UPDATE [��������] " +
                                             "SET ������������������� = '" + ������������������� + "' " +
                            //"FlagPersonData = '" + row["FlagPersonData"] + "' " +
                                             ",����� = 'True' " +
                                              ",��� = '" + form.FlagDsp + "' " +
                                             "where id_�������� = " + Convert.ToInt32(row2["id_������������������"]) + " ";
                        // ������ ������ ��������� �� ���������� ������ � �� �������������� � ������ ������, ����� ��������� �� � ����� ����������.
                        builder.Append(queryUpdate);
                    }
                }
                else
                {
                    string ������������������ = string.Empty;

                    DataRow[] rowsSelect = ds11.���������������������.Select("id_�������������= " + Convert.ToInt32(row2["id_�������������"]) + " ");
                    foreach (DataRow item in rowsSelect)
                    {
                        ������������������ = item["������������������"].ToString().Trim();
                    }

                    string ������������������� = "��� �����. � ���. ��������� " + row2["�������������"].ToString().Trim() + "-" + row2["�������������������"].ToString().Trim() + "-" + ������������������ + "/" + row2["���������������"].ToString().Trim();

                    if (row2["id_������������������"] != DBNull.Value)
                    {
                        // ��������� ���� � TRUE.
                        string queryUpdate = " UPDATE [��������] " +
                                             "SET ������������������� = '" + ������������������� + "' " +
                            //"FlagPersonData = '" + row["FlagPersonData"] + "' " +
                                             ",����� = 'True' " +
                                              ",��� = '" + form.FlagDsp + "' " +
                                             "where id_�������� = " + Convert.ToInt32(row2["id_������������������"]) + " ";
                        // ������ ������ ��������� �� ���������� ������ � �� �������������� � ������ ������, ����� ��������� �� � ����� ����������.
                        builder.Append(queryUpdate);
                    }
                }

              
                
                // �������� ��������� � ��.
                ������������ ������� = new ������������();
                string sConnect = �������.�����������������();

                SqlConnection con = new SqlConnection(sConnect);
                SqlCommand com = new SqlCommand(builder.ToString(), con);
                con.Open();
                com.ExecuteNonQuery();
                con.Close();
                    
                                   
                ��������������();
            }
        }

        /// <summary>
        /// ����������� ���� �������� ������ � ������� ��������� � ����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContext��������������_Click2(object sender, EventArgs e)
        {
            // �������� ������ ������������ � ���������� ������:
            BindingManagerBase bmb = this.BindingContext[dataGrid��������������.DataSource, dataGrid��������������.DataMember];
            bmb.Position = dataGrid��������������.CurrentCell.RowNumber;
            dataGrid��������������.Select(dataGrid��������������.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            int id�������� = (int)drv["id_��������"];
            Form�������� form = new Form��������(ds11, id��������, ������������);
            form.ShowDialog(this);
            if (form.DialogResult == DialogResult.OK)
            {

                DS1TableAdapters.��������TableAdapter ������� = new RegKor.DS1TableAdapters.��������TableAdapter();

                // ����� ��� �������� ������ ��������� � �� ��� ���������� ��.
                 StringBuilder builderUpdate = new StringBuilder();

                // ������� ������ ����������� ���������.
                Item�������������������������� �������������������� = form.�����������������;

                // �������� ������ ������� � ������ ����������.
                builderUpdate.Append("SET TRANSACTION ISOLATION LEVEL serializable begin transaction  ");

                string queryUpdate = "UPDATE [��������] " +
                                     "SET [id_���������] = " + form.��������������.id_��������� + " " +
                                     ",[id_��������������] = " + form.��������������.id_�������������� + " " +
                                     ",[�����] = '" + form.��������������.����� + "' " +
                                     ",[����������] = '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                     ",[����������] = '" + ����SQL.����(form.��������������.����������.ToShortDateString()) + "' " +
                                     ",[�����������������] = '" + form.��������������.�����������������.Trim() + "' " +
                                     ",[����������] = '" + form.��������������.���������� + "' " +
                    //",[���������] = '" + form.��������������.���������.Trim() + "' " +
                    //",[���������] = '" + docNumNext.������� + "' " +
                                     ",[����������] = '" + form.��������������.���������� + "' " +
                                     ",[���������] = '" + form.��������������.���������.Trim() + "' " +
                                     ",[�������������������] = '" + form.��������������.�������������������.Trim() + "'  " +
                                     ",[��������������] = '" + ����SQL.����(form.��������������.��������������.ToShortDateString()) + "' " +
                    //",[�������] = " + form.��������������.������� + " " +
                    //",[�������] = " + docNumNext.����� + " " +
                    //",[�������] = " + docNum.����� + " " +
                                     ",[����������������������] = '' " +
                                     ",[FlagPersonData] = '" + form.��������������.FlagPersonData + "' " +
                                     ",[FlagCardRepeet] = '" + form.��������������.FlagCardRepeet + "' " +
                                      ",[NameFileDocument] = '" + fileNameCopy + "' " +
                                      ",GuidName = '" + form.PathFileServer.Trim() + "' " +
                                      ",id����������������������� = " + form.�����������������.Id + " " +
                    // ",md5 = '" + strMd5.Trim() + "' " +
                                     "WHERE id_�������� = " + form.��������������.id_�������� + " ";

                // ������� ������ ������� �� ����������.
                builderUpdate.Append(queryUpdate);

                // �������� ������ ������ ��� ���.


                // ������ �� �������� ��������� � ���������� ����� �������.
                builderUpdate.Append(form.QueryPersonDateForCardInput);

                builderUpdate.Append("COMMIT TRANSACTION ");

                // ������ �� ���������� �������.
                string queryUpdateCard = builderUpdate.ToString();


                ExecuteQuery exec = new ExecuteQuery(queryUpdateCard);
                exec.Excecute();

                //�������.Update(form.��������������);
                ��������������();
            }
            this.Refresh();
        }

        /// <summary>
        /// ����������� ���� ������� ������ � ������� ������� ���������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContext�������������_Click(object sender, EventArgs e)
        {
            // �������� ������ ������������ � ���������� ������:
            BindingManagerBase bmb = this.BindingContext[dataGrid����������������.DataSource, dataGrid����������������.DataMember];
            bmb.Position = dataGrid����������������.CurrentCell.RowNumber;
            dataGrid����������������.Select(dataGrid����������������.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            DialogResult ����������������� = MessageBox.Show(this, "�� ������������� ������ ������� �������� '" + drv["�����������������"] + "' �� ������������� '" + drv["����������������������"] + "'?", "�������� ������", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2);
            if (����������������� == DialogResult.Yes)
            {
                int id�������� = (int)drv["id_��������"];

                DataRow[] rows = ds11.��������.Select("id_�������� = " + id��������);
                //rows[0].Delete();

                StringBuilder build = new StringBuilder();

                string queryDelete = "DELETE FROM [��������] " +
                                     "WHERE id_�������� = "+ id�������� +" ";

                build.Append(queryDelete);

                string queryDeleteD = "DELETE FROM �������������� " +
                                     "WHERE id_���������������� = " + id�������� + " ";

                build.Append(queryDeleteD);

                 //� ��� ����� ����� � �������������� ����� id �� ���� ��������

                ExecuteQuery eq = new ExecuteQuery(build.ToString().Trim());
                eq.Excecute();

                ��������������();
            }
            this.Refresh();
        }

        /// <summary>
        /// ����������� ���� ������� ������ � ������� ��������� � ����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContext�������������_Click2(object sender, EventArgs e)
        {
            // �������� ������ ������������ � ���������� ������:
            BindingManagerBase bmb = this.BindingContext[dataGrid��������������.DataSource, dataGrid��������������.DataMember];
            bmb.Position = dataGrid��������������.CurrentCell.RowNumber;
            dataGrid��������������.Select(dataGrid��������������.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            DialogResult ����������������� = MessageBox.Show(this, "�� ������������� ������ ������� �������� '" + drv["�����������������"] + "' �� ������������� '" + drv["����������������������"] + "'?", "�������� ������", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2);
            if (����������������� == DialogResult.Yes)
            {
                int id�������� = (int)drv["id_��������"];

                DataRow[] rows = ds11.��������.Select("id_�������� = " + id��������);
                rows[0].Delete();

                ��������������();
            }
            this.Refresh();
        }

        /// <summary>
        /// ����������� ���� ������� ������ � ������� ��������� ���������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItemContext����������������������_Click(object sender, EventArgs e)
        {
            // �������� ������ ������������ � ���������� ������:
            BindingManagerBase bmb = this.BindingContext[dataGrid������������������.DataSource, dataGrid������������������.DataMember];
            bmb.Position = dataGrid������������������.CurrentCell.RowNumber;
            dataGrid������������������.Select(dataGrid������������������.CurrentCell.RowNumber);
            DataRowView drv = (DataRowView)bmb.Current;
            DialogResult ����������������� = MessageBox.Show(this, "�� ������������� ������ ������� �������� �� '" + drv["���������������������"] + "' \n���\n'" + drv["����������������"] + "'?", "�������� ������", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button2);
            if (����������������� == DialogResult.Yes)
            {
                int id�������� = (int)drv["id_��������"];

                if (drv["id_������������������"] != System.DBNull.Value)
                {
                    int id������������� = (int)drv["id_������������������"];
                    DataRow[] ������ = ds11.��������.Select("id_��������=" + id�������������);
                    if (������.Length > 0)
                    {
                        ������[0]["�����"] = false;
                        ������[0]["�������������������"] = "";
                    }
                }


                DataRow[] rows = ds11.�����������������.Select("id_�������� = " + id��������);
                rows[0].Delete();

                string query = "DELETE FROM ����������������� " +
                               "WHERE id_�������� = "+ id�������� +" ";

                Classess.������������ clConn = new ������������();
                string sConn = clConn.�����������������();

                SqlConnection con = new SqlConnection(sConn);
                con.Open();
                SqlCommand com = new SqlCommand(query, con);
                com.ExecuteNonQuery();
                con.Close();

                ��������������();
            }
            this.Refresh();
        }

        private void menuItemContext��������������_Click(object sender, EventArgs e)
        {
            ��������������();
            this.Refresh();
        }

        /// <summary>
        /// �������� ���� ���������� � ����
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem�������_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// �������� ���������� "���������"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem��������������������_Click(object sender, System.EventArgs e)
        {
            Form��������� form = new Form���������();
            this.Enabled = false;
            form.ShowDialog(this);
            this.Refresh();
            ��������������������������();
            this.Enabled = true;
            this.Refresh();
        }

        /// <summary>
        /// �������� ���������� "��������������"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem�������������������������_Click(object sender, System.EventArgs e)
        {
            Form�������������� form = new Form��������������();
            this.Refresh();
            this.Enabled = false;
            form.ShowDialog(this);
            this.Refresh();
            ��������������������������();
            this.Enabled = true;
            this.Refresh();
        }

        /// <summary>
        /// �������� ���������� "����������"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem���������������������_Click(object sender, System.EventArgs e)
        {
            Form���������� form = new Form����������();
            this.Refresh();
            this.Enabled = false;
            form.ShowDialog(this);
            this.Refresh();
            ��������������������������();
            this.Enabled = true;
            this.Refresh();
        }

        ///// <summary>
        ///// �������� ���������� "��������"
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void menuItem������������������_Click ( object sender, EventArgs e )
        //{
        //    Form����������������� form = new Form�����������������();
        //    this.Refresh( );
        //    this.Enabled = false;
        //    form.ShowDialog( this );
        //    this.Refresh( );
        //    ��������������������������( );
        //    this.Enabled = true;
        //    this.Refresh( );	
        //}

        private void menuItem4_Click(object sender, System.EventArgs e)
        {
            this.Enabled = false;
            Form����������� frm = new Form�����������(this.ds11);
            frm.ShowDialog(this);
            this.Enabled = true;
        }

        /// <summary>
        /// ������� Click ���� "���������� ����������� ���������������"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem5_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            Form�����������2 frm = new Form�����������2(this.ds11);
            frm.ShowDialog(this);
            this.Enabled = true;
        }

        private void menuItem����������������������_Click(object sender, EventArgs e)
        {
            ����������������������������();
        }

        private void menuItem����������������_Click(object sender, EventArgs e)
        {
            ������������� = new System.Threading.Thread(new System.Threading.ThreadStart(����������������������));
            �������������.Start();

            ����������������������������();
        }

        private void tabControl��������������_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl��������������.SelectedTab.Text == "���������")
            {
                //MessageBox.Show( "�������� ������� " + tabControl��������������.SelectedTab.Name);
                menuItem����������������������.Enabled = false;
                menuItem����������������.Enabled = false;

                menuItemContext��������������.Enabled = false;
                menuItem4.Enabled = false;
                menuItem6.Enabled = true;
            }
            if (tabControl��������������.SelectedTab.Text == "��������")
            {
                menuItem����������������������.Enabled = true;
                menuItem����������������.Enabled = true;
                
                menuItemContext��������������.Enabled = true;
                menuItem4.Enabled = true;
                menuItem6.Enabled = false;
            }
        }

        #endregion

        private void menuItem������������������������_Click(object sender, EventArgs e)
        {
            Form������������� form = new Form�������������(this.ds11);
            this.Refresh();
            this.Enabled = false;
            form.ShowDialog(this);
            this.Refresh();
            ��������������������������();
            this.Enabled = true;
            this.Refresh();
        }

        private void dataGrid����������������_Resize(object sender, EventArgs e)
        {
            int ������������� = dataGrid����������������.Width;

            int ���� = dataGridTextBoxColumn1.Width;
            int ����� = dataGridTextBoxColumn2.Width;
            int ���1 = dataGridTextBoxColumn3.Width;
            int ����1 = dataGridTextBoxColumn4.Width;
            int ���2 = dataGridTextBoxColumn5.Width;
            int ����2 = dataGridTextBoxColumn6.Width;
            int ������� = dataGridTextBoxColumn7.Width;
            int ������ = dataGridTextBoxColumn8.Width;

            dataGridTextBoxColumn7.Width = ������������� - 20 - ���� - ����� - ���1 - ����1 - ���2 - ����2 - ������;
        }

        private void dataGrid��������������_Resize(object sender, EventArgs e)
        {
            int ������������� = dataGrid��������������.Width;

            int ���� = dataGridTextBoxColumn9.Width;
            int ����� = dataGridTextBoxColumn10.Width;
            int ���1 = dataGridTextBoxColumn11.Width;
            int ����1 = dataGridTextBoxColumn12.Width;
            int ���2 = dataGridTextBoxColumn13.Width;
            int ����2 = dataGridTextBoxColumn14.Width;
            int ������� = dataGridTextBoxColumn15.Width;
            int ������ = dataGridTextBoxColumn16.Width;

            dataGridTextBoxColumn15.Width = ������������� - 20 - ���� - ����� - ���1 - ����1 - ���2 - ����2 - ������;
        }

        private void dataGrid������������������_Resize(object sender, EventArgs e)
        {
            int ������������� = dataGrid������������������.Width;

            int ����� = dataGridTextBoxColumn����������������.Width;
            int ������ = dataGridTextBoxColumn�����������.Width;
            int ������ = dataGridTextBoxColumn����������������������.Width;
            int ������� = dataGridTextBoxColumn����������������.Width;
            int ������� = dataGridTextBoxColumn��������������������.Width;

            dataGridTextBoxColumn����������������.Width = ������������� - 20 - ����� - ������ - ������ - �������;
        }



        /// <summary>
        /// ������������ ������ ��� ���������� ��������� ����������
        /// </summary>
        private string ��������
        {
            get
            {
                string ������ = string.Empty;

                if (textBox�������������������������������.Text.Trim().ToLower() != "���".ToLower().Trim())
                {
                    ������ = "(�������������� LIKE '%" + textBox�������������������������������.Text + "%'" +
                                                        " OR ���������� LIKE '%" + textBox�������������������������������.Text + "%'" +
                                                        " OR ��������������������� LIKE '%" + textBox�������������������������������.Text + "%'" +
                                                        " OR �������������������� LIKE '%" + textBox�������������������������������.Text + "%'" +
                                                        " OR ���������������� LIKE '%" + textBox�������������������������������.Text + "%')";
                }
                else
                {
                    ������ = "���Desc = '���' ";
                }
                if (comboBox��������������.SelectedItem.ToString() == "���� ���")
                {
                    DateTime min = Convert.ToDateTime("01.12." + ������������ + "");
                    DateTime max = Convert.ToDateTime("31.12." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "������")
                {
                    DateTime min = Convert.ToDateTime("01.12." + ������������ + "");
                    DateTime max = Convert.ToDateTime("31.01." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "�������")
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
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "����")
                {
                    DateTime min = Convert.ToDateTime("01.03." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.03." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "������")
                {
                    DateTime min = Convert.ToDateTime("01.04." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("30.04." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "���")
                {
                    DateTime min = Convert.ToDateTime("01.05." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.05." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "����")
                {
                    DateTime min = Convert.ToDateTime("01.06." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("30.06." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "����")
                {
                    DateTime min = Convert.ToDateTime("01.07." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.07." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "������")
                {
                    DateTime min = Convert.ToDateTime("01.08." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.08." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "��������")
                {
                    DateTime min = Convert.ToDateTime("01.09." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("30.09." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "�������")
                {
                    DateTime min = Convert.ToDateTime("01.10." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.10." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "������")
                {
                    DateTime min = Convert.ToDateTime("01.11." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("30.11." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }
                if (comboBox��������������.SelectedItem.ToString() == "�������")
                {
                    DateTime min = Convert.ToDateTime("01.12." + selectedYear.ToString() + "");
                    DateTime max = Convert.ToDateTime("31.12." + selectedYear.ToString() + "");
                    ������ += " AND ����>='" + min + "' AND ����<='" + max + "'";
                }

                return ������;
            }
        }

        private void comboBox��������������_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataView view = (DataView)dataGrid������������������.DataSource;
            view.RowFilter = ��������;
            label��������������������������������������������.Text = "�������� ����������: " + view.Count;
        }

        //private void dataGrid����������������_Navigate(object sender, NavigateEventArgs ne)
        //{

        //}

        //private void dataGrid������������������_Navigate(object sender, NavigateEventArgs ne)
        //{

        //}

        

        private void checkBox��������������_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox��������������.Checked == true)
            {
                this.comboBox��������������.Visible = false;
            }
            else
            {
                this.comboBox��������������.Visible = true;
            }
        }

        private void menuItem6_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            //Form������������������� frm = new Form�������������������(this.ds11);
            Form������������������� frm = new Form�������������������();
            frm.ShowDialog(this);
            this.Enabled = true;
        }

        private void menuItem8_Click(object sender, EventArgs e)
        {
            Form��������������� �������������� = new Form���������������();
            ��������������.ShowDialog(this);
        }

        private void Form�������_Load(object sender, EventArgs e)
        {
            //��� �������� ������� ����� ���� � �������� �� ��������
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
            // ������� ���� �������������� ������������ ������.
            FormPD formPD = new FormPD();
            formPD.Show();
        }

        private void menuItem12_Click(object sender, EventArgs e)
        {
            Form��������������������������� formPD = new Form���������������������������();
            formPD.Show();
        }

        private void menuItem13_Click(object sender, EventArgs e)
        {
            Form�������������� form = new Form��������������();
            form.���������� = ����.����������(selectedYear.ToString());
            form.���������� = �������������;
            form.ShowDialog(this);
            this.Enabled = true;
        }

        private void Form�������_FormClosing(object sender, FormClosingEventArgs e)
        {
            // ������� ���������� ������.

            // ������� ���� � ������������ �����.
            string pathExe = Application.StartupPath;

            string patch = pathExe;
        }

        private void contextMenu1_Popup(object sender, EventArgs e)
        {

        }

        private void menuItem15_Click(object sender, EventArgs e)
        {
            Form��������������� form = new Form���������������();
            form.Show();
        }

        private NumOutputCardVipNet GetNumDocOutVipNet(string strGuid)
        {

            // ��������� ���������������� ������.
            NumOutputCardVipNet numCard = new NumOutputCardVipNet();

            string num = string.Empty;

            string query = "select id_��������,��������������� from dbo.����������������� " +
                           "where [GUID] = '" + strGuid.Trim() + "' ";

            ������������ strCon = new ������������();
            SqlConnection con = new SqlConnection(strCon.�����������������());
            con.Open();

            SqlCommand com = new SqlCommand(query, con);
            SqlDataReader read = com.ExecuteReader();

            while (read.Read())
            {
                numCard.Id = Convert.ToInt32(read["id_��������"]);
                numCard.��������������� = read["���������������"].ToString().Trim();

            }

            return numCard;

        }

        /// <summary>
        /// ������� ����� ���������� ���������.
        /// </summary>
        /// <returns></returns>
        private string GetNumDocOut(string strGuid)
        {
            
            // ��������� ���������������� ������.
            NumOutputCardVipNet numCard = new NumOutputCardVipNet();
            
            string num = string.Empty;

            string query = "select id_��������,��������������� from dbo.����������������� " +
                           "where [GUID] = '" + strGuid.Trim() + "' ";

            ������������ strCon = new ������������();
            SqlConnection con = new SqlConnection(strCon.�����������������());
            con.Open();

            SqlCommand com = new SqlCommand(query, con);
            SqlDataReader read = com.ExecuteReader();

            while (read.Read())
            {
                num = read["���������������"].ToString().Trim();
            }

            return num.Trim();

        }

        private void menuItem16_Click(object sender, EventArgs e)
        {
            Form������������������ form = new Form������������������();
            form.YearSelect = ������������;
            form.Show();
        }

        /// <summary>
        /// ��������� ������� ��������.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem18_Click(object sender, EventArgs e)
        {
            string iTest = ������������;

            int seletYear = Convert.ToInt16(this.������������) + 1;

            string filePatchLog = Application.StartupPath + @"\fileLog.txt";

            if (File.Exists(filePatchLog) == true)
            {
                File.Delete(filePatchLog);
                Log.WriteLine(filePatchLog, "�������� ���");
            }
            else
            {
                Log.WriteLine(filePatchLog, "������� ��� ����");
            }


            // ���������� ��� �������� ���������� ���.
            int inc = 0;
            StringBuilder builder = new StringBuilder();

            // ������ ��� �������� ������� � �� ��� ��������� id �����������.
            StringBuilder build��� = new StringBuilder();

            Form�������� form = new Form��������(ds11, seletYear.ToString(), true);

            form.ShowDialog(this);

            // �������� �����.-
            �������������� docNumNext = form.�����������������������;

            if (form.DialogResult == DialogResult.OK)
            {
                DS1.��������Row row = form.��������������;

                // ����������� ���� ��� ������������� ���������.
                Guid guidCard = Guid.NewGuid();

                inc = form.IncrementDate;

                string patchToServer = string.Empty;

                // ���������� ��� �������� ����� ����� �� �������.
                string namFileServer = string.Empty;

                // ������� ��������� ������ ����������� ��������� ������ ������������.
                Item�������������������������� �������������������������� = form.�����������������;

                // �������� ������ ����������� ���������� � ������� ������� ������� ��������.
                this.ListPerson = form.ListPerson;

                // ���������� ����.

                //���� ���������� ���� ���������� ���������� ��������� �� �������.
                if (form.SaveDocServer == true)
                {
                    if (form.���������������� == true)
                    {
                        // ������� ���� � �����.
                        string filePatch = form.PathFileServer;

                        // ��� ��������� ����������.
                        //string archiver = @"C:\Program Files\7-Zip\7z.exe";

                        // ������� ��� ����� ������� ����� ��������������.
                        string archive = form.FileName;// +@"\*.*";

                        // GUID ������������ �������� �����.
                        string file = form.PathFileServer;

                        // �������� ��� ��� ����� ������ ����������� ������������ �����.
                        string namFileS = docNumNext.�����.ToString() + "-" + docNumNext.������� + "_" + file;
                        string namFile = docNumNext.�����.ToString() + "-" + docNumNext.�������;

                        // ���� � ���������� ���������� ����� � �������.
                        string patch = Application.StartupPath + @"\Archive\" + namFile + ".7z";

                        fileName = patch;

                        namFileServer = namFile;// +".7z";

                        // ��������� ���� ����� ������������ ����.
                        string patchDir = Application.StartupPath + @"\Archive\";

                        // ���������� �����. (������ ����������)
                        //Archiver.AddToArchive(archiver, archive, patch,patchDir);

                        Log.WriteLine(filePatchLog, "������������� ����� ������");

                        // ���� � 7z.dll.
                        string sevenZipDll = Application.StartupPath + @"\7z.dll";
                        if (archive.Length > 0)
                        {
                            // �������, ��� �������� ������� ��� ������ �� ������.
                            flagInsertCopyDoc = true;

                            Archiver.AddToArchive(sevenZipDll, archive, patch, patchDir);
                        }
                        else
                        {
                            // �������, ��� �������� ��� ������ �� ������ �� ��������.
                            flagInsertCopyDoc = false;

                            MessageBox.Show("�� �� ������� ����� ����� � ����������� �������� �� ������", "��������", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }


                        Log.WriteLine(filePatchLog, "������������� ����� �����");

                        // ���� ���� ����� ������������ �����.
                        patchToServer = patchServerFile + @"\" + namFileS.Trim();

                        fileNameCopy = patchToServer;
                    }
                    else
                    {
                        return;
                    }
                }

                // ===Begin========������� ������� ���� �������� �������� � ���� ������.
                // �������� ������ �� �������. (������ ,)
                string[] s����s = row["���������"].ToString().Split(',');

                int id_�������� = Convert.ToInt32(row["id_��������"]);

                // �������� ����� ������.
                DateTime todoy = DateTime.Now;

                // ������� ������.
                int iCount = 1;

                //// ���������� ������ 
                //foreach (string str in s����s)
                //{
                //    string insert = "declare @id_" + iCount + "  int " +
                //                    "SELECT @id_" + iCount + " = id_���������� " +
                //                    "FROM [����������] " +
                //                    "where [������������������] = '" + str.Trim() + "' " +
                //                    "INSERT INTO [�����������������������������] " +
                //                               "([id����������] " +
                //                               ",[���������������] " +
                //                               ",[����������������] " +
                //                               ",[�����������������] " +
                //                               ",[id��������] " +
                //                               ",[�������������������]) " +
                //                         "VALUES " +
                //                               "(@id_" + iCount + " " +
                //                               ",'" + todoy + "' " +
                //                               ",NULL " +
                //                               ",NULL " +
                //                               ","+ id_�������� +" " +
                //                               ",NULL) ";

                //    // ������� � ������.
                //    builder.Append(insert);

                //    iCount++;
                //}
                //=============End=====================

                // ���� ������������ �� ������ �������� ������� ����� ������������ � ���������� �� ������
                // ����� ��������� ���� ������ ��������� �� ������ ��� �������� ������ ���������.
                if (flagInsertCopyDoc == false)
                {
                    form.���������������� = false;
                }

                // �� ���������, ��������� ����� �������� ������ �� ������������.
                if (form.FlagRecordRepeet == false)
                {
                    string queryInsert = string.Empty;

                    // ������ ��� �������� ��� ������ �����.
                    string md5 = string.Empty;

                    if (form.���������������� == true)
                    {
                        if (form.FlagAddDoc == true)
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          " declare @������� int  " +
                                          "select top 1 @������� = ������� from �������� " +
                                          "where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                //"select top 1 @������� = [�������] from �������� " +
                                //"where ���������� >= '" + seletYear.ToString().Trim() + "0101' " +  and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                                 "order by ������� desc " +
                                                 "INSERT INTO �������� " +
                                                 "([id_���������] " +
                                                 ",[id_��������������] " +
                                                 ",[�����] " +
                                                ",[����������] " +
                                                ",[����������] " +
                                                ",[�����������������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[�������������������] " +
                                                ",[��������������] " +
                                                ",[�������] " +
                                                ",[����������������������] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                ",MD5 " +
                                                ",id�����������������������  " +
                                                ", FlagAuto )" +
                                                "VALUES " +
                                                "( " + row["id_���������"] + " " +
                                                "," + row["id_��������������"] + " " +
                                                ",'" + row["�����"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["�����������������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                //",'" + row["���������"] + "' " +
                                                ",'"+ docNumNext.������� +"'" +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["�������������������"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["�������"] + " " +
                                                "," + docNumNext.����� + " " +
                                                //", @������� + 1 " +
                                                ",'" + row["����������������������"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",'" + namFileServer + "'  " +
                                                ",'" + form.PathFileServer + "' " +
                                                ",'md5' " +
                                                "," + ��������������������������.Id + "  " +
                                                ",'True' ) " +
                                                "SELECT @id_�������� = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                        else
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          " declare @������� int  " +
                                              "select top 1 @������� = ������� from �������� " +
                                          "where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                //"select top 1 @������� = [�������] from �������� " +
                                //"where ���������� >= '" + seletYear.ToString().Trim() + "0101' " +  and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                                 "order by ������� desc " +
                                                 "INSERT INTO �������� " +
                                                 "([id_���������] " +
                                                 ",[id_��������������] " +
                                                 ",[�����] " +
                                                ",[����������] " +
                                                ",[����������] " +
                                                ",[�����������������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[�������������������] " +
                                                ",[��������������] " +
                                                ",[�������] " +
                                                ",[����������������������] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                 ",MD5 " +
                                                ",id�����������������������  " +
                                                ", FlagAuto )" +
                                                "VALUES " +
                                                "( " + row["id_���������"] + " " +
                                                "," + row["id_��������������"] + " " +
                                                ",'" + row["�����"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["�����������������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                //",'" + row["���������"] + "' " +
                                                ",'" + docNumNext.������� + "'" +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["�������������������"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["�������"] + " " +
                                                "," + docNumNext.����� + " " +
                                                //", @������� + 1 " +
                                                ",'" + row["����������������������"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",'" + namFileServer + "'  " +
                                                ",'" + form.PathFileServer + "' " +
                                                ",NULL " +
                                                 "," + ��������������������������.Id + "  " +
                                                ",'True' ) " +
                                                "SELECT @id_�������� = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                    }
                    else
                    {

                        if (form.FlagAddDoc == true)
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                        " declare @������� int  " +
                                          " select top 1 @������� = ������� from �������� " +
                                          " where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          " id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                                          " where FlagAuto is null) " +
                                                 "order by ������� desc " +
                                                "INSERT INTO �������� " +
                                                 "([id_���������] " +
                                                 ",[id_��������������] " +
                                                 ",[�����] " +
                                                ",[����������] " +
                                                ",[����������] " +
                                                ",[�����������������] " +
                                                ",[����������] " +
                                //",'" + row["���������"] + "' " +
                                                 ",[���������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[�������������������] " +
                                                ",[��������������] " +
                                                ",[�������] " +
                                                ",[����������������������] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                                ",MD5 " +
                                                  ",id�����������������������  " +
                                                ", FlagAuto )" +
                                                "VALUES " +
                                                "( " + row["id_���������"] + " " +
                                                "," + row["id_��������������"] + " " +
                                                ",'" + row["�����"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["�����������������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                //",'" + row["���������"] + "' " +
                                                ",'" + docNumNext.������� + "'" +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["�������������������"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["�������"] + " " +
                                                "," + docNumNext.����� + " " +
                                                ", @������� + 1 " +
                                                ",'" + row["����������������������"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",NULL  " +
                                                ",'" + guidCard + "' " +
                                                ",'md5' " +
                                                "," + ��������������������������.Id + "  " +
                                                ",'True' ) " +
                                                "SELECT @id_�������� = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                        else
                        {
                            queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                          "begin transaction  " +
                                          "declare @id_�������� int " +
                                          "declare @������� int " +
                                              "select top 1 @������� = ������� from �������� " +
                                          "where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                //"select top 1 @������� = [�������] from �������� " +
                                //"where ���������� >= '" + seletYear.ToString().Trim() + "0101' " +  and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                                 "order by ������� desc " +
                                                "INSERT INTO �������� " +
                                                 "([id_���������] " +
                                                 ",[id_��������������] " +
                                                 ",[�����] " +
                                                ",[����������] " +
                                                ",[����������] " +
                                                ",[�����������������] " +
                                                ",[����������] " +
                                //",'" + row["���������"] + "' " +
                                                  ",[���������] " +
                                                ",[����������] " +
                                                ",[���������] " +
                                                ",[�������������������] " +
                                                ",[��������������] " +
                                                ",[�������] " +
                                                ",[����������������������] " +
                                                ",[FlagPersonData] " +
                                                ",[FlagCardRepeet] " +
                                                ",NameFileDocument  " +
                                                ",GuidName " +
                                               ",MD5 " +
                                                 ",id�����������������������  " +
                                                ", FlagAuto )" +
                                                "VALUES " +
                                                "( " + row["id_���������"] + " " +
                                                "," + row["id_��������������"] + " " +
                                                ",'" + row["�����"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                                ",'" + row["�����������������"] + "' " +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + docNumNext.������� + "'" +
                                                ",'" + row["����������"] + "' " +
                                                ",'" + row["���������"] + "' " +
                                                ",'" + row["�������������������"] + "' " +
                                                ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                                //"," + row["�������"] + " " +
                                                "," + docNumNext.����� + " " +
                                                //", @������� + 1 " +
                                                ",'" + row["����������������������"] + "' " +
                                                ",'" + row["FlagPersonData"] + "' " +
                                                ",'" + form.FlagRecordRepeet + "' " +
                                                ",NULL  " +
                                                ",'" + guidCard + "' " +
                                                ",NULL " +
                                                 "," + ��������������������������.Id + "  " +
                                                ",'True' ) " +
                                                "SELECT @id_�������� = @@IDENTITY  ";

                            builder.Append(queryInsert);
                        }
                    }

                    // ���������� ������ ����������� ����� �������� � ������������� ���� ������� ��������.
                    foreach (string str in s����s)
                    {
                        string insert = "declare @id_" + iCount + "  int " +
                                        "SELECT @id_" + iCount + " = id_���������� " +
                                        "FROM [����������] " +
                                        "where [������������������] = '" + str.Trim() + "' " +
                                        "INSERT INTO [�����������������������������] " +
                                                   "([id����������] " +
                                                   ",[���������������] " +
                                                   ",[����������������] " +
                                                   ",[�����������������] " +
                                                   ",[id��������] " +
                                                   ",[�������������������]) " +
                                             "VALUES " +
                                                   "(@id_" + iCount + " " +
                            //",'" + ����SQL.����(todoy.ToShortDateString()) + "' " +
                                                   ",GETDATE() " +
                                                   ",NULL " +
                                                   ",NULL " +
                                                   ",@id_�������� " +
                                                   ",NULL) ";

                        // ������� � ������.
                        builder.Append(insert);

                        iCount++;
                    }

                    // ���������� ������ � ��������� ������� ���������, ���� ��������� ��������� � ������������ ������� � ���������� ������� ������� ������� ��������.
                    foreach (PersonRecepient person in this.ListPerson)
                    {
                        string insert = "INSERT INTO [������������������������������������] " +
                                        "([id_person] " +
                                       ",[id_�����������������] " +
                                       ",[id_��������]) " +
                                       "VALUES " +
                                       "(" + person.ID + " " +
                                       "," + ��������������������������.Id + " " +
                                       ",@id_�������� ) ";

                        // ������� � ������.
                        builder.Append(insert);
                    }


                    //builder.Append(queryInsert + "COMMIT TRANSACTION ");
                    builder.Append("COMMIT TRANSACTION ");

                    string sTest = builder.ToString().Trim();
                }

                // ��������� ����� ������ ������������.
                if (form.FlagRecordRepeet == true)
                {
                    string queryInsert = string.Empty;
                    if (form.���������������� == true)
                    {
                        queryInsert = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                      "begin transaction  " +
                                   " declare @������� int  " +
                                          " select top 1 @������� = ������� from �������� " +
                                          " where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          " id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                                          " where FlagAuto is null) " +
                                                 "order by ������� desc " +
                                            " INSERT INTO �������� " +
                                             "([id_���������] " +
                                             ",[id_��������������] " +
                                             ",[�����] " +
                                            ",[����������] " +
                                            ",[����������] " +
                                            ",[�����������������] " +
                                            ",[����������] " +
                            //",'" + row["���������"] + "' " +
                                              ",[���������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[�������������������] " +
                                            ",[��������������] " +
                                            ",[�������] " +
                                            ",[����������������������] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                             ",[FlagCardRepeet] " +
                                            ",NameFileDocument  " +
                                              ",id�����������������������  " +
                                                ", FlagAuto )" +
                                            "VALUES " +
                                            "( " + row["id_���������"] + " " +
                                            "," + row["id_��������������"] + " " +
                                            ",'" + row["�����"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["�����������������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                           ",'" + docNumNext.������� + "'" +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["�������������������"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["�������"] + " " +
                                            "," + docNumNext.����� + " " +
                                            //", @������� + 1 " +
                                            ",'" + row["����������������������"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                             ",'" + namFileServer + "'  " +
                                            ",'" + form.PathFileServer + "'  " +
                                             "," + ��������������������������.Id + "  " +
                                                ",'True' ) " +
                                           "INSERT INTO �������������� " +
                                             "([id_���������] " +
                                             ",[id_��������������] " +
                                             ",[�����] " +
                                            ",[����������] " +
                                            ",[����������] " +
                                            ",[�����������������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[�������������������] " +
                                            ",[��������������] " +
                                            ",[�������] " +
                                            ",[����������������������] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                            ",id_����������������  " +
                                            ",�������������� " +
                                            ",FlagControl)" +
                                            "VALUES " +
                                            "( " + row["id_���������"] + " " +
                                            "," + row["id_��������������"] + " " +
                                            ",'" + row["�����"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["�����������������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["�������������������"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["�������"] + " " +
                                            "," + docNumNext.����� + " " +
                                            ",'" + row["����������������������"] + "' " +
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
                                  " declare @������� int  " +
                                          " select top 1 @������� = ������� from �������� " +
                                          " where ���������� >= '" + seletYear.ToString().Trim() + "0101' and ���������� <= '" + seletYear.ToString().Trim() + "1231' " +
                                          " id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                                          " where FlagAuto is null) " +
                                                 "order by ������� desc " +
                                      "INSERT INTO �������� " +
                                             "([id_���������] " +
                                             ",[id_��������������] " +
                                             ",[�����] " +
                                            ",[����������] " +
                                            ",[����������] " +
                                            ",[�����������������] " +
                                            ",[����������] " +
                            //",'" + row["���������"] + "' " +
                                            ",[���������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[�������������������] " +
                                            ",[��������������] " +
                                            ",[�������] " +
                                            ",[����������������������] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                             ",[FlagCardRepeet] " +
                                            ",NameFileDocument  " +
                                               ",id�����������������������  " +
                                                ", FlagAuto )" +
                                            "VALUES " +
                                            "( " + row["id_���������"] + " " +
                                            "," + row["id_��������������"] + " " +
                                            ",'" + row["�����"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["�����������������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                          ",'" + docNumNext.������� + "'" +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["�������������������"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["�������"] + " " +
                                            "," + docNumNext.����� + " " +
                                            //", @������� + 1 " +
                                            ",'" + row["����������������������"] + "' " +
                                            ",'" + row["FlagPersonData"] + "' " +
                                            ",'" + form.FlagRecordRepeet + "' " +
                                             ",NULL  " +
                                             ",NULL  " +
                                               "," + ��������������������������.Id + "  " +
                                                ",'True' ) " +
                                           "INSERT INTO �������������� " +
                                             "([id_���������] " +
                                             ",[id_��������������] " +
                                             ",[�����] " +
                                            ",[����������] " +
                                            ",[����������] " +
                                            ",[�����������������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[����������] " +
                                            ",[���������] " +
                                            ",[�������������������] " +
                                            ",[��������������] " +
                                            ",[�������] " +
                                            ",[����������������������] " +
                                            ",[FlagPersonData] " +
                                            ",[FlagCardRepeet] " +
                                            ",id_����������������  " +
                                            ",�������������� " +
                                            ",FlagControl)" +
                                            "VALUES " +
                                            "( " + row["id_���������"] + " " +
                                            "," + row["id_��������������"] + " " +
                                            ",'" + row["�����"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString()) + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["����������"]).ToShortDateString().Trim()) + "' " +
                                            ",'" + row["�����������������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["����������"] + "' " +
                                            ",'" + row["���������"] + "' " +
                                            ",'" + row["�������������������"] + "' " +
                                            ",'" + ����SQL.����(Convert.ToDateTime(row["��������������"]).ToShortDateString().Trim()) + "' " +
                            //"," + row["�������"] + " " +
                                            "," + docNumNext.����� + " " +
                                            ",'" + row["����������������������"] + "' " +
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

                //// �������� ������.
                ������������ connectBD = new ������������();
                string sCon = connectBD.�����������������();

                // ���� �������� �������� ����� �����.
                bool flagCopyServer = false;

                // �������� ������ �� ������� (� ��������� �� � ������ ����������.
                using (SqlConnection con = new SqlConnection(sCon))
                {

                    //Log.WriteLine(filePatchLog, "����� ��������� ���� �� ������");

                    //if (form.SaveDocServer == true)
                    //{
                    //    try
                    //    {
                    //        if (form.���������������� == true)
                    //        {

                    //            Log.WriteLine(filePatchLog, "�������� ���� �� ������");

                    //            // �������� ������� �� �������� ��� ������ �� ������.
                    //            if (flagInsertCopyDoc == true)
                    //            {
                    //                //��������� ���� �� ������ �������� ����������.
                    //                //File.Copy(fileName, fileNameCopy, true);
                    //            }
                    //            Log.WriteLine(filePatchLog, "�������� ���������� ���� �� ������");

                    //            //���� ���� ������������ ������� ������� ���� � true.
                    //            flagCopyServer = true;
                    //        }
                    //    }
                    //    catch(Exception exp)
                    //    {
                    //        Log.WriteLine(filePatchLog, "������ ��� ����������� - ");
                    //        Log.WriteLine(filePatchLog, exp.Message);
                    //        MessageBox.Show("������ ��� ����������� �����");

                    //        flagCopyServer = false;

                    //        return;
                    //    }

                    //    string fileTest = fileNameCopy;
                    //    if (File.Exists(fileNameCopy) == true)
                    //    {
                    //        Log.WriteLine(filePatchLog, "�������� ������ �� ������");

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
                    // ���� ���� ������������ ������� ������� ���� � true.
                    flagCopyServer = true;

                    con.Open();
                    SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    com.ExecuteNonQuery();
                    //}
                }


                //ds11.��������.Add��������Row(row);
                ��������������();

                // ����� �� ������ �������� � ������� ����� ���������.
                string queryNumDoc = "select id_��������,�������,��������� from [��������] " +
                                     "where GuidName = '" + guidCard + "' ";

                string �������� = string.Empty;

                DataTable tabNum;

                using (SqlConnection con = new SqlConnection(sCon))
                {
                    con.Open();

                    SqlDataAdapter da = new SqlDataAdapter(queryNumDoc, con);

                    DataSet ds = new DataSet();

                    da.Fill(ds, "numDoc");

                    tabNum = ds.Tables["numDoc"];
                }

                �������� = tabNum.Rows[0]["�������"].ToString().Trim() + "/" + tabNum.Rows[0]["���������"].ToString().Trim();

                string ����� = ��������;

                // ������� ����� id ��������.
                string idCard = tabNum.Rows[0]["id_��������"].ToString().Trim();

                // ������� ����� ������������������� ���������.
                FormMessage frmMessage = new FormMessage(�����);
                frmMessage.NumCardDoc = idCard.Trim();
                frmMessage.�������������� = ��������;
                frmMessage.�������������������������� = ��������������������������;
                frmMessage.TopMost = true;
                frmMessage.ShowDialog();

            }
        }


        /// <summary>
        /// ��������� �������� ���������.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void menuItem19_Click(object sender, EventArgs e)
        {
            string iTest = ������������;

            int seletYear = Convert.ToInt16(this.������������) + 1;

            Form����������������� form = new Form�����������������(ds11, ������������, true);
            //Form����������������� form = new Form�����������������(ds11, seletYear.ToString(), true);

            // ��������� ���� � false.
            form.Flag����������� = false;

            // ��������� �������.
            form.������� = "";

            DialogResult result = form.ShowDialog(this);
            if (result == DialogResult.OK)
            {

                // ������� ��������� ������ ����������� ��������� ������ ������������.
                Item�������������������������� �������������������������� = form.�����������������;

                DS1.�����������������Row row = form.�����������������������;

                �������������� doc = new ��������������();

                if (form.FlagNumStopDoc == false)
                {
                    doc.����� = Convert.ToInt16(form.�����Doc.�����);
                }
                else
                {
                    // ���� �� ��������� �������������� ������������� ������.
                    doc.����� = form.NumDocNoAutomat;
                }


                // ������� ����������.
                numberPrefix = string.Empty;

                // ������� ������� ������ ���������.
                numberPrefix = form.���������������������;

                //ds11.�����������������.Add�����������������Row(row); 

                List<int> listId������� = form.ListID��������;
                //List<�����������������> listOP = form.List�����������������;

                // ���� �� ����� ������������ ������.
                if (form.Flag����������� == false)
                {
                    // ������ ��� �������� SQL ����������, ��� ���������� � ����� ����������.
                    StringBuilder buildInsert = new StringBuilder();

                    // ���������� ��� �������� ������
                    string numDirect = string.Empty;

                    if (form.FlagNumStopDoc == false)
                    {
                        string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                       "begin transaction  " +
                                       "declare @numDoc int " +
                                       "select top 1 @numDoc = ��������������� from ����������������� " +
                                       "where ���� >= '" + seletYear.ToString().Trim() + "0101' and ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                                       "order by id_�������� desc " +
                                       "declare @key int " +
                                       "INSERT INTO ����������������� " +
                                       "([����] " +
                                       ",[�������������] " +
                                       ",[id_�������������] " +
                                       ",[�������������������] " +
                                       ",[���������������] " +
                                       ",[id_��������] " +
                                       ",[����������] " +
                                       ",[id_������������������] " +
                                       ",[����������������������] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ",id����������������������� ) " +
                                       "VALUES " +
                                       "('" + ����SQL.����(Convert.ToDateTime(row["����"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row["�������������"] + "' " +
                                       "," + row["id_�������������"] + " " +
                                       ",'" + row["�������������������"] + "' " +
                            //"," + row["���������������"] + " " +
                            //", "+ doc.����� + " " +
                                       ", @numDoc + 1  " +
                                       "," + row["id_��������"] + " " +
                                       ",'" + row["����������"] + "' " +
                            //","+ row["id_������������������"]+" " +
                                       ",NULL " +
                            //",'"+ form.�������.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       // ",NULL " +
                                       //",NULL " +
                                       ", " + ��������������������������.Id + " ) " +
                                       "set @key = @@IDENTITY ";

                        buildInsert.Append(query);
                    }
                    else
                    {
                        string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                       "begin transaction  " +
                                       "declare @numDoc int " +
                                       "select top 1 @numDoc = ��������������� from ����������������� " +
                                       "where ���� >= '" + seletYear.ToString().Trim() + "0101' and ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                                       "order by id_�������� desc " +
                                       "declare @key int " +
                                       "INSERT INTO ����������������� " +
                                       "([����] " +
                                       ",[�������������] " +
                                       ",[id_�������������] " +
                                       ",[�������������������] " +
                                       ",[���������������] " +
                                       ",[id_��������] " +
                                       ",[����������] " +
                                       ",[id_������������������] " +
                                       ",[����������������������] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ",id����������������������� " +
                                       ", FlagAutho ) " + 
                                       "VALUES " +
                                       "('" + ����SQL.����(Convert.ToDateTime(row["����"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row["�������������"] + "' " +
                                       "," + row["id_�������������"] + " " +
                                       ",'" + row["�������������������"] + "' " +
                            //"," + row["���������������"] + " " +
                                       ", "+ doc.����� + " " +
                                       //", @numDoc + 1  " +
                                       "," + row["id_��������"] + " " +
                                       ",'" + row["����������"] + "' " +
                            //","+ row["id_������������������"]+" " +
                                       ",NULL " +
                            //",'"+ form.�������.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       // ",NULL " +
                                       //",NULL " +
                                       ", " + ��������������������������.Id + " " +
                                       ", 'True' ) " + 
                                       "set @key = @@IDENTITY ";

                        buildInsert.Append(query);
                    }

                    // ��������� � ������ ������� �� ������� SQL ���������� �� ������� � ������� [���������������������������������������].

                    string sInsert = string.Empty;
                    sInsert = String.Format(form.QueryInsert.Trim(), "@key");

                    buildInsert.Append(sInsert.Trim());

                    // �������� ����������.
                    buildInsert.Append("COMMIT TRANSACTION ");

                    // ������� ������ ��� �������� ��������� �������� ����� ��������������.
                    //form.List�����������������.Clear();

                    ������������ connBD = new ������������();
                    string sCon = connBD.�����������������();

                    SqlConnection con = new SqlConnection(sCon);
                    con.Open();
                    SqlCommand com = new SqlCommand(buildInsert.ToString(), con);
                    com.ExecuteNonQuery();
                    con.Close();
                }
                else
                {
                    // ���������.
                    DS1.�����������������Row row2 = form.�����������������������;

                    int i = row2.id_������������������;

                    // ������ ��� �������� �������.
                    System.Text.StringBuilder builder = new System.Text.StringBuilder();

                    if (form.FlagNumStopDoc == false)
                    {
                        // ��������� ���� � FALSE.
                        string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                        "begin transaction  " +
                                        "declare @numCard  int " +
                                        "select top 1 @numCard = ���������������  from ����������������� " +
                                        "where ���� >= '" + seletYear.ToString().Trim() + "0101' and ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                                        "order by id_�������� desc " +
                                       "INSERT INTO ����������������� " +
                                       "([����] " +
                                       ",[�������������] " +
                                       ",[id_�������������] " +
                                       ",[�������������������] " +
                                       ",[���������������] " +
                                       ",[id_��������] " +
                                       ",[����������] " +
                                       ",[id_������������������] " +
                                       ",[����������������������] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ", id�����������������������)" +
                                       "VALUES " +
                                       "('" + ����SQL.����(Convert.ToDateTime(row2["����"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row2["�������������"] + "' " +
                                       "," + row2["id_�������������"] + " " +
                                       ",'" + row2["�������������������"] + "' " +
                            //"," + row2["���������������"] + " " +
                            //", " + doc.����� + " " +
                                        ", @numCard + 1  " +
                                       "," + row2["id_��������"] + " " +
                                       ",'" + row2["����������"] + "' " +
                                       "," + row2["id_������������������"] + " " +
                            //",NULL " +
                            //",'"+ form.�������.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row2["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       //",NULL " +
                                       //",NULL " +
                                       ", " + ��������������������������.Id + " )" +
                                       " declare @idCard int " +
                                       "select top 1 @idCard = id_��������  from ����������������� " +
                                       "order by id_�������� desc ";

                        builder.Append(query);
                    }
                    else
                    {
                        // ��������� ���� � FALSE.
                        string query = "SET TRANSACTION ISOLATION LEVEL serializable " +
                                        "begin transaction  " +
                                        "declare @numCard  int " +
                                        "select top 1 @numCard = ���������������  from ����������������� " +
                                        "where ���� >= '" + seletYear.ToString().Trim() + "0101' and ���� <= '" + seletYear.ToString().Trim() + "1231' " +
                                        "order by id_�������� desc " +
                                       "INSERT INTO ����������������� " +
                                       "([����] " +
                                       ",[�������������] " +
                                       ",[id_�������������] " +
                                       ",[�������������������] " +
                                       ",[���������������] " +
                                       ",[id_��������] " +
                                       ",[����������] " +
                                       ",[id_������������������] " +
                                       ",[����������������������] " +
                                       ",[FlagPersonData] " +
                                       ",[GUID] " +
                                       //",FileData " +
                                       //",FileDateTitlePage " +
                                       ", id�����������������������" +
                                       ",FlagAutho )" +
                                       "VALUES " +
                                       "('" + ����SQL.����(Convert.ToDateTime(row2["����"]).ToShortDateString().Trim()) + "' " +
                                       ",'" + row2["�������������"] + "' " +
                                       "," + row2["id_�������������"] + " " +
                                       ",'" + row2["�������������������"] + "' " +
                            //"," + row2["���������������"] + " " +
                                        ", " + doc.����� + " " +
                                        //", @numCard + 1  " +
                                       "," + row2["id_��������"] + " " +
                                       ",'" + row2["����������"] + "' " +
                                       "," + row2["id_������������������"] + " " +
                            //",NULL " +
                            //",'"+ form.�������.Trim() +"' " +
                                       ",NULL" +
                                       ",'" + row2["FlagPersonData"] + "' " +
                                       ",'" + form.StrGuid.Trim() + "'  " +
                                       //",NULL " +
                                       //",NULL " +
                                       ", " + ��������������������������.Id + " " +
                                       ", 'True' ) " +
                                       " declare @idCard int " +
                                       "select top 1 @idCard = id_��������  from ����������������� " +
                                       "order by id_�������� desc ";
                        builder.Append(query);
                    }

                    string ������������������ = string.Empty;

                    DataRow[] rowsSelect = ds11.���������������������.Select("id_�������������= " + Convert.ToInt32(row2["id_�������������"]) + " ");
                    foreach (DataRow item in rowsSelect)
                    {
                        ������������������ = item["������������������"].ToString().Trim();
                    }

                    string ������������������� = "��� �����. � ���. ��������� " + row2["�������������"].ToString().Trim() + "-" + row2["�������������������"].ToString().Trim() + "-" + ������������������ + "/" + doc.�����.ToString().Trim();// row2["���������������"].ToString().Trim();
                    //string ������������������� = "��� �����. � ���. ��������� " + row2["�������������"].ToString().Trim() + "-" + row2["�������������������"].ToString().Trim() + "-" + ������������������ + "/CAST(@numCard + 1 AS nvarchar) ";// +row2["���������������"].ToString().Trim();

                    // ��������� ���� � TRUE.
                    string queryUpdate = "UPDATE [��������] " +
                                         "SET ������������������� = '" + ������������������� + "' " + //' + CAST(@numCard + 1 AS nvarchar) " +
                        //"FlagPersonData = '" + row["FlagPersonData"] + "' " +
                                         ",����� = 'True' " +
                                         "where id_�������� = " + row["id_������������������"] + " ";
                    // ������ ������ ��������� �� ���������� ������ � �� �������������� � ������ ������, ����� ��������� �� � ����� ����������.
                    builder.Append(queryUpdate);

                    string sTestNum = builder.ToString().Trim();

                    // ������ �� ������� id � ��������� ������� �����������������������.
                    foreach (����������������� itm in form.List�����������������)
                    {
                        string queryIns = "INSERT INTO [���������������������������������������] " +
                                       "([id_��������] " +
                                       ",[id_�����������������]) " +
                                       "VALUES " +
                            //"('" + row.id_�������� + "' " +
                                       "( @idCard " +
                                       ",'" + itm.Id_����������������� + "' ) ";

                        builder.Append(queryIns);
                    }

                    // �������� ��������� ������� �������� �����������������.
                    foreach (int id�� in listId�������)
                    {

                        string queryId�� = "INSERT INTO [����������������������������������] " +
                                           "([id_����������������] " +
                                           ",[id_�����������������]) " +
                                           "VALUES " +
                                           "(" + id�� + " " +
                            //"," + row.id_�������� + " ) " +
                                            ",@idCard ) " +
                                           "update �������� " +
                                           "set ������������������� = '" + form.��������������.Trim() + "' " + " + CAST(@numCard + 1 AS nvarchar) " +
                                           "where id_�������� = " + id�� + " ";

                        builder.Append(queryId��);
                    }

                    // ��������, ��� �������� �� ������� �� �������� ����� � ������� ��������� �������.
                    �������������� card = new ��������������(Convert.ToInt32(row["id_������������������"]));
                    bool flagStatusRepeet = card.������������������������();

                    // ���� ������ = true ������ �� ����� ���� � ���������� �� ������� ������������ ���������� ������ �����.
                    if (flagStatusRepeet == true)
                    {
                        // ������ ��� ���������� ������� � ����� ����������.
                        //StringBuilder querTransact = new StringBuilder();

                        /*
                         * �������� �������� �� �� ���� �������� ������� ��� ���.
                         * ��� ����� ������ �������� � ���� ����� � ������� ��������, ���� ����������� �������� False 
                         * ����� �� �������� �� �������� ������� � ��������� ������ ���.
                        */
                        ������������ bdConnect = new ������������();
                        using (SqlConnection conn = new SqlConnection(bdConnect.�����������������().Trim()))
                        {
                            conn.Open();
                            �������������� card2 = new ��������������(Convert.ToInt32(row["id_������������������"]));
                            bool flagVD = card2.Get��������������(conn);

                            // ���� �� �������� �������� �������� �������.
                            if (flagVD == false)
                            {
    
                                string queryUp = " update �������� " +
                                              "set ����� = 'True' " +
                                              "where id_�������� = " + Convert.ToInt32(row["id_������������������"]) + " " +
                                              " declare @date datetime " +
                                              "declare @day int " +
                                              "declare @SetDate datetime " +
                                              "select @date = ��������������,@day = �������������� from �������������� " +
                                              "where id_���������������� = " + Convert.ToInt32(row["id_������������������"]) + " " +
                                              "SELECT @SetDate = DATEADD(day, @day, @date); " +
                                              "update �������������� " +
                                              "set FlagControl = 'True' " +
                                              ",�������������� = @SetDate " +
                                              "where id_���������������� = " + Convert.ToInt32(row["id_������������������"]) + " ";

                                builder.Append(queryUp);
                            }

                            // ���� ����� ���������.
                            if (flagVD == true)
                            {
                                // �������� �������� � ���� �������������� � ������� �������������� �� ���������� ���� ��������� � ���� ��������������.
                                string queryUpdatDate = " declare @date datetime " +
                                                        "declare @day int " +
                                                        "declare @SetDate datetime " +
                                                        "select @date = ��������������,@day = �������������� from �������������� " +
                                                        "where id_���������������� = " + Convert.ToInt32(row["id_������������������"]) + " " +
                                                        "SELECT @SetDate = DATEADD(day, @day, @date); " +
                                                        "update �������������� " +
                                                        "set �������������� = @SetDate " +
                                                        "where id_���������������� = " + Convert.ToInt32(row["id_������������������"]) + " ";

                                //querTransact.Append(queryUpdatDate);
                                builder.Append(queryUpdatDate);
                            }
                        }
                    }

                    // �������� ����������.
                    builder.Append("COMMIT TRANSACTION ");

                    string queryTest = builder.ToString().Trim();

                    // �������� ������.
                    ������������ strConnectBD = new ������������();
                    string strConn = strConnectBD.�����������������();

                    // ������� ���������� � �������� ������.
                    SqlConnection con = new SqlConnection(strConn);
                    con.Open();
                    SqlCommand com = new SqlCommand(builder.ToString().Trim(), con);
                    com.ExecuteNonQuery();

                    // ������� ����������.
                    con.Close();

                }

                ��������������();

                string iTest2 = "test";

                // ������� ����� ���������.
                NumOutputCardVipNet numDoc = GetNumDocOutVipNet(form.StrGuid);

                //string numberDocument = form.���������������������.Trim() + "/" + numDoc.Trim();

                string numberDocument = numberPrefix + "/" + numDoc.���������������.Trim();

                // ������� ��������� � ����� �������.
                FormMessage message = new FormMessage(numberDocument.Trim());
                message.TopMost = true;
                message.�������������������������� = ��������������������������;
                message.NumCardDoc = numDoc.Id.ToString().Trim();
                message.�������������� = numberDocument;
                message.ShowDialog();

            }
        }

        private void menuItem24_Click(object sender, EventArgs e)
        {

            ����������������������������();
            //������������� = new System.Threading.Thread(new System.Threading.ThreadStart(����������������������));
            //�������������.Start();

            //����������������������������();
        }

        private void menuItem25_Click(object sender, EventArgs e)
        {
            //������������� = new System.Threading.Thread(new System.Threading.ThreadStart(����������������������));
            //�������������.Start();

            //����������������������������();

            //string querySelect = "SELECT * FROM [�������] " +
            //                    "where ��������������<'" + ����SQL.����(DateTime.Today.ToShortDateString()) + "' AND ���������� >= '" + ������������ + "0112' AND ����������='True' AND �����='False'";

            //string query = "SELECT     convert(VARCHAR,dbo.��������.�������) + '/' + dbo.��������.��������� as '���������',  dbo.��������.����������, dbo.��������������.����������������������, " +
            //               " dbo.��������.�����������������, dbo.��������.����������, dbo.��������.����������, dbo.��������.��������������, " +
            //               " dbo.��������.��������� " +
            //               "FROM         dbo.�������� INNER JOIN " +
            //               "  dbo.�������������� ON dbo.��������.id_�������������� = dbo.��������������.id_�������������� " +
            //               " where ����� = 'False' and �������������� < CONVERT(DATE,GETDATE()) ";

            string query = "select ���������,����������,����������������������,�����������������, ����������, ����������, ��������������,������������������ as ��������� from dbo.View������������ " +
                           "where ����� = 'False' and �������������� < CONVERT(DATE,GETDATE()) ";

            GetDataTable getTable = new GetDataTable(query);
            DataTable tab = getTable.DataTable("�������");

            Form��������������� form = new Form���������������();
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
                RangeDate rd = formSelD.����������;

                FormStatInputKorr form = new FormStatInputKorr();
                form.TopMost = true;
                form.����������� = rd;
                form.Show();
            }

        }

        private void menuItem27_Click(object sender, EventArgs e)
        {
            FormSelectDate formSelD = new FormSelectDate();
            formSelD.ShowDialog();

            if (formSelD.DialogResult == DialogResult.OK)
            {
                RangeDate rd = formSelD.����������;

                FormViewInputDoc form = new FormViewInputDoc();
                //form.TopMost = true;
                form.����������� = rd;
                form.Show();
                
            }
        }

        private void menuItem28_Click(object sender, EventArgs e)
        {
            Form�������������� form = new Form��������������();
            form.���������� = ����.����������(selectedYear.ToString());
            form.���������� = �������������;
            form.ShowDialog(this);
            this.Enabled = true;
        }

        private void menuItem29_Click(object sender, EventArgs e)
        {
            FormSelectDate formSelD = new FormSelectDate();
            formSelD.ShowDialog();

            if (formSelD.DialogResult == DialogResult.OK)
            {
                RangeDate rd = formSelD.����������;

                FormStatOutpotCorr form = new FormStatOutpotCorr();
                form.TopMost = true;
                form.����������� = rd;
                form.Show();
            }
        }

        private void menuItem30_Click(object sender, EventArgs e)
        {
            FormSelectDate formSelD = new FormSelectDate();
            formSelD.ShowDialog();

            if (formSelD.DialogResult == DialogResult.OK)
            {
                RangeDate rd = formSelD.����������;

                FormViewOutputDoc form = new FormViewOutputDoc();
                form.TopMost = true;
                form.����������� = rd;
                form.Show();

            }
        }

        private void menuItem31_Click(object sender, EventArgs e)
        {
            Form������������������������� form = new Form�������������������������(�������������);
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
            string query = "SELECT     convert(VARCHAR,dbo.��������.�������) + '/' + dbo.��������.��������� as '���������',  dbo.��������.����������, dbo.��������������.����������������������, " +
                         " dbo.��������.�����������������, dbo.��������.����������, dbo.��������.����������, dbo.��������.��������������, " +
                         " dbo.��������.��������� " +
                         "FROM         dbo.�������� INNER JOIN " +
                         "  dbo.�������������� ON dbo.��������.id_�������������� = dbo.��������������.id_�������������� " +
                         " where ����� = 'False' and �������������� < CONVERT(DATE,GETDATE()) ";

            GetDataTable getTable = new GetDataTable(query);
            DataTable tab = getTable.DataTable("�������");

            Form��������������� form = new Form���������������();
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
            // �������� �����.
            FormSelectInputDatePerson forminput = new FormSelectInputDatePerson();
            forminput.Show();
        }

        

       

        

    }
}

