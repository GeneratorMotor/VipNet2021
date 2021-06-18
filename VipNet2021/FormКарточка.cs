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
    /// Summary description for Form��������.
    /// </summary>
    public class Form�������� : System.Windows.Forms.Form
    {
        public RegKor.DS1.��������Row ��������������;
        private System.Windows.Forms.Label label�������ID;
        private System.Windows.Forms.Label label�������������������;
        private System.Windows.Forms.Label label���������;
        private System.Windows.Forms.Label label����������;
        private System.Windows.Forms.Label label��������;
        private System.Windows.Forms.Label label�������������;
        public System.Windows.Forms.CheckBox checkBox�����;
        private System.Windows.Forms.Button button������;
        private System.Windows.Forms.Button button���������;
        private System.Windows.Forms.GroupBox groupBox����������;
        private System.Windows.Forms.Label label��������������;
        private System.Windows.Forms.Label label���������������;
        public System.Windows.Forms.DateTimePicker dateTime���������������;
        public System.Windows.Forms.TextBox textBox��������������;
        public System.Windows.Forms.TextBox textBox���������;
        public System.Windows.Forms.TextBox textBox����������;
        public System.Windows.Forms.DateTimePicker dateTime��������;
        public System.Windows.Forms.CheckBox checkBox��������;
        public System.Windows.Forms.ComboBox combo��������;
        public System.Windows.Forms.ComboBox combo�������������;
        private System.Windows.Forms.GroupBox groupBox���������;
        private System.Windows.Forms.Label label�������������;
        private System.Windows.Forms.Label label���������������;
        public System.Windows.Forms.DateTimePicker dateTime���������������;
        private RegKor.DS1 ds1;
        public System.Windows.Forms.TextBox textBox�������������������;
        private System.Windows.Forms.Panel panelKontrol;
        private string ��������������������� = "";
        private System.Windows.Forms.Button button�������������������;
        private System.Windows.Forms.ToolTip toolTip1;
        private Panel panel1;
        private System.ComponentModel.IContainer components;
        private Label label1;

        // ������� ���.
        private int currentYear;

        /// <summary>
        /// ����� ����������, ������� ������ ���� �������� � ����
        /// </summary>
        private int �����������;

        /// <summary>
        /// ����� ���������� ��� ����������
        /// </summary>
        private int ������� = 0;
        private Button btnElementPS;

        /// <summary>
        /// ��� ����� �������� ��� ���������� ������������
        /// </summary>
        private bool ������������� = false;

        private int id��������Properti;

        /// <summary>
        /// ������ ��������� ���.
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
        /// �������� ������ id ��������.
        /// </summary>
        public int Id��������
        {
            get
            {
                return id��������Properti;
            }
            set
            {
                id��������Properti = value;
            }
        }
        private string queryStringProperty = string.Empty;

        /// <summary>
        /// �������� ������ ������ ������� ��� ���������� ��������� ������� ����������������������.
        /// </summary>
        public string QueryString��������������
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

        int ���� = 0;

        /// <summary>
        /// ������ id ������� ���� ��������� ������������ ������.
        /// </summary>
        public int Id�������������������������������
        {
            get
            {
                return ����;
            }
            set
            {
                ���� = value;
            }
        }

        bool _�������;

        /// <summary>
        ///  ������ ������� � �������� ��� ������ � �������� ������������ ������.
        /// </summary>
        public bool ��������������������
        {
            get
            {
                return _�������;
            }
            set
            {
                _������� = value;
            }
        }

        private �������������������� �������������;
        private CheckBox chBoxRepet;

        /// <summary>
        /// ������ �������� ��� ����� � ��������� ������� ������.
        /// </summary>
        public �������������������� �������������
        {
            get
            {
                return �������������;
            }
            set
            {
                ������������� = value;
            }
        }

        private DatePersonal dPerson = new DatePersonal();
        private Label label2;
        private TextBox txtPeriod;

        /// <summary>
        /// ������ ������������������ ������������ ������.
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
        /// ������ ��������� ����������� ��� ������ ����� ����� ������������ �����.
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
        /// ������ ���������� ����.
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
        /// ���� ��������� ��� ��������� ��������� ����������� �� �������.
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

        //���������� ��� �������� ���� � ����� � �����������.
        private string pathFileServer = string.Empty;

        // ���������� ��� �������� ���� � ���������� �����.
        private string pathFileServerTitlePage = string.Empty;

        /// <summary>
        /// ������ ���� � ����� � ����������� �� ��������� ������.
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

        // ���������� ��� �������� ����� ����� ������� ���������� �� ������.
        private string fileName = string.Empty;
        private string fileNameCopy = string.Empty;

        /// <summary>
        /// �������� ��� ����� ������� ����� ������������.
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


        // ���������� ��� �������� ���������� ������ ���������.
        private int �����������������������;
        private LinkLabel linkLabel1;

        // ���������� ������� ������ ��������� ���������� ����� ���������.
        private string lastNumberDoc = string.Empty;

        // ���� ��������� ��� ������������ ����������� ���������� ��������� ���������.
        private bool flagLastNumberDoc = false;

        // ��������� ������ ����������� ��������� ����� ���������.
        �������������� numDoc = new ��������������();

        /// <summary>
        /// ������ ����� ����� ������������������� ���������.
        /// </summary>
        public �������������� �����������������������
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

        private string ����������� = string.Empty;

        /// <summary>
        /// ������ ����� ��������� ������������ � ��.
        /// </summary>
        public string ������������
        {
            get
            {
                return �����������;
            }
            set
            {
                ����������� = value;
            }
        }

        private string ������������������ = string.Empty;

        /// <summary>
        /// ������ ��������� ����.
        /// </summary>
        public string ������������������
        {
            get
            {
                return ������������������;
            }
            set
            {
                ������������������ = value;
            }
        }

        // ���������� ��� �������� ���� � ������� ��� ����������� ����� �� ������.
        private string patchServerSave = string.Empty;

        // ���������� ��� �������� ����� ����� �� �������.
        private string fileNameServer = string.Empty;

        private bool ����������������;
        private MaskedTextBox textBox�������������;
        private CheckBox chcDop;

        /// <summary>
        /// �������� ���������, ��� ����� � ���������� ����� ���������� �� ������.
        /// </summary>
        public bool ����������������
        {
            get
            {
                return ����������������;
            }
            set
            {
                ���������������� = value;
            }
        }


        private bool _flagUpdateRecord;

        /// <summary>
        /// �������� ���������� ������.
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
        /// �������� ���������� ��� � �������� � ������� �������� ���������� ������.
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

        // ���������� ��� �������� ���������� ������� ��������� ���������.
        private Item�������������������������� item;
        private LinkLabel linkLabel2;

        /// <summary>
        /// ���������� ��������� ����� ����������� ���������.
        /// </summary>
        public Item�������������������������� �����������������
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
        /// �������� ������ ������ ���������� ������� ������� ��� ��������� ����������� ���������.
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

        private TextBox textBox�������������2;

        // ���������� ��������� ��� ������� ������� ���������� ����� ����������.
        private bool flagAutoNumberDocStoip = false;
        private CheckBox chboxDsp;

        // ���������� ��� �������� ������ ��������� ���������.
        private string sNumStart = string.Empty;

        private string flagDsp = "False";

        /// <summary>
        /// ������ ��������� �������� �������� ��� ��� ���.
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

        // ��������� ��� �������� id ��������� �������� ������������ ������.
        private int idPersonDate = 0;

        // �������� ��� �������� id ��������� �������� ������������ ������.
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

        // ���������� ��� �������� ������ ������� � ������������ ������ ��� �������� ��������.
        private string queryPersonDateForCardInput = string.Empty;

        /// <summary>
        /// ���������� ��� �������� ������ ������� � ������������ ������ ��� �������� ��������.
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
        /// ���������� ��� �������� ������� ����� ��������.
        /// </summary>
        private string _���������������������������� = string.Empty;

        /// <summary>
        /// ����������� �����. � �������� ��������� ��������� �������. ������������ ��� �������� ����� ������
        /// </summary>
        /// <param name="ds">�������</param>
        public Form��������(RegKor.DS1 ds, string ������������, bool flagAutoStop)
        {
            InitializeComponent();

            // �������� ��� �������� �������������� ���������� �������� ������� ����������.
            flagAutoNumberDocStoip = flagAutoStop;

            this.ds1 = ds;
            �������������� = ds1.��������.New��������Row();

            // ������� id ��������.
            int id_�������� = ��������������.id_��������;

            // ��������� id �������� � �������� �����.
            this.Id�������� = id_��������;

            ������������� = true;

            string ���������������� = string.Empty;

            // �������� ������ ������� ������������� ��������.
            LoadNumberDepartments();

            
            // ����� � �� ������ �������� ���������� ����� ������� � ������� �������� � ������ 1 ����� ����� ������ ����.

            // ��������� ������������� �������� �� ������� �����������������
            //DataRow[] dr = ds.��������.Select("���������� >='01.12." + ������������ + "'", "������� DESC");

            //string query = "declare @numDoc int  " +
            //               " select top 1 @numDoc = ������� from �������� " +
            //                " where ���������� >= '" + ������������ + "0101' and ���������� <= '" + ������������ + "1231' and " +
            //                " id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
            //                " where FlagAuto is null) " +
            //                " order by id_�������� desc ";

            string query = " select top 1 ������� from �������� " +
                " where ���������� <= '" + ������������ + "1231' and ���������� >= '" + (Convert.ToInt32(������������) - 1).ToString().Trim() + "1231' and FlagAuto is null " +
                  "order by id_�������� desc ";
                //" where ���������� <= '" + ������������ + "1231' and  FlagAuto is null " +
                          // "where ���������� <= '" + ������������ + "1231' and FlagAuto is null " +
                //" id_�������� in (SELECT MAX(id_��������) FROM [��������] " +
                //" where FlagAuto is null) " +
                //" order by id_�������� asc ";
                         

            DataRow[] dr = DataTableSql.GetDataTableRows(query);

            if (dr.Length > 0)
            {
                if (flagAutoStop == false)
                {
                    ����������� = 1 +(int)dr[0]["�������"];
                    label1.Text = "����. ����� �\\� " + (�����������);
                    //textBox�������������.Text = �����������.ToString() + "/12-02-0";

                    if (this.rb02.Checked == true)
                    {
                        //textBox�������������.Text = "12-02-";
                        textBox�������������.Text = "02-";
                    }
                    else if (this.rb04.Checked == true)
                    {
                        textBox�������������.Text = "04-";
                    }

                    // ��������� ��������� ����� ���������.
                    ����������������������� = �����������;

                    // ������ ���� TextBox �� MasckTextBox.
                    //���������������� = "12-02-" + textBox�������������.Text.Trim();//0";
                    ���������������� = textBox�������������.Text.Trim();//0";

                    // ������ ���� ����� ������.
                    this.textBox�������������2.Visible = false;
                }
                else
                {
                    // ����� ����� ����� ������.
                    this.textBox�������������.Visible = false;

                    ����������� = 1 + (int)dr[0]["�������"];
                    label1.Text = "����. ����� �\\� " + (�����������);
                    textBox�������������2.Text = �����������.ToString() + "/12-02-0";

                    //this.maskedTextBox1.Visible = false;
                }
            }
            else
            {
                ����������� = 1;
                label1.Text = "����. ����� �\\� " + (�����������);

                    //textBox�������������.Text = �����������.ToString() + "/12-02-0";

                    // ������ ���� TextBox �� MasckTextBox.
                    //textBox�������������.Text = "";
                    textBox�������������.Text = "12-02-";

                    // ��������� ��������� ����� ���������.
                    ����������������������� = �����������;

                    // ������ ���� TextBox �� MasckTextBox.
                    //���������������� = "12-02-" + textBox�������������.Text.Trim();
                    ���������������� = textBox�������������.Text.Trim();

                    // Test.
                    this.textBox�������������2.Visible = true;

                    this.textBox�������������.Visible = false;
               

            }

            // ������� ����� ����� ������������ ��������� � �����.
            �������������� doc = new ��������������();
            doc.����� = �����������������������;
            doc.������� = ����������������;

            // �������� ����� ��������� � �������� �����.
            ����������������������� = doc;

            combo�������������.DataSource = ds1.��������������;
            combo�������������.DisplayMember = ds1.��������������.Columns["����������������������"].ToString();
            combo�������������.ValueMember = ds1.��������������.Columns["id_��������������"].ToString();
            combo�������������.Text = "";

            combo��������.DataSource = ds1.���������;
            combo��������.DisplayMember = ds1.���������.Columns["�����������������"].ToString();
            combo��������.ValueMember = ds1.���������.Columns["id_���������"].ToString();

            combo�������������.Focus();


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
        /// ����������� �����. � �������� ���������� ��������� �������, � ������������� ������ ��� ���������
        /// </summary>
        /// <param name="ds">�������</param>
        /// <param name="id��������">������������� ������ ��� ���������</param>
        public Form��������(RegKor.DS1 ds, int id��������, string ������������)
        {
            //// ������ ���� ��������������.
            //this.textBox�������������2.Visible = false;
            bool flagDsp = false;

            InitializeComponent();

            // ������ ���� ��������������.
            this.textBox�������������2.Visible = false;

            // �������� ������ ������� ������������� ��������.
            LoadNumberDepartments();

            this.ds1 = ds;
            DataRow[] dr = ds1.��������.Select("id_��������=" + id��������);
            DataRow[] dr2 = ds1.�������.Select("id_��������=" + id��������);

            // ������� ��������.
            string queryCard = "select * from �������� where id_�������� = "+ id�������� +" ";

            DataRow rowCurrCard = DataTableSql.GetDataTable(queryCard).Rows[0];

            if (rowCurrCard["���"] != DBNull.Value)
            {
                // �������� �������� �� �������� ���������� ���.
                if (Convert.ToBoolean(rowCurrCard["���"]) == true)
                {
                    // ������ �������� �������� ���.
                    flagDsp = true;
                }
            }

            // ��������� � �������� ����� id ��������.
            this.Id�������� = id��������;

            �������������� = (DS1.��������Row)dr[0];


            // ���������� ��� �������� �������� ����������.
            string ���������������� = string.Empty;

            �������������� doc = new ��������������();

            // ��������� ������������� �������� �� ������� �����������������
            DataRow[] dr3 = ds.��������.Select("���������� >='01.12." + ������������ + "'", "������� DESC");
            if (dr3.Length > 0)
            {
                ����������� = 1 + (int)dr3[0]["�������"];
                label1.Text = "����. ����� �\\� " + (�����������);
            }
            else
            {
                ����������� = 1;
                label1.Text = "����. ����� �\\� " + (�����������);
            }

            combo�������������.DataSource = ds1.��������������;
            combo�������������.DisplayMember = ds1.��������������.Columns["����������������������"].ToString();
            combo�������������.ValueMember = ds1.��������������.Columns["id_��������������"].ToString();
            combo�������������.SelectedItem = ��������������["id_��������������"];

            combo��������.DataSource = ds1.���������;
            combo��������.DisplayMember = ds1.���������.Columns["�����������������"].ToString();
            combo��������.ValueMember = ds1.���������.Columns["id_���������"].ToString();

            combo��������.SelectedValue = (int)��������������["id_���������"];
            combo�������������.SelectedValue = (int)��������������["id_��������������"];
            checkBox�����.Checked = (Boolean)��������������["�����"];
            dateTime���������������.Value = Convert.ToDateTime(��������������["����������"]);
            dateTime���������������.Value = Convert.ToDateTime(��������������["����������"]);
            textBox����������.Text = (string)��������������["�����������������"];

            if (flagDsp == false)
            {
                _���������������������������� = dr2[0]["���������"].ToString();
                textBox�������������.Text = dr2[0]["���������"].ToString().Split('/')[1].ToString().Trim();
            }
            else
            {
                _���������������������������� = dr2[0]["���������"].ToString() + "���";
                this.chboxDsp.Checked = true;
                textBox�������������.Text = dr2[0]["���������"].ToString().Split('/')[1].ToString().Trim() + "���";
            }

            //textBox�������������.Text = dr2[0]["���������"].ToString().Split('/')[1].ToString().Trim();
           

            string[] nums = dr2[0]["���������"].ToString().Split('/');

            doc.����� = Convert.ToInt32(nums[0]); //�����������;
            doc.������� = nums[1].ToString();

            ����������������������� = doc;


            textBox��������������.Text = (string)��������������["����������"];
            textBox���������.Text = (string)��������������["���������"];

            // ���� ����� ��������� ���������� ��������� null �� ������ �������� ��������.
            if (��������������["�������������������"] == DBNull.Value)
            {
                ��������������["�������������������"] = "";
            }

            if ((string)��������������["�������������������"] != "")
            {
                textBox�������������������.Enabled = true;
                textBox�������������������.Text = (string)��������������["�������������������"];
            }
            checkBox��������.Checked = (Boolean)��������������["����������"];
            if (��������������["��������������"] != System.DBNull.Value)
            {
                dateTime��������.Value = Convert.ToDateTime(��������������["��������������"]);
            }
            if (checkBox��������.Checked)
            {
                dateTime��������.Enabled = true;
            }

            chBoxRepet.Checked = (Boolean)��������������["FlagCardRepeet"];

            if (this.chBoxRepet.Checked == true)
            {
                this.txtPeriod.Enabled = true;
            }
            else
            {
                this.txtPeriod.Enabled = false;
            }

            if (���������������Config.�����������������������() == true)
            {
                string queryDoc = "select FileDate,FileDateTitlePage from ����������������� " +
                                  "where id_�������� = " + id�������� + " ";
                GetDataTable tab = new GetDataTable(queryDoc);
                DataTable tabRow = tab.DataTable();

                // ������� ���������� ��� �������� ���� � �������� ��������� VipNet.
                pathFileServerTitlePage = null;

                //if (tabRow.Rows[0]["FileDate"].ToString() != "" || tabRow.Rows[0]["FileDate"] != null)
                if (tabRow.Rows.Count > 0)
                {
                    if (tabRow.Rows[0]["FileDate"] != DBNull.Value)
                    {
                        this.linkLabel1.Text = "����������� ���� ���������";

                        ////this.lblFile.Text = "���� ��������� - " + tabRow.Rows[0]["NameFileDocument"].ToString();
                        //this.linkLabel1.Text = "���� ��������� - " + tabRow.Rows[0]["NameFileDocument"].ToString().Split('_')[0].Trim();

                        //if (tabRow.Rows[0]["GuidName"].ToString().Trim() != "")
                        //{
                        //    pathFileServer = tabRow.Rows[0]["NameFileDocument"].ToString().Trim();// +"_" + tabRow.Rows[0]["GuidName"].ToString().Trim();
                        //}
                        //else
                        //{
                        //    pathFileServer = tabRow.Rows[0]["NameFileDocument"].ToString().Trim();
                        //}


                        //������������ = null;
                        //������������ = pathFileServer;

                    }
                    else
                    {
                        //this.lblFile.Text = "";

                        ������������ = null;
                    }
                }
                else
                {
                    ������������ = null;
                }



                //if (tabRow.Rows[0]["MD5"].ToString() == "md5")
                //{
                //    this.chcDop.Checked = true;
                //}
                //else
                //{
                //    this.chcDop.Checked = false;
                //}

                //// ������� ���� � �������.
                //string patchServerQuery = "select PatchServer from ����������";

                //GetDataTable tabServer = new GetDataTable(patchServerQuery);
                //DataTable tabServFile = tabServer.DataTable();

                //patchServerSave = tabServFile.Rows[0]["PatchServer"].ToString().Trim();

                //// ��������� ���� FlagUpdateDocument � true, � ����� � ��� ��� ����� ������� ��� ���������.
                //FlagUpdateDocument = true;

                // ��������� ������ �� ��������� ����.
                //if (tabRow.Rows[0]["FileDateTitlePage"].ToString() != "" || tabRow.Rows[0]["FileDateTitlePage"].ToString() != null)
                if (tabRow.Rows.Count > 0)
                {
                    if (tabRow.Rows[0]["FileDateTitlePage"] != DBNull.Value)
                    {
                        this.linkLabel2.Text = "����������� ���� ���������� �����";

                        //this.linkLabel2.Text = "����������� a��� ���������� ����� - " + tabRow.Rows[0]["NameFileDocumentVipNetEmailTitlePage"].ToString().Split('_')[0].Trim();

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form��������));
            this.label�������ID = new System.Windows.Forms.Label();
            this.label������������������� = new System.Windows.Forms.Label();
            this.label��������� = new System.Windows.Forms.Label();
            this.label���������� = new System.Windows.Forms.Label();
            this.label�������� = new System.Windows.Forms.Label();
            this.label������������� = new System.Windows.Forms.Label();
            this.checkBox����� = new System.Windows.Forms.CheckBox();
            this.button������ = new System.Windows.Forms.Button();
            this.button��������� = new System.Windows.Forms.Button();
            this.groupBox���������� = new System.Windows.Forms.GroupBox();
            this.label�������������� = new System.Windows.Forms.Label();
            this.label��������������� = new System.Windows.Forms.Label();
            this.dateTime��������������� = new System.Windows.Forms.DateTimePicker();
            this.textBox�������������� = new System.Windows.Forms.TextBox();
            this.textBox��������� = new System.Windows.Forms.TextBox();
            this.textBox���������� = new System.Windows.Forms.TextBox();
            this.dateTime�������� = new System.Windows.Forms.DateTimePicker();
            this.checkBox�������� = new System.Windows.Forms.CheckBox();
            this.combo�������� = new System.Windows.Forms.ComboBox();
            this.combo������������� = new System.Windows.Forms.ComboBox();
            this.groupBox��������� = new System.Windows.Forms.GroupBox();
            this.rb04 = new System.Windows.Forms.RadioButton();
            this.rb02 = new System.Windows.Forms.RadioButton();
            this.chboxDsp = new System.Windows.Forms.CheckBox();
            this.textBox�������������2 = new System.Windows.Forms.TextBox();
            this.textBox������������� = new System.Windows.Forms.MaskedTextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label������������� = new System.Windows.Forms.Label();
            this.label��������������� = new System.Windows.Forms.Label();
            this.dateTime��������������� = new System.Windows.Forms.DateTimePicker();
            this.ds1 = new RegKor.DS1();
            this.textBox������������������� = new System.Windows.Forms.TextBox();
            this.panelKontrol = new System.Windows.Forms.Panel();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.button������������������� = new System.Windows.Forms.Button();
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
            this.groupBox����������.SuspendLayout();
            this.groupBox���������.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ds1)).BeginInit();
            this.panelKontrol.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label�������ID
            // 
            this.label�������ID.Location = new System.Drawing.Point(516, 2);
            this.label�������ID.Name = "label�������ID";
            this.label�������ID.Size = new System.Drawing.Size(26, 24);
            this.label�������ID.TabIndex = 0;
            // 
            // label�������������������
            // 
            this.label�������������������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label�������������������.Location = new System.Drawing.Point(4, 404);
            this.label�������������������.Name = "label�������������������";
            this.label�������������������.Size = new System.Drawing.Size(172, 14);
            this.label�������������������.TabIndex = 0;
            this.label�������������������.Text = "��������� ����������";
            this.label�������������������.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label���������
            // 
            this.label���������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label���������.Location = new System.Drawing.Point(4, 344);
            this.label���������.Name = "label���������";
            this.label���������.Size = new System.Drawing.Size(172, 13);
            this.label���������.TabIndex = 0;
            this.label���������.Text = "���������";
            this.label���������.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label����������
            // 
            this.label����������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label����������.Location = new System.Drawing.Point(4, 284);
            this.label����������.Name = "label����������";
            this.label����������.Size = new System.Drawing.Size(172, 14);
            this.label����������.TabIndex = 0;
            this.label����������.Text = "������� ����������";
            this.label����������.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label��������
            // 
            this.label��������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label��������.Location = new System.Drawing.Point(8, 32);
            this.label��������.Name = "label��������";
            this.label��������.Size = new System.Drawing.Size(108, 22);
            this.label��������.TabIndex = 0;
            this.label��������.Text = "��� ���������";
            this.label��������.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label�������������
            // 
            this.label�������������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label�������������.Location = new System.Drawing.Point(10, 3);
            this.label�������������.Name = "label�������������";
            this.label�������������.Size = new System.Drawing.Size(108, 24);
            this.label�������������.TabIndex = 0;
            this.label�������������.Text = "�������������";
            this.label�������������.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // checkBox�����
            // 
            this.checkBox�����.Checked = true;
            this.checkBox�����.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox�����.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox�����.Location = new System.Drawing.Point(354, 215);
            this.checkBox�����.Name = "checkBox�����";
            this.checkBox�����.Size = new System.Drawing.Size(118, 24);
            this.checkBox�����.TabIndex = 12;
            this.checkBox�����.Text = "� ����";
            this.checkBox�����.CheckedChanged += new System.EventHandler(this.checkBox�����_CheckedChanged);
            // 
            // button������
            // 
            this.button������.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button������.Dock = System.Windows.Forms.DockStyle.Right;
            this.button������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button������.Location = new System.Drawing.Point(370, 0);
            this.button������.Name = "button������";
            this.button������.Size = new System.Drawing.Size(178, 28);
            this.button������.TabIndex = 17;
            this.button������.Text = "������";
            this.toolTip1.SetToolTip(this.button������, "������� ���� ��� ���������� ���������");
            this.button������.Click += new System.EventHandler(this.button������_Click);
            // 
            // button���������
            // 
            this.button���������.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.button���������.Dock = System.Windows.Forms.DockStyle.Left;
            this.button���������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button���������.Location = new System.Drawing.Point(0, 0);
            this.button���������.Name = "button���������";
            this.button���������.Size = new System.Drawing.Size(178, 28);
            this.button���������.TabIndex = 16;
            this.button���������.Text = "���������";
            this.toolTip1.SetToolTip(this.button���������, "��������� ��������� � ������� ����");
            this.button���������.Click += new System.EventHandler(this.button���������_Click);
            // 
            // groupBox����������
            // 
            this.groupBox����������.Controls.Add(this.label��������������);
            this.groupBox����������.Controls.Add(this.label���������������);
            this.groupBox����������.Controls.Add(this.dateTime���������������);
            this.groupBox����������.Controls.Add(this.textBox��������������);
            this.groupBox����������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBox����������.Location = new System.Drawing.Point(8, 56);
            this.groupBox����������.Name = "groupBox����������";
            this.groupBox����������.Size = new System.Drawing.Size(258, 103);
            this.groupBox����������.TabIndex = 3;
            this.groupBox����������.TabStop = false;
            this.groupBox����������.Text = "����������";
            // 
            // label��������������
            // 
            this.label��������������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label��������������.Location = new System.Drawing.Point(10, 56);
            this.label��������������.Name = "label��������������";
            this.label��������������.Size = new System.Drawing.Size(94, 20);
            this.label��������������.TabIndex = 0;
            this.label��������������.Text = "��������� �";
            this.label��������������.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label���������������
            // 
            this.label���������������.Location = new System.Drawing.Point(10, 22);
            this.label���������������.Name = "label���������������";
            this.label���������������.Size = new System.Drawing.Size(92, 20);
            this.label���������������.TabIndex = 0;
            this.label���������������.Text = "����";
            this.label���������������.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dateTime���������������
            // 
            this.dateTime���������������.CalendarTrailingForeColor = System.Drawing.SystemColors.Control;
            this.dateTime���������������.Location = new System.Drawing.Point(106, 22);
            this.dateTime���������������.Name = "dateTime���������������";
            this.dateTime���������������.Size = new System.Drawing.Size(144, 22);
            this.dateTime���������������.TabIndex = 4;
            // 
            // textBox��������������
            // 
            this.textBox��������������.BackColor = System.Drawing.SystemColors.HighlightText;
            this.textBox��������������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox��������������.Location = new System.Drawing.Point(106, 56);
            this.textBox��������������.MaxLength = 20;
            this.textBox��������������.Name = "textBox��������������";
            this.textBox��������������.Size = new System.Drawing.Size(144, 22);
            this.textBox��������������.TabIndex = 5;
            this.textBox��������������.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox���������
            // 
            this.textBox���������.BackColor = System.Drawing.SystemColors.HighlightText;
            this.textBox���������.Enabled = false;
            this.textBox���������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox���������.Location = new System.Drawing.Point(2, 359);
            this.textBox���������.MaxLength = 250;
            this.textBox���������.Multiline = true;
            this.textBox���������.Name = "textBox���������";
            this.textBox���������.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox���������.Size = new System.Drawing.Size(512, 42);
            this.textBox���������.TabIndex = 0;
            this.textBox���������.TabStop = false;
            // 
            // textBox����������
            // 
            this.textBox����������.BackColor = System.Drawing.SystemColors.HighlightText;
            this.textBox����������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox����������.Location = new System.Drawing.Point(2, 300);
            this.textBox����������.MaxLength = 250;
            this.textBox����������.Multiline = true;
            this.textBox����������.Name = "textBox����������";
            this.textBox����������.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox����������.Size = new System.Drawing.Size(540, 42);
            this.textBox����������.TabIndex = 13;
            // 
            // dateTime��������
            // 
            this.dateTime��������.Enabled = false;
            this.dateTime��������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dateTime��������.Location = new System.Drawing.Point(136, 4);
            this.dateTime��������.Name = "dateTime��������";
            this.dateTime��������.Size = new System.Drawing.Size(144, 22);
            this.dateTime��������.TabIndex = 11;
            // 
            // checkBox��������
            // 
            this.checkBox��������.BackColor = System.Drawing.SystemColors.Control;
            this.checkBox��������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkBox��������.Location = new System.Drawing.Point(4, 4);
            this.checkBox��������.Name = "checkBox��������";
            this.checkBox��������.Size = new System.Drawing.Size(126, 24);
            this.checkBox��������.TabIndex = 10;
            this.checkBox��������.Text = "�� �������� ��";
            this.checkBox��������.UseVisualStyleBackColor = false;
            this.checkBox��������.CheckedChanged += new System.EventHandler(this.checkBox��������_CheckedChanged);
            // 
            // combo��������
            // 
            this.combo��������.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.combo��������.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combo��������.DisplayMember = "id_���������";
            this.combo��������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.combo��������.Location = new System.Drawing.Point(124, 30);
            this.combo��������.MaxDropDownItems = 40;
            this.combo��������.Name = "combo��������";
            this.combo��������.Size = new System.Drawing.Size(372, 24);
            this.combo��������.TabIndex = 2;
            this.combo��������.ValueMember = "id_���������";
            // 
            // combo�������������
            // 
            this.combo�������������.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.combo�������������.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.combo�������������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.combo�������������.Location = new System.Drawing.Point(124, 1);
            this.combo�������������.MaxDropDownItems = 40;
            this.combo�������������.Name = "combo�������������";
            this.combo�������������.Size = new System.Drawing.Size(372, 24);
            this.combo�������������.TabIndex = 1;
            // 
            // groupBox���������
            // 
            this.groupBox���������.Controls.Add(this.rb04);
            this.groupBox���������.Controls.Add(this.rb02);
            this.groupBox���������.Controls.Add(this.chboxDsp);
            this.groupBox���������.Controls.Add(this.textBox�������������2);
            this.groupBox���������.Controls.Add(this.textBox�������������);
            this.groupBox���������.Controls.Add(this.label1);
            this.groupBox���������.Controls.Add(this.label�������������);
            this.groupBox���������.Controls.Add(this.label���������������);
            this.groupBox���������.Controls.Add(this.dateTime���������������);
            this.groupBox���������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBox���������.Location = new System.Drawing.Point(284, 60);
            this.groupBox���������.Name = "groupBox���������";
            this.groupBox���������.Size = new System.Drawing.Size(258, 145);
            this.groupBox���������.TabIndex = 6;
            this.groupBox���������.TabStop = false;
            this.groupBox���������.Text = "���������";
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
            this.chboxDsp.Text = "���";
            this.chboxDsp.UseVisualStyleBackColor = true;
            this.chboxDsp.CheckedChanged += new System.EventHandler(this.chboxDsp_CheckedChanged);
            // 
            // textBox�������������2
            // 
            this.textBox�������������2.Location = new System.Drawing.Point(112, 75);
            this.textBox�������������2.Name = "textBox�������������2";
            this.textBox�������������2.Size = new System.Drawing.Size(133, 22);
            this.textBox�������������2.TabIndex = 11;
            this.textBox�������������2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox�������������
            // 
            this.textBox�������������.Location = new System.Drawing.Point(107, 75);
            this.textBox�������������.Mask = "00-00-00";
            this.textBox�������������.Name = "textBox�������������";
            this.textBox�������������.Size = new System.Drawing.Size(123, 22);
            this.textBox�������������.TabIndex = 10;
            this.textBox�������������.TextChanged += new System.EventHandler(this.textBox�������������_TextChanged);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(6, 102);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(239, 18);
            this.label1.TabIndex = 9;
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label�������������
            // 
            this.label�������������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label�������������.Location = new System.Drawing.Point(12, 76);
            this.label�������������.Name = "label�������������";
            this.label�������������.Size = new System.Drawing.Size(88, 20);
            this.label�������������.TabIndex = 0;
            this.label�������������.Text = "�������� �";
            this.label�������������.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label���������������
            // 
            this.label���������������.Location = new System.Drawing.Point(12, 22);
            this.label���������������.Name = "label���������������";
            this.label���������������.Size = new System.Drawing.Size(88, 20);
            this.label���������������.TabIndex = 0;
            this.label���������������.Text = "����";
            this.label���������������.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dateTime���������������
            // 
            this.dateTime���������������.CalendarTrailingForeColor = System.Drawing.SystemColors.Control;
            this.dateTime���������������.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.ds1, "�������.����������", true));
            this.dateTime���������������.Location = new System.Drawing.Point(102, 22);
            this.dateTime���������������.Name = "dateTime���������������";
            this.dateTime���������������.Size = new System.Drawing.Size(144, 22);
            this.dateTime���������������.TabIndex = 7;
            // 
            // ds1
            // 
            this.ds1.DataSetName = "DS";
            this.ds1.Locale = new System.Globalization.CultureInfo("ru-RU");
            this.ds1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // textBox�������������������
            // 
            this.textBox�������������������.Enabled = false;
            this.textBox�������������������.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBox�������������������.Location = new System.Drawing.Point(2, 419);
            this.textBox�������������������.MaxLength = 250;
            this.textBox�������������������.Multiline = true;
            this.textBox�������������������.Name = "textBox�������������������";
            this.textBox�������������������.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBox�������������������.Size = new System.Drawing.Size(512, 42);
            this.textBox�������������������.TabIndex = 15;
            // 
            // panelKontrol
            // 
            this.panelKontrol.BackColor = System.Drawing.SystemColors.Control;
            this.panelKontrol.Controls.Add(this.dateTime��������);
            this.panelKontrol.Controls.Add(this.checkBox��������);
            this.panelKontrol.Location = new System.Drawing.Point(6, 211);
            this.panelKontrol.Name = "panelKontrol";
            this.panelKontrol.Size = new System.Drawing.Size(290, 30);
            this.panelKontrol.TabIndex = 9;
            // 
            // button�������������������
            // 
            this.button�������������������.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.button�������������������.Image = global::RegKor.Properties.Resources.add;
            this.button�������������������.Location = new System.Drawing.Point(520, 367);
            this.button�������������������.Name = "button�������������������";
            this.button�������������������.Size = new System.Drawing.Size(22, 22);
            this.button�������������������.TabIndex = 14;
            this.toolTip1.SetToolTip(this.button�������������������, "�������� ���, ������� ��������� ��������");
            this.button�������������������.Click += new System.EventHandler(this.button�������������������_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.btnTie);
            this.panel1.Controls.Add(this.button���������);
            this.panel1.Controls.Add(this.button������);
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
            this.btnTie.Text = "�������� ���������";
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
            this.btnElementPS.Text = "������ ������������ ������";
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
            this.chBoxRepet.Text = "����� � ��������������";
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
            this.label2.Text = "������������� ����";
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
            this.chekDocServer.Text = "��������� ����� ���������";
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
            this.chcDop.Text = "��������";
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
            this.btnLastNumber.Text = "��������� ��������";
            this.btnLastNumber.UseVisualStyleBackColor = true;
            this.btnLastNumber.Click += new System.EventHandler(this.btnLastNumber_Click);
            // 
            // Form��������
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.button������;
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
            this.Controls.Add(this.button�������������������);
            this.Controls.Add(this.panelKontrol);
            this.Controls.Add(this.textBox�������������������);
            this.Controls.Add(this.label�������ID);
            this.Controls.Add(this.label�������������������);
            this.Controls.Add(this.label���������);
            this.Controls.Add(this.label����������);
            this.Controls.Add(this.label��������);
            this.Controls.Add(this.label�������������);
            this.Controls.Add(this.checkBox�����);
            this.Controls.Add(this.groupBox����������);
            this.Controls.Add(this.textBox���������);
            this.Controls.Add(this.textBox����������);
            this.Controls.Add(this.combo��������);
            this.Controls.Add(this.combo�������������);
            this.Controls.Add(this.groupBox���������);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(554, 436);
            this.Name = "Form��������";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "�������� ���������";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.Form��������_FormClosed);
            this.Shown += new System.EventHandler(this.Form��������_Shown);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form��������_FormClosing);
            this.Load += new System.EventHandler(this.Form��������_Load);
            this.groupBox����������.ResumeLayout(false);
            this.groupBox����������.PerformLayout();
            this.groupBox���������.ResumeLayout(false);
            this.groupBox���������.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ds1)).EndInit();
            this.panelKontrol.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion
        private void checkBox��������_CheckedChanged(object sender, System.EventArgs e)
        {
            if (checkBox��������.Checked)
            {
                dateTime��������.Enabled = true;
            }
            else
            {
                dateTime��������.Enabled = false;
            }
        }

        private void checkBox�����_CheckedChanged(object sender, System.EventArgs e)
        {
            if (checkBox�����.Checked)
            {
                textBox�������������������.Enabled = true;
                textBox�������������������.Text = ���������������������;
                panelKontrol.Enabled = false;
            }
            else
            {
                textBox�������������������.Enabled = false;
                ��������������������� = textBox�������������������.Text;
                textBox�������������������.Text = "";
                panelKontrol.Enabled = true;
            }
        }

        private void button�������������������_Click(object sender, System.EventArgs e)
        {
            Form��������� form = new Form���������();
            DialogResult result = form.ShowDialog(this);

            if (result == DialogResult.OK)
            {
                textBox���������.Text = form.���������������;

                // ������� ������ ����������� ������� � ���������� ������� ������� ��������.
                this.ListPerson = form.ListPerson;

                string sTest = "";
            }
        }

        private void Form��������_Shown(object sender, EventArgs e)
        {
            combo�������������.Focus();
        }


        /// <summary>
        /// ��������� ���������� ����������, 
        /// ���� � ������� ���� ����� �� ���������� ����������� �� ����������, 
        /// � ���� ������ ���������� ������, �� � ���������� ����������� �� ��������, 
        /// � ������ ��� ���� ��������
        /// </summary>
        private void ��������������������()
        {
            Form��������� form;
            string ������������� = combo��������.Text;

            if (������������� != "")
            {
                form = new Form���������(�������������);
                form.Text = "���������� \"���������\"";
            }
            else
            {
                form = new Form���������();
                form.Text = "���������� \"���������\"";
            }

            form.ShowDialog(this);
            ds1.���������.Clear();
            DS1TableAdapters.���������TableAdapter adapter = new RegKor.DS1TableAdapters.���������TableAdapter();
            adapter.Fill(ds1.���������);
            combo��������.DataSource = null;
            combo��������.DisplayMember = "";
            combo��������.ValueMember = "";
            combo��������.DataSource = ds1.���������;
            combo��������.DisplayMember = ds1.���������.Columns["�����������������"].ToString();
            combo��������.ValueMember = ds1.���������.Columns["id_���������"].ToString();
            combo��������.Text = �������������;
            if (combo��������.Text != �������������)
            {
                combo��������.Text = "";
                combo��������.SelectedText = �������������;
            }
        }

        /// <summary>
        /// ��������� ���������� ���������������, 
        /// ���� � ������� ���� ����� �� ���������� ����������� �� ����������, 
        /// � ���� ������ ��������������� ������, 
        /// �� � ���������� ����������� �� ��������, 
        /// � ������ ��� ���� ��������
        /// </summary>
        private void �������������������������()
        {
            Form�������������� form;
            string ������������������ = combo�������������.Text;

            if (������������������ != "")
            {
                form = new Form��������������(������������������);
                form.Text = "���������� \"��������������\"";
            }
            else
            {
                form = new Form��������������();
                form.Text = "���������� \"��������������\"";
            }

            form.ShowDialog(this);
            ds1.��������������.Clear();
            DS1TableAdapters.��������������TableAdapter adapter = new RegKor.DS1TableAdapters.��������������TableAdapter();
            adapter.Fill(ds1.��������������);
            combo�������������.DataSource = null;
            combo�������������.DisplayMember = "";
            combo�������������.ValueMember = "";
            combo�������������.DataSource = ds1.��������������;
            combo�������������.DisplayMember = ds1.��������������.Columns["����������������������"].ToString();
            combo�������������.ValueMember = ds1.��������������.Columns["id_��������������"].ToString();
            combo�������������.Text = ������������������;
            if (combo�������������.Text != ������������������)
            {
                combo�������������.Text = "";
                combo�������������.SelectedText = ������������������;
            }
        }

        private void textBox�������������_Enter(object sender, EventArgs e)
        {
            textBox�������������.Select(textBox�������������.Text.Length, 0);
        }


        private void button������_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        private void button���������_Click(object sender, System.EventArgs e)
        {
            #region ��������
            // ������������� id ���������:
            DataRow[] rows = ds1.���������.Select("�����������������='" + combo��������.Text.Trim() + "'");
            if (rows.Length > 0)
            {
                ��������������["id_���������"] = (int)combo��������.SelectedValue;
            }
            else if (combo��������.Text != "")
            {
                DialogResult result = MessageBox.Show(this,
                    "�� ������� ��������, ������� �� ��������������� � ����������� \"���������\". ����� ��������� ��� � ���������� ��� ���?",
                    "����������� ��� ���������",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);
                if (result == DialogResult.No)
                {// ���� ���� ������ ���, ��������� ���������� � ������� �� ���������
                    this.DialogResult = DialogResult.None;
                    return;
                }
                if (result == DialogResult.Yes)
                {// ���� ���� ������ ��, ��������� ����������, ��������� ���������� � ������� �� ���������
                    ��������������������();
                    this.DialogResult = DialogResult.None;
                    return;
                }
            }
            else if (combo��������.Text.Trim() == "")
            {
                MessageBox.Show(this,
                "�� �� ������� ��� ���������",
                "��� ���������",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.None;
                return;
            }
            #endregion

            #region �������������
            // ������������� id ��������������:
            DataRow[] rows2 = ds1.��������������.Select("����������������������='" + combo�������������.Text.Trim() + "'");
            if (rows2.Length > 0)
            {
                ��������������["id_��������������"] = (int)combo�������������.SelectedValue;
            }
            else if (combo�������������.Text != "")
            {
                DialogResult result = MessageBox.Show(this,
                    "�� ������� ��������������, ������� �� ��������������� � ����������� \"��������������\". ����� ��������� ��� � ���������� ��� ���?",
                    "����������� �������������",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question,
                    MessageBoxDefaultButton.Button1);
                if (result == DialogResult.No)
                {// ���� ���� ������ ���, ��������� ���������� � ������� �� ���������
                    this.DialogResult = DialogResult.None;
                    return;
                }
                if (result == DialogResult.Yes)
                {// ���� ���� ������ ��, ��������� ����������, ��������� ���������� � ������� �� ���������
                    �������������������������();
                    this.DialogResult = DialogResult.None;
                    return;
                }
            }
            else if (combo�������������.Text.Trim() == "")
            {
                MessageBox.Show(this,
                "�� �� ������� ��������������",
                "�������������",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
                this.DialogResult = DialogResult.None;
                return;
            }
            #endregion

            #region � ����, �� ��������

            ��������������["�����"] = (Boolean)checkBox�����.Checked;
            ��������������["��������������"] = dateTime��������.Value.ToShortDateString();
            ��������������["����������"] = checkBox��������.Checked;

            if ((checkBox�����.Checked == true && checkBox��������.Checked == true && textBox�������������������.Text.Trim().Length > 0) || (checkBox�����.Checked == true && checkBox��������.Checked == false && textBox�������������������.Text.Trim().Length > 0))
            {
                ��������������["�����"] = true;
                ��������������["����������"] = true;
            }

            #endregion

            #region ����� ��������
            if (textBox�������������.Text == "�/�" || textBox�������������.Text == "�.�" || textBox�������������.Text == "�.�." || textBox�������������.Text == "��")
            {
                ��������������["���������"] = "�/�";
            }
            else
            {
                // �������� ������������ ��������� ������.
                string[] arr = textBox�������������.Text.Split('-');

                bool flagErrorNuber = true;
                foreach (string sKey in this.numbersDepartment.Keys)
                {
                    //if(arr[3].Trim() == sKey.Trim())
                    if (arr[2].Trim() == sKey.Trim())
                    {
                        // ��������� ���� ������ � true.
                        flagErrorNuber = false;
                    }
                }

                // ���� ������ �� ������� �� ���� ������������.
                if (flagErrorNuber == true)
                {
                    MessageBox.Show(this,
                           "������� ������ ����� ��������",
                           "������ ������",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error);
                    this.DialogResult = DialogResult.None;
                    return;
                }

                // ������ ���������� ��������� ������ ���������, ������� ����� ������������.
                //string[] arr = textBox�������������.Text.Split('/');


                //if (arr.Length != 2)
                //{
                //    MessageBox.Show(
                //                        this,
                //                       "������� ������ ����� ��������",
                //                       "������ ������",
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
                //        if (Convert.ToInt32(arr[0]) > �����������)
                //        {
                //            MessageBox.Show(this,
                //               "������� ������ ���������� �������� �����. �� ������ ������� ����� �� ������ ��� " + �����������,
                //               "������ ������",
                //               MessageBoxButtons.OK,
                //               MessageBoxIcon.Error);
                //            this.DialogResult = DialogResult.None;
                //            return;

                //        }
                //        else if (Convert.ToInt32(arr[0]) < ����������� && �������������)
                //        {
                //            DialogResult result = MessageBox.Show(this,
                //                "�� ������� ���������� �������� �����, ������� �� ������������� ��������������.\n���� �� �������� ��������� �����, �������� ������������ ������� � ���� ������",
                //                "�������� ������������ �������",
                //                MessageBoxButtons.YesNo,
                //                MessageBoxIcon.Warning,
                //                MessageBoxDefaultButton.Button2);
                //            if (result == DialogResult.No)
                //            {
                //                this.DialogResult = DialogResult.None;
                //                return;
                //            }
                //        }


                //        //�����������������������
                //        ��������������["�������"] = arr[0];
                //        ��������������["���������"] = arr[1];
                //    }
                //    else
                //    {
                //        MessageBox.Show(this,
                //           "������� ������ ����� ��������",
                //           "������ ������",
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

                        // �������� ���� ����� ������ ��������� �������.
                        //if (this.flagNumberDoc.Checked == true)
                        //{
                        //    FormNumberDoc fNumDoc = new FormNumberDoc();
                        //    fNumDoc.ShowDialog();

                        //    if (fNumDoc.DialogResult == DialogResult.OK)
                        //    {
                        //        ��������������["�������"] = fNumDoc.NumberDoc;
                        //    }
                        //    else if (fNumDoc.DialogResult == DialogResult.Cancel)
                        //    {
                        //        return;
                        //    }
                        //}
                        //else
                        //{
                        //    ��������������["�������"] = �����������������������.�����;
                        //}

                        ��������������["�������"] = �����������������������.�����;
                        //��������������["���������"] = �����������������������.������� + textBox�������������.Text.Trim();
                        ��������������["���������"] = textBox�������������.Text.Trim();

                        // �������� � �������� ����� ����� ������ ���������.
                        �������������� nextNumDoc = this.�����������������������;

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
                        //        // ������������ ��������� ������ ���� ��������� ����������.
                        //        if (this.flagLastNumberDoc == false)
                        //        {
                        //            nextNumDoc.����� = Convert.ToInt16(fNumDoc.NumberDoc);
                        //            ��������������["�������"] = Convert.ToInt16(fNumDoc.NumberDoc);
                        //        }
                        //        else
                        //        {
                        //            // �������������� �������������� ��������� ���������� ����� ������� �����.
                        //            fNumDoc.NumberDoc = string.Empty;

                        //            // ��������� ���������� ���������� ����������.
                        //            fNumDoc.NumberDoc = this.lastNumberDoc;

                        //            // ��������� �����������.


                        //            nextNumDoc.����� = Convert.ToInt16(fNumDoc.NumberDoc);
                        //            ��������������["�������"] = Convert.ToInt16(fNumDoc.NumberDoc);
                        //        }
                        //    }
                        //}
                        //else
                        //{
                        //    nextNumDoc.����� = �����������������������.�����;
                        //}

                        //nextNumDoc.������� = �����������������������.������� + textBox�������������.Text.Trim();
                        nextNumDoc.������� = textBox�������������.Text.Trim();

                        // �������� ��� ����������� ������������ ��������� ����� ���������.
                        ����������������������� = nextNumDoc;
                    }
                    else
                    {
                        // �������� � �������� ����� ����� ������ ���������.
                        �������������� nextNumDoc = this.�����������������������;

                        // ������ ����� � ������� ���������.
                        this.textBox�������������2.Text.Split('/')[0].Trim();

                        string[] arrayNum  = this.textBox�������������2.Text.Split('/');

                        nextNumDoc.����� = Convert.ToInt32(arrayNum[0]);

                        //nextNumDoc.������� = �����������������������.������� + textBox�������������.Text.Trim();
                        nextNumDoc.������� = arrayNum[1].Trim();

                        // �������� ��� ����������� ������������ ��������� ����� ���������.
                        ����������������������� = nextNumDoc;
                    }
                }
                else
                {
                    ��������������["�������"] = �����������������������.�����;

                    string[] arry = textBox�������������.Text.Trim().Split('/');

                    // ������ ����������.
                    //��������������["���������"] = arry[1].Trim();
                    ��������������["���������"] = arry[0].Trim();

                    string iTest = ��������������["���������"].ToString().Trim();


                    // �������� � �������� ����� ����� ������ ���������.
                    �������������� nextNumDoc = new ��������������();
                    nextNumDoc.FlagUpdate = true;
                    nextNumDoc.����� = �����������������������.�����;
                    nextNumDoc.������� = textBox�������������.Text.Trim();

                    // �������� ��� ����������� ������������ ��������� ����� ���������.
                    ����������������������� = nextNumDoc;
                }
            }
            #endregion

            #region ����� ���������
            if (textBox��������������.Text == "")
            {
                ��������������["����������"] = "�/�";
            }
            else
            {
                ��������������["����������"] = textBox��������������.Text;
            }
            #endregion

            ��������������["����������"] = dateTime���������������.Value.ToShortDateString();

            ��������������["����������"] = dateTime���������������.Value.ToShortDateString();

            ��������������["�����������������"] = textBox����������.Text;

            ��������������["���������"] = textBox���������.Text;

            string myTest = textBox�������������������.Text;

            ��������������["�������������������"] = textBox�������������������.Text;

            ��������������["FlagPersonData"] = false;

            ��������������.FlagCardRepeet = this.chBoxRepet.Checked;

            //// ��������� � �������� id ���� ��������� ������.
            //������������ coonectDB = new ������������();
            //string sConn = coonectDB.�����������������();

            //string query = "select [id_�����������������������]from ������������������������������� " +
            //               "where [�������������������������������] = '" + this.cmbBox.Text + "' ";

            //GetDataTable getTable = new GetDataTable(query);
            //DataTable tabPD = getTable.DataTable();

            //this.Id������������������������������� = Convert.ToInt32(tabPD.Rows[0][0]);

            // �������� � ����� �������� ����� ������������, ��� ������ ������ ����� ����� ������������� ��������� ������.


            string ttt = dateTime��������.Value.ToShortDateString();

            this.FlagRecordRepeet = this.chBoxRepet.Checked;

            if (this.chBoxRepet.Checked == true)
            {
                this.IncrementDate = Convert.ToInt32(txtPeriod.Text.Trim()) - 2;

                // �������� ���������� ����.
                ��������������["��������������"] = dateTime��������.Value.AddDays(this.IncrementDate);

                string sTestDate = Convert.ToDateTime(��������������["��������������"]).ToShortDateString();
            }

            // ���� ������ ��� �������� �� �������� ������������ ���������.
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

            // ������� ���� ������� �������� ������ ����� ����������� ����������.
            FormTypeCompanyDocument formType = new FormTypeCompanyDocument();
            formType.ShowDialog();

            if (formType.DialogResult == DialogResult.OK)
            {
                // ��������� � ����� ������ ����������� ���������.
                ����������������� = formType.�����������������;
            }

            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Form���������������� form = new Form����������������();
            //form.Id�������� = this.Id��������;
            //form.TopMost = true;
            //form.ShowDialog();

            //if (form.DialogResult == DialogResult.OK)
            //{
            //    this.QueryString�������������� = form.�������������������������������;
            //}

            Form������������������ form = new Form������������������();
            form.Id�������� = this.Id��������;
            form.TopMost = true;
            form.ShowDialog();

            if (form.DialogResult == DialogResult.OK)
            {
                this.ConfigDatePerosnal = form.����������������������������������;
            }

        }

        private void Form��������_Load(object sender, EventArgs e)
        {
            //// �������� �������������� ������: ���� ��������� ������������ ������ ����������� �� ���� ������.
            // // ������ ����������� ����������� ������������ ������.
            //������������ coonectDB = new ������������();
            //string sConn = coonectDB.�����������������();

            //string query = "select [id_�����������������������],[�������������������������������] from �������������������������������";

            //GetDataTable getTable = new GetDataTable(query);
            //DataTable tabPD = getTable.DataTable();

            //this.cmbBox.DataSource = tabPD;
            //this.cmbBox.DisplayMember = "�������������������������������";
            //this.cmbBox.ValueMember = "id_�����������������������";
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            //�������������������� ������� = new ��������������������();
            //�������.������� = true;
            //�������.������������� = "NULL";

            //// �������� �������� ������ �������������������� � ��������.
            //this.������������� = �������;

            //this.lblMarks.Text = "��������";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Form����� fo = new Form�����();
            //fo.ShowDialog();

            //if (fo.DialogResult == DialogResult.OK)
            //{
            //    �������������������� ������� = new ��������������������();
            //    �������.������� = false;
            //    �������.������������� = fo.�����������;

            //    // �������� �������� ������ �������������������� � ��������.
            //    this.������������� = �������;


            //    this.lblMarks.Text = "��������";
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
                // ������� ���� � ����� ������ ������� ����� ������� ����� � ������� ����������.
                string patchDir = ConfigurationSettings.AppSettings["����������������������������"].Trim();

                //// �������� ���������� ��� �������� ���������.
                string nameDir = this._����������������������������.Trim().Replace("/", "-") + "-id" + this.Id��������.ToString().Trim();

                // ������� ���������� � �������� ��������.
                DirectoryInfo dirInfo = new DirectoryInfo(patchDir);

                // �������� ��������������.
                dirInfo.CreateSubdirectory(nameDir);

                string query = "update �������� " +
                                "set NameFileDocument = NULL, " +
                                "DataWriterServerDoc = NULL, " +
                                "NameFileDocumentVipNetEmailTitlePage = NULL " +
                                "where id_�������� = " + this.Id�������� + " ";

                ExecuteQuery exq = new ExecuteQuery(query);
                exq.Excecute();

                MessageBox.Show("���� �������");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            #region ������ ����������

            //// ������ ��� ����� ���������� ���� �� ������.
            //���������������� = true;

            //// ��������� �������� ����� ���������� ����������� ���� � true.
            //this.SaveDocServer = true;

            //string docServer = ConfigurationSettings.AppSettings["���������������"].ToString();

            //// ������� ���� ��������� �������.
            //FolderBrowserDialog openFileDialog1 = new FolderBrowserDialog();

            ////OpenFileDialog openFileDialog1 = new OpenFileDialog();

            //if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //{
            //    // ������� ���� � �����.
            //    fileName = openFileDialog1.SelectedPath;

            //    // ������� ��� ���� ����� ����� ������ ��������� � ������������� ����������.
            //    DirectoryInfo dif = new DirectoryInfo(fileName);
            //    if (dif.GetFiles().Length == 0)
            //    {
            //        MessageBox.Show("���������� ����� �����!", "��������", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            //        // � ������ ���� ����� � ���������� ������, ������ ������ ���������.
            //        ���������������� = false;

            //        return;
            //    }

            //    // ������� � �������� �������� �����.
            //    FileName = fileName.Trim();

            //    string ���������� = Path.GetExtension(fileName);

            //    // �������� ����� ��� �����.
            //    string newFileName = Guid.NewGuid().ToString().Trim() + ����������;

            //    // ���� ����� ����������� �� ������.
            //    //this.PathFileServer = docServer + Path.GetFileName(newFileName).Trim();
            //    this.PathFileServer = newFileName;

            //}

            #endregion
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ������������ connDb = new ������������();

            using (SqlConnection connection = new SqlConnection(connDb.�����������������()))
            {
                connection.Open();
                string sql = "select FileDate from ����������������� where id_�������� = " + this.Id�������� + "";
                SqlCommand command = new SqlCommand(sql, connection);
                
                SqlDataReader reader = command.ExecuteReader();


                while (reader.Read())
                {
                    fileByteArray = (byte[])reader["FileDate"];
                    //fileByteArray = (byte[])command.ExecuteScalar();
                    //fileByteArray = (byte[])reader["FileDate"];
                }

                // ������ ����� ������ �� ��.
                //byte[] fileByteArray = (byte[])command.ExecuteScalar();

                string dir = @"d:\Recor";

                string fileName = dir + @"\TempView.zip";

                FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                BinaryWriter binWriter = new BinaryWriter(fileStream);
                binWriter.Write(fileByteArray);
                binWriter.Close();

                // ������� �����.
                System.Diagnostics.Process.Start(fileName);
            }

            //// ������� ��� ����� �� �������.
            //string ���������������� = this.PathFileServer.Trim() + ".zip";

            //try
            //{

            //    // ���� � ��������.
            //    string ����������� = patchServerSave.Trim();

            //    // ������� ���� � ����� �� �������.
            //    string fileServer = ����������� + @"\" + ����������������.Replace("/", "-");

            //    // ���� � ����� �� ��������� �����.
            //    string tempPath = Path.GetTempPath();

            //    // ������� ���� � ��� ����� ������� �� ����� ����� ����� ����������� �� ��������� ������� �� ������.
            //    string fileTo = tempPath + ����������������;

            //    // ��������� ����� �� ��������� �����.
            //    File.Copy(fileServer, fileTo, true);

            //    // ������� ���� � ����� �� ��������� ������� �� �������.
            //    string fileTemp = tempPath + @"\" + ����������������;

            //    // ������� �����.
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
                ������������ connDb = new ������������();

                using (SqlConnection connection = new SqlConnection(connDb.�����������������()))
                {
                    connection.Open();
                    string sql = "select FileDateTitlePage from ����������������� where id_�������� = " + this.Id�������� + "";
                    SqlCommand command = new SqlCommand(sql, connection);
                    //SqlDataReader reader = command.ExecuteReader();

                    // ������ ����� ������ �� ��.
                    byte[] fileByteArray = (byte[])command.ExecuteScalar();

                    string dir = @"d:\Recor";

                    string fileName = dir + @"\TempView.zip";

                    FileStream fileStream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                    BinaryWriter binWriter = new BinaryWriter(fileStream);
                    binWriter.Write(fileByteArray);
                    binWriter.Close();

                    // ������� �����.
                    System.Diagnostics.Process.Start(fileName);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //    // ������� ��� ����� �� �������.
            //    string ���������������� = this.pathFileServerTitlePage.Trim() + ".zip";

            //    try
            //    {

            //    // ���� � ��������.
            //    string ����������� = patchServerSave.Trim();

            //    // ������� ���� � ����� �� �������.
            //    string fileServer = ����������� + @"\�����������������\" + ����������������.Replace("/", "-");

            //    // ���� � ����� �� ��������� �����.
            //    string tempPath = Path.GetTempPath();

            //    // ������� ���� � ��� ����� ������� �� ����� ����� ����� ����������� �� ��������� ������� �� ������.
            //    string fileTo = tempPath + ����������������;

            //    // ��������� ����� �� ��������� �����.
            //    File.Copy(fileServer, fileTo, true);

            //    // ������� ���� � ����� �� ��������� ������� �� �������.
            //    string fileTemp = tempPath + @"\" + ����������������;

            //    // ������� �����.
            //    System.Diagnostics.Process.Start(fileTemp);
            //}
            //catch(Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
        }

        private void Form��������_FormClosing(object sender, FormClosingEventArgs e)
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
                MessageBox.Show(@"�������� ����� d:\Recor ");
            }

            

        }

        private void Form��������_FormClosed(object sender, FormClosedEventArgs e)
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
                MessageBox.Show(@"�������� ����� d:\Recor");
            }
        }

        private void chboxDsp_CheckedChanged(object sender, EventArgs e)
        {
           if (this.chboxDsp.Checked == true)
           {
               this.textBox�������������.Mask = "00-00-00-00-aaa";
               textBox�������������.Text += "���";
               this.FlagDsp = "True";
           }
           else
           {
               string stest = sNumStart;
               this.textBox�������������.Mask = "00-00-00-00";
               textBox�������������.Text = stest;
               this.FlagDsp = "False";
           }
        }

        private void textBox�������������_TextChanged(object sender, EventArgs e)
        {
            //if (textBox�������������.Text.Length == 11)
            //{
            //    this.chboxDsp.Enabled = true;
            //}
            //else
            //{
            //    this.chboxDsp.Enabled = true;
            //}
            sNumStart = this.textBox�������������.Text;
        }

        /// <summary>
        /// ��������� ������ �������.
        /// </summary>
        private void LoadNumberDepartments()
        {
            string query = "select ������������������ from ���������������������";

            foreach(DataRow row in DataTableSql.GetDataTableRows(query))
            {
                numbersDepartment.Add(row["������������������"].ToString().Trim(), row["������������������"].ToString().Trim());
            }
        }

        private void rb04_CheckedChanged(object sender, EventArgs e)
        {
             textBox�������������.Text = "04-";
        }

        private void rb02_CheckedChanged(object sender, EventArgs e)
        {
            textBox�������������.Text = "02-";
        }

        private void btnLastNumber_Click(object sender, EventArgs e)
        {
            #region ������ ���������� ���� �������
            /*
            string query = "select MAX(�������) from �������� " +
                           "where YEAR(����������) = "+ this.CurrentYear +" ";

            GetDataTable tab = new GetDataTable(query);
            string numLastDoc = tab.DataTable("SelectedYear").Rows[0][0].ToString();

            int lastNumberDocumnet = Convert.ToInt32(numLastDoc) + 1;

            numLastDoc = string.Empty;
            numLastDoc = lastNumberDocumnet.ToString();

            MessageBox.Show("��������� ����� ���������  - " + numLastDoc.ToString());

            this.lastNumberDoc = numLastDoc;

            flagLastNumberDoc = true;

            label1.Text = "����. ����� �\\� " + (this.lastNumberDoc);

            //this.flagNumberDoc.Checked = true;
             */

            #endregion

            bool flagEdit = �������������;

            // ���� ��������� ��� �� �������� � ��������� �������� ����������.
            bool flagInputCard = true;

            Form����������������� form��������� = new Form�����������������(this.Id��������, flagEdit, flagInputCard);

            // ������� ������ ����������������� ����� ��������������.
            form���������.List�����������������.Clear();

            // ��������� � ����� id ��������.
            form���������.Id�������� = this.Id��������;

            form���������.ShowDialog();

            if (form���������.DialogResult == DialogResult.OK)
            {
                // �.�. ��������� ����� ��������.
                if (flagEdit == true && flagInputCard == true)
                {
                    // ���������� ������ ������� ��� ���������� � ��������� �������, ��������� ��� �������� ������������ ������.
                    IQueryStringSQL queryInsert = new InsertQuery�����������������(form���������.List�����������������, this.Id��������);

                    // ��������� � �����: �������� ������ ������� �� ���������� ��������� � �������� ������������ ������.
                    this.QueryPersonDateForCardInput = queryInsert.Query();
                }
                else if (flagEdit == false && flagInputCard == true)
                {
                    // ���������� ������ ������� ��� ���������� ��������� �������.
                    IQueryStringSQL queryUpdate = new UpdateQuery�����������������(form���������.List�����������������, this.Id��������);

                    // ��������� � �����: �������� ������ ������� �� ���������� ��������� � �������� ������������ ������.
                    this.QueryPersonDateForCardInput = queryUpdate.Query();

                }
            }

            // ��������� ����� ��� ������ ��������� ��� �������� ������������ ������
        }

    }
}
