using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Collections.Generic;
using System.Text;
using RegKor.Classess;

namespace RegKor
{
    /// <summary>
    /// Summary description for FormResolution.
    /// </summary>
    public class Form��������� : System.Windows.Forms.Form
    {

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox List1;
        private System.Windows.Forms.ListBox List2;
        private System.Windows.Forms.Button btnAddAll;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnDelAll;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.ComponentModel.IContainer components;

        /// <summary>
        /// ������� 
        /// </summary>
        private DS1 ds11;

        /// <summary>
        /// ���������� ������ � ������������
        /// </summary>
        public string ��������������� = null;

        /// <summary>
        /// ������ ���������� �����������
        /// </summary>
        string[] Arr;

        /// <summary>
        /// ����������� ��� �������������� � ���������� ������
        /// </summary>
        DS1TableAdapters.����������TableAdapter ����������TableAdapter;

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


        public Form���������()
        {
            InitializeComponent();

            ����������TableAdapter = new RegKor.DS1TableAdapters.����������TableAdapter();

            ��������������������������();

        }

       

        private void ��������������������������()
        {

            ds11.����������.Clear();
            //����������TableAdapter.Fill(ds11.����������);

            �����������������(ds11.����������);


            //----------------------------------������ ���----------------------------
            //oleDbConnectionAccess.Open();
            //oleDbDataAdapter����������.Fill(this.ds1);
            //oleDbConnectionAccess.Close();
            //------------------------------------------------------------------------


            // �������� ��� ������ �� ������� ����������
            DataRow[] dr = this.ds11.����������.Select("", "������������������");
            // ������� ��������� ������ ������ ������ ���������� ������� � ������� ����������
            Arr = new String[dr.Length];
            // ��������� ��������� ������ ���������� ���������������
            for (int i = 0; i < dr.Length; i++)
            {
                DataRow temp = dr[i];

                //if (temp["������������������"].ToString().ToLower() != "��������� �.�.".Trim().ToLower())
                //{
                if (temp["������������������"].ToString().ToLower() != "�������� �.�.".Trim().ToLower())
                    {
                        Arr[i] = (string)temp["������������������"];
                        List1.Items.Add(Arr[i]);// � ����� ��������� ������ ListBox
                    }
                    else
                    {
                        Arr[i] = (string)temp["������������������"];
                        List2.Items.Add(Arr[i]);// � ����� ��������� ������ ListBox
                    }
                //}
            }
        }

        /// <summary>
        /// ����������� ������������ ������� � ��������� ����� �� ������
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

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form���������));
            this.List1 = new System.Windows.Forms.ListBox();
            this.List2 = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnAddAll = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnDelAll = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.ds11 = new RegKor.DS1();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).BeginInit();
            this.SuspendLayout();
            // 
            // List1
            // 
            this.List1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.List1.Location = new System.Drawing.Point(0, 22);
            this.List1.Name = "List1";
            this.List1.Size = new System.Drawing.Size(256, 303);
            this.List1.Sorted = true;
            this.List1.TabIndex = 0;
            this.toolTip1.SetToolTip(this.List1, "������ ��������� �����������");
            // 
            // List2
            // 
            this.List2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.List2.Location = new System.Drawing.Point(0, 22);
            this.List2.Name = "List2";
            this.List2.Size = new System.Drawing.Size(256, 303);
            this.List2.TabIndex = 1;
            this.toolTip1.SetToolTip(this.List2, "������ ��������� �����������");
            // 
            // label1
            // 
            this.label1.Dock = System.Windows.Forms.DockStyle.Top;
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(256, 22);
            this.label1.TabIndex = 2;
            this.label1.Text = "������ ��������� �����������";
            this.label1.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // label2
            // 
            this.label2.Dock = System.Windows.Forms.DockStyle.Top;
            this.label2.Location = new System.Drawing.Point(0, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(256, 22);
            this.label2.TabIndex = 3;
            this.label2.Text = "���� ��������� ��������";
            this.label2.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // btnAddAll
            // 
            this.btnAddAll.Location = new System.Drawing.Point(261, 94);
            this.btnAddAll.Name = "btnAddAll";
            this.btnAddAll.Size = new System.Drawing.Size(62, 34);
            this.btnAddAll.TabIndex = 4;
            this.btnAddAll.Text = ">>";
            this.toolTip1.SetToolTip(this.btnAddAll, "�������� ���� ����������� �� ������");
            this.btnAddAll.Click += new System.EventHandler(this.btnAddAll_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(261, 144);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(62, 34);
            this.btnAdd.TabIndex = 5;
            this.btnAdd.Text = ">";
            this.toolTip1.SetToolTip(this.btnAdd, "�������� ���������� ����������");
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(261, 196);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(62, 34);
            this.btnDel.TabIndex = 6;
            this.btnDel.Text = "<";
            this.toolTip1.SetToolTip(this.btnDel, "������� ���������� ����������");
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnDelAll
            // 
            this.btnDelAll.Location = new System.Drawing.Point(261, 246);
            this.btnDelAll.Name = "btnDelAll";
            this.btnDelAll.Size = new System.Drawing.Size(62, 34);
            this.btnDelAll.TabIndex = 7;
            this.btnDelAll.Text = "<<";
            this.toolTip1.SetToolTip(this.btnDelAll, "������� ���� ����������� �����������");
            this.btnDelAll.Click += new System.EventHandler(this.btnDelAll_Click);
            // 
            // btnSave
            // 
            this.btnSave.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnSave.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnSave.Location = new System.Drawing.Point(70, 8);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(183, 32);
            this.btnSave.TabIndex = 94;
            this.btnSave.Text = "���������";
            this.toolTip1.SetToolTip(this.btnSave, "��������� ��� ��������� � ������� ����");
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnCancel.Location = new System.Drawing.Point(332, 8);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(183, 32);
            this.btnCancel.TabIndex = 95;
            this.btnCancel.Text = "������";
            this.toolTip1.SetToolTip(this.btnCancel, "������� ���� ��� ���������� ���������");
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnSave);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 327);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(584, 48);
            this.panel1.TabIndex = 96;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.List1);
            this.panel2.Controls.Add(this.label1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(256, 327);
            this.panel2.TabIndex = 97;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.List2);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel3.Location = new System.Drawing.Point(328, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(256, 327);
            this.panel3.TabIndex = 98;
            // 
            // ds11
            // 
            this.ds11.DataSetName = "DS1";
            this.ds11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // Form���������
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(584, 375);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnDelAll);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.btnAddAll);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form���������";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "���������";
            this.Load += new System.EventHandler(this.FormResolution_Load);
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.ResumeLayout(false);

        }
        #endregion

        private void FormResolution_Load(object sender, System.EventArgs e)
        {

        }

        /// <summary>
        /// ������� Click ������ ������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnCancel_Click(object sender, System.EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// ������� Click ������ �������� ���� �����������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAddAll_Click(object sender, System.EventArgs e)
        {
            if (List1.Items.Count > 0)// ���� �������� ������ �� ������
            {
                List2.Items.Clear(); // ������� ����������� ������
                // ��������� � ����������� ������ ��� ������ �� ���������
                for (int i = 0; i < Arr.Length; i++)
                {
                    List2.Items.Add(Arr[i]);
                }
                // ������� �������� ������
                List1.Items.Clear();
                empty_List1(true);
                empty_List2(false);
            }
        }

        /// <summary>
        /// ������� Click ������ ������� ���� �����������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDelAll_Click(object sender, System.EventArgs e)
        {
            if (List2.Items.Count > 0)// ���� ����������� ������ �� ������
            {
                List1.Items.Clear();// ������� �������� ������
                // ��������� � �������� ������ ��� ������ �� ������������
                for (int i = 0; i < Arr.Length; i++)
                {
                    List1.Items.Add(Arr[i]);
                }
                // ������� ����������� ������
                List2.Items.Clear();
                empty_List2(true);
                empty_List1(false);
            }
        }

        /// <summary>
        /// ������� Click ������ �������� ����������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnAdd_Click(object sender, System.EventArgs e)
        {
            if (List1.Items.Count > 0)// ���� �������� ������ �� ������
            {
                if (List1.SelectedItem == null)// ���� � �������� ������ �� ������� �� ���� �������
                {
                    List1.SetSelected(0, true);// �������� ������ ������� � �������� ������
                }
                // �������� � ������ ���������� ����� �� ��������� ������:
                string str = (string)List1.SelectedItem;
                // ������� ���������� ������� � �������� ������:
                List1.Items.Remove(List1.SelectedItem);
                // � ����������� ������ ��������� ����������� ������:
                List2.Items.Add(str);
                empty_List2(false);
                if (List1.Items.Count == 0)
                {
                    empty_List1(true);
                }
                else
                {
                    empty_List1(false);
                }
            }
            else
            {
                empty_List1(true);
            }
        }

        /// <summary>
        /// ������� Click ������ ������� ����������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnDel_Click(object sender, System.EventArgs e)
        {
            if (List2.Items.Count > 0)// ���� ����������� ������ �� ������
            {
                if (List2.SelectedItem == null)// ���� � ����������� ������ �� ������� �� ���� �������
                {
                    List2.SetSelected(0, true);// �������� ������ ������� � ����������� ������
                }
                // �������� � ������ ���������� ����� �� ������������ ������:
                string str = (string)List2.SelectedItem;
                // ������� ���������� ������� � �������� ������:
                List2.Items.Remove(List2.SelectedItem);
                // � ����������� ������ ��������� ����������� ������:
                List1.Items.Add(str);
                empty_List1(false);
                if (List2.Items.Count == 0)
                {
                    empty_List2(true);
                }
                else
                {
                    empty_List2(false);
                }
            }
            else
            {
                empty_List2(true);
            }
        }

        /// <summary>
        /// ��������-��������� ������ ���������� ����������� � ��������
        /// </summary>
        /// <param name="val"></param>
        private void empty_List1(bool val)
        {
            btnAdd.Enabled = !val;
            btnAddAll.Enabled = !val;
        }

        /// <summary>
        /// ��������-��������� ������ �������� ����������� �� ��������
        /// </summary>
        /// <param name="val"></param>
        private void empty_List2(bool val)
        {
            btnDel.Enabled = !val;
            btnDelAll.Enabled = !val;
        }

        /// <summary>
        /// ������� Click ������ ���������
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, System.EventArgs e)
        {
            listPerson = new List<PersonRecepient>();

            // ������ ��� �������� ��� ��������� ����������.
            StringBuilder build = new StringBuilder();

            if (List2.Items.Count < 1)
            {
                return;
            }

            string str = "";
            for (int i = 0; i < List2.Items.Count; i++)
            {
                str += (string)List2.Items[i];

                build.Append( "'"+ (string)List2.Items[i] +"' ");
                build.Append(", ");

                if ((i + 1) != List2.Items.Count)
                {
                    str += ", ";
                }
            }
            if (str != "" || str != null)
            {
                ��������������� = str;
            }

            // ������� �� �� ��� ����������.
            int aaa = build.ToString().Trim().Length;
            int length = build.ToString().Trim().Length - 1;
            string stringPerson = build.ToString().Trim().Remove(length, 1);

            string query = "SELECT id_����������, ������������������  FROM [����������] " +
                           "where ������������������ in ("+ stringPerson +")";

            GetDataTable getTable = new GetDataTable(query);
            DataTable tabPerson = getTable.DataTable("�����������������");

            foreach (DataRow row in tabPerson.Rows)
            {
                PersonRecepient person = new PersonRecepient();
                person.ID = Convert.ToInt32(row["id_����������"]);
                person.Famili = row["������������������"].ToString().Trim();

                listPerson.Add(person);
            }

            this.ListPerson = listPerson;

            this.Close();
        }

        //��������� ������ �����������
        public void �����������������(DS1.����������DataTable dt)
        {
            string query = "SELECT [id_����������] " +
                           ",[������������������] " +
                           ",[��������������] " +
                           ",[������] " +
                           "FROM [����������] " +
                           "where [������] is null";
            SqlConnection con = new SqlConnection();
            con.ConnectionString = ConfigurationSettings.AppSettings["���������������������"].ToString();
            SqlCommand com = new SqlCommand(query, con);

            con.Open();
            SqlDataReader read = com.ExecuteReader();

            while (read.Read())
            {
                DataRow row = dt.NewRow();
                row["id_����������"] = read["id_����������"];
                row["������������������"] = read["������������������"];
                row["��������������"] = read["��������������"];
                row["������"] = read["������"];
                dt.Rows.Add(row);
            }

            //return dt;
        }
    }
}
