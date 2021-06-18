using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using RegKor.Classess;
namespace RegKor
{
	/// <summary>
	/// Summary description for Form2.
	/// </summary>
	public class Form���������� : System.Windows.Forms.Form
	{
		/// <summary>
		/// ��������� �����.
		/// 0 - ��������, 
		/// 1 - ����������, 
		/// 2 - ���������, 
		/// 3 - ��������.
		/// </summary>
		private int flag = 0;

		private System.Windows.Forms.Button btnAdd;
		private System.Windows.Forms.Button btnEdit;
		private System.Windows.Forms.Button btnDel;
		private System.Windows.Forms.Button btnClose;
		private System.Windows.Forms.ListBox List;
		private System.Windows.Forms.RichTextBox txt;
		private System.Windows.Forms.Button btnSave;
		private System.Windows.Forms.Button btnCancel;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        /// <summary>
        /// ����������� ��� �������������� � ���������� ������
        /// </summary>
        DS1TableAdapters.����������TableAdapter ����������TableAdapter;

        /// <summary>
        /// ������� - ��������� ����� ������
        /// </summary>
        private DS1 ds11;

		public Form����������()
		{
			InitializeComponent();
            ����������TableAdapter = new RegKor.DS1TableAdapters.����������TableAdapter();
			��������������������������();
		}

		private void ��������������������������()
		{
			ds11.����������.Clear();

            �����������������(ds11.����������);

            //����������TableAdapter.Fill(ds11.����������);

			// �������� � ��������� ������
			this.List.DataSource = ds11.����������.Select("", "������������������");
			this.List.ValueMember = "id_����������";
			this.List.DisplayMember = "������������������";
			this.List.SelectedValue = "id_����������";
			// �������� ������ ������� � ������:
			if(this.List.Items.Count > 0)
			{
				this.List.SetSelected(0, true);
			}
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form����������));
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.txt = new System.Windows.Forms.RichTextBox();
            this.List = new System.Windows.Forms.ListBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.ds11 = new RegKor.DS1();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).BeginInit();
            this.SuspendLayout();
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(268, 94);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(106, 30);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "��������";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(268, 130);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(106, 30);
            this.btnEdit.TabIndex = 3;
            this.btnEdit.Text = "��������";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(268, 168);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(106, 30);
            this.btnDel.TabIndex = 4;
            this.btnDel.Text = "�������";
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(420, 232);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(130, 30);
            this.btnClose.TabIndex = 5;
            this.btnClose.Text = "�������";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // txt
            // 
            this.txt.DetectUrls = false;
            this.txt.Dock = System.Windows.Forms.DockStyle.Top;
            this.txt.Enabled = false;
            this.txt.Location = new System.Drawing.Point(264, 0);
            this.txt.Multiline = false;
            this.txt.Name = "txt";
            this.txt.Size = new System.Drawing.Size(290, 20);
            this.txt.TabIndex = 6;
            this.txt.Text = "";
            this.txt.WordWrap = false;
            // 
            // List
            // 
            this.List.Dock = System.Windows.Forms.DockStyle.Left;
            this.List.Location = new System.Drawing.Point(0, 0);
            this.List.Name = "List";
            this.List.Size = new System.Drawing.Size(264, 264);
            this.List.TabIndex = 7;
            this.List.SelectedIndexChanged += new System.EventHandler(this.List_SelectedIndexChanged);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(288, 26);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(106, 30);
            this.btnSave.TabIndex = 8;
            this.btnSave.Text = "OK";
            this.btnSave.Visible = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(430, 26);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(106, 30);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "������";
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // ds11
            // 
            this.ds11.DataSetName = "DS1";
            this.ds11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // Form����������
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(554, 265);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.txt);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnEdit);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.List);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(560, 290);
            this.Name = "Form����������";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "���������� \"����������\"";
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// ��������� �������� ����������
		/// </summary>
		/// <param name="val"></param>
		private void LockConponent(bool val)
		{
			btnSave.Visible = val;
			btnCancel.Visible = val;
			btnAdd.Enabled = !val;
			btnDel.Enabled = false;
			btnEdit.Enabled = !val;
			btnClose.Enabled = !val;
			List.Enabled = !val;
			txt.Enabled = val;
		}
	
		/// <summary>
		/// ������� Click ������ �������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnClose_Click(object sender, System.EventArgs e)
		{
			Dispose(true);
		}

		/// <summary>
		/// ������� ��������� ��������� � ������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void List_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			txt.Text = Convert.ToString(List.Text);
		}

		/// <summary>
		/// ������� Click ������ ��������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnAdd_Click(object sender, System.EventArgs e)
		{
			flag = 1;
			LockConponent(true);
			txt.Text = "";
			txt.Focus();
		}


		/// <summary>
		/// ������� Click ������ ��������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			flag = 2;
			LockConponent(true);
		}

		/// <summary>
		/// ������� Click ������ �������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnDel_Click(object sender, System.EventArgs e)
		{
            flag = 3;
            LockConponent(true);
		}

		/// <summary>
		/// ������� Click ������ ���������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnSave_Click(object sender, System.EventArgs e)
		{
			if(txt.Text.Trim() == "")
			{
				flag = 0;
				LockConponent(false);
				return;
			}

			if(flag == 1)// ����������:
			{
			    DataRow[] row = ds11.����������.Select("������������������ = '" + txt.Text.Trim() + "'");
			    if(row.Length > 0)
			    {
				    MessageBox.Show(this, "���������� � ����� ��������� ��� ���� � �����������.", "���������� ����������", MessageBoxButtons.OK, MessageBoxIcon.Information);
				    flag = 0;
				    LockConponent(false);
				    return;
			    }
                ����������TableAdapter.Insert(txt.Text.Trim(), null, null);
                ��������������������������();
			}
			if(flag == 2)// ���������: 
			{
                DataRow[] rows = ds11.����������.Select("id_����������=" + (int)this.List.SelectedValue);
                rows[0]["������������������"] = txt.Text.Trim();
                ����������TableAdapter.Update(rows[0]);
                ��������������������������();
			}

            // ������� ����������.
            if (flag == 3)
            {
               string query = "update dbo.���������� " +
                               "set ������ = 'True' " +
                               "where LOWER(RTRIM(LTRIM(������������������))) =  '"+ this.txt.Text.Trim().ToLower() +"' ";

               ExecuteQuery execQuery = new ExecuteQuery(query);
               execQuery.Excecute();

                //DataRow[] rows = ds11.����������.Select("id_����������=" + (int)this.List.SelectedValue);
                //rows[0]["������������������"] = txt.Text.Trim();
                //����������TableAdapter.Update(rows[0]);
                ��������������������������();
            }
			flag = 0;
			LockConponent(false);
		}

		/// <summary>
		/// ������� Click ������ ������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			flag = 0;
			LockConponent(false);
			this.List_SelectedIndexChanged(null, null);
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
