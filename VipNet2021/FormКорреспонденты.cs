using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using RegKor.Classess;

namespace RegKor
{
	/// <summary>
	/// Summary description for Form2.
	/// </summary>
	public class Form�������������� : System.Windows.Forms.Form
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
        DS1TableAdapters.��������������TableAdapter ��������������TableAdapter;
        private CheckBox checkBox1;
        private Button btnEditClose;
        private GroupBox groupBox1;
        private Button btnOpen;
        private CheckBox flag��������������;

        /// <summary>
        /// ������� - ��������� ����� ������
        /// </summary>
        private DS1 ds11;

        public Form�������������� ( string ������������������ )
        {
            InitializeComponent( );
            ��������������TableAdapter = new RegKor.DS1TableAdapters.��������������TableAdapter( );
            ��������������������������( );
            flag = 1;
            LockConponent( true );
            txt.Text = ������������������;
        }

		public Form��������������()
		{
            InitializeComponent();
            ��������������TableAdapter = new RegKor.DS1TableAdapters.��������������TableAdapter();
			��������������������������();
		}

		private void ��������������������������()
		{
            //�������� DataSet
            ds11.��������������.Clear();

            //=====================

            ////��������� ���� � ���������������� ����� ���� ������������ = "������" ����� 2012 ��� ��� ������
            ////���� ������������ = "�����" ����� 2011 ��� ��� ������
            //Classess.������������������� ��� = new RegKor.Classess.�������������������();
            //bool flag = ���.�������������������������();

            ////���� 2011 ��� ��� ������ ����� ds11 ���� DataSet ��������� ��� ������
            //if (flag == false)
            //{
            //    //��������� DataSet ds11 �������
            //    ��������������TableAdapter.Fill(ds11.��������������);
            //}

            ////���� 2012 ��� ��� ������ ����� ��������� ds11 ���� DataSet ������� ��� � �������
            ////�������������� ���� ����� = NULL  

            //if (flag == true && this.flag��������������.Checked == false)
            //{
            //    Classess.�������������� �������������� = new RegKor.Classess.��������������();
            //    DataSet ds�������������� = ��������������.�����������������������_DataSet();

            //    //foreach (DataRow row�������������� in ds��������������.Tables[0].Rows)
            //    //{
            //    //    DataRow row1 = ds11.��������������.NewRow();
            //    //    row1[0] = row��������������[0];
            //    //    row1[1] = row��������������[1];
            //    //    row1[2] = row��������������[2];
            //    //    ds11.��������������.Rows.Add(row1);
            //    //}

            //}

            //////��������� ������� ������
            //if (flag == true && this.flag��������������.Checked == true)
            //{

            //    �������������� �������������� = new ��������������();
            //    DataSet ds��������������������� = ��������������.���������������();

            //    //foreach (DataRow row�������������� in ds���������������������.Tables[0].Rows)
            //    //{
            //    //    DataRow row1 = ds11.��������������.NewRow();
            //    //    row1[0] = row��������������[0];
            //    //    row1[1] = row��������������[1];
            //    //    row1[2] = row��������������[2];
            //    //    ds11.��������������.Rows.Add(row1);
            //    //}

            //}


            //=======================

            ////��������� DataSet ds11 �������
            ��������������TableAdapter.Fill(ds11.��������������); //- ������ ������� �����

			// �������� � ��������� ������
			this.List.DataSource = ds11.��������������.Select("", "����������������������");
    		this.List.ValueMember = "id_��������������";
			this.List.DisplayMember = "����������������������";
			this.List.SelectedValue = "id_��������������";

            int iCount =   ds11.��������������.Rows.Count;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form��������������));
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.txt = new System.Windows.Forms.RichTextBox();
            this.List = new System.Windows.Forms.ListBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.ds11 = new RegKor.DS1();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.btnEditClose = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnOpen = new System.Windows.Forms.Button();
            this.flag�������������� = new System.Windows.Forms.CheckBox();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(268, 93);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(106, 30);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "��������";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(268, 129);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(106, 30);
            this.btnEdit.TabIndex = 3;
            this.btnEdit.Text = "��������";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnDel
            // 
            this.btnDel.Enabled = false;
            this.btnDel.Location = new System.Drawing.Point(268, 167);
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
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(271, 69);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(198, 17);
            this.checkBox1.TabIndex = 10;
            this.checkBox1.Text = "��� ���������������� 2012 ����";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Visible = false;
            this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
            // 
            // btnEditClose
            // 
            this.btnEditClose.Enabled = false;
            this.btnEditClose.Location = new System.Drawing.Point(268, 204);
            this.btnEditClose.Name = "btnEditClose";
            this.btnEditClose.Size = new System.Drawing.Size(106, 30);
            this.btnEditClose.TabIndex = 11;
            this.btnEditClose.Text = "������";
            this.btnEditClose.UseVisualStyleBackColor = true;
            this.btnEditClose.Visible = false;
            this.btnEditClose.Click += new System.EventHandler(this.btnEditClose_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnOpen);
            this.groupBox1.Controls.Add(this.flag��������������);
            this.groupBox1.Location = new System.Drawing.Point(381, 93);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(169, 104);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "��������������";
            this.groupBox1.Visible = false;
            // 
            // btnOpen
            // 
            this.btnOpen.Enabled = false;
            this.btnOpen.Location = new System.Drawing.Point(7, 60);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(106, 30);
            this.btnOpen.TabIndex = 1;
            this.btnOpen.Text = "�������";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // flag��������������
            // 
            this.flag��������������.AutoSize = true;
            this.flag��������������.Location = new System.Drawing.Point(6, 36);
            this.flag��������������.Name = "flag��������������";
            this.flag��������������.Size = new System.Drawing.Size(118, 17);
            this.flag��������������.TabIndex = 0;
            this.flag��������������.Text = "������� �������";
            this.flag��������������.UseVisualStyleBackColor = true;
            this.flag��������������.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // Form��������������
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(554, 265);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnEditClose);
            this.Controls.Add(this.checkBox1);
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
            this.Name = "Form��������������";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "���������� \"��������\"";
            this.Load += new System.EventHandler(this.Form��������������_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

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
			this.Close();
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
//			flag = 3;
//			LockConponent(true);
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
			    DataRow[] row = ds11.��������������.Select("���������������������� = '" + txt.Text.Trim() + "'");
			    if(row.Length > 0)
			    {
				    MessageBox.Show(this, "������������� � ����� ��������� ��� ���� � �����������.", "���������� ��������������", MessageBoxButtons.OK, MessageBoxIcon.Information);
				    flag = 0;
				    LockConponent(false);
				    return;
			    }
                ��������������TableAdapter.Insert(txt.Text.Trim(), null);
                ��������������������������();
			}
			if(flag == 2)// ���������: 
			{
				DataRow[] rows = ds11.��������������.Select("id_��������������=" + (int)this.List.SelectedValue);
                rows[0]["����������������������"] = txt.Text.Trim();
                ��������������TableAdapter.Update(rows[0]);
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

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (this.checkBox1.Checked == true)
            {
                this.btnEditClose.Enabled = true;
            }
            else
            {
                this.btnEditClose.Enabled = false;
            }
        }

        private void btnEditClose_Click(object sender, EventArgs e)
        {
            �������������� �������������� = new ��������������();
            ��������������.������(this.txt.Text);

            ��������������������������();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (this.flag��������������.Checked == true)
            {
                btnAdd.Enabled = false;
                btnEdit.Enabled = false;
                btnEditClose.Enabled = false;
                btnOpen.Enabled = true;
            }
            else
            {
                btnAdd.Enabled = true;
                btnEdit.Enabled = true;
                btnEditClose.Enabled = true;
            }
            ��������������������������();


            //Classess.�������������� �������������� = new RegKor.Classess.��������������();
            //DataSet ds�������������� = ��������������.�����������������������_DataSet();

            //foreach (DataRow row�������������� in ds��������������.Tables[0].Rows)
            //{
            //    DataRow row1 = ds11.��������������.NewRow();
            //    row1[0] = row��������������[0];
            //    row1[1] = row��������������[1];
            //    ds11.��������������.Rows.Add(row1);
            //}
        }

        private void Form��������������_Load(object sender, EventArgs e)
        {
            RegKor.Classess.������������������� ������������ = new �������������������();
            bool flag = ������������.�������������������������();

            if (flag == true)
            {
                if (this.checkBox1.Checked == true)
                {
                    this.btnEditClose.Enabled = true;
                }
                else
                {
                    this.btnEditClose.Enabled = false;
                }
                this.groupBox1.Visible = true;
                this.flag��������������.Visible = true;
                this.btnOpen.Visible = true;
                this.btnEditClose.Visible = true;
                this.checkBox1.Visible = true;
            }
            else
            {
                this.groupBox1.Visible = false;
                this.flag��������������.Visible = false;
                this.btnOpen.Visible = false;
                this.btnEditClose.Visible = false;
                this.checkBox1.Visible = false;
            }
            
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            �������������� �������������� = new ��������������();
            ��������������.�������(this.txt.Text);
            this.flag��������������.Checked = false;
            ��������������������������();
            this.btnOpen.Enabled = false;
        }

        

	}
}
