using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace RegKor
{
	/// <summary>
	/// Summary description for Form2.
	/// </summary>
	public class Form������������� : System.Windows.Forms.Form
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
        private System.Windows.Forms.TextBox textBox������������;
        private System.Windows.Forms.TextBox textBox������������������;
        private System.Windows.Forms.TextBox textBox��������������������;
		private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private IContainer components;

        /// <summary>
        /// ����������� ��� �������������� � ���������� ������
        /// </summary>
        DS1TableAdapters.���������������������TableAdapter �������������TableAdapter;
        DS1TableAdapters.����������TableAdapter ������������TableAdapter;
        private Panel panel1;
        private Label label1;
        private Panel panel2;
        private Label label2;
        private Label label3;
        private Button button�����������������;
        private Label label4;
        private ToolTip toolTip1;
        private Label label������������;
        private object id_������������;
        private DS1 ds11;
        private RichTextBox txt���������������������;


        public Form������������� ( DS1 dataset )
		{
			InitializeComponent();
            this.ds11 = dataset;

            �������������TableAdapter = new RegKor.DS1TableAdapters.���������������������TableAdapter( );
            ������������TableAdapter = new RegKor.DS1TableAdapters.����������TableAdapter( );

			��������������������������();
		}

		private void ��������������������������()
		{
            ds11.���������������������.Clear( );
            �������������TableAdapter.Fill( ds11.��������������������� );

			// �������� � ��������� ������
            this.List.DataSource = ds11.���������������������.Select( "������=False", "������������������, ��������������������" );
            this.List.ValueMember = "id_�������������";
            this.List.DisplayMember = "���������������������";
            this.List.SelectedValue = "id_�������������";
			// �������� ������ ������� � ������:
			if(this.List.Items.Count > 0)
			{
				this.List.SetSelected(0, true);
			}
		}


		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form�������������));
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.List = new System.Windows.Forms.ListBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txt��������������������� = new System.Windows.Forms.RichTextBox();
            this.label������������ = new System.Windows.Forms.Label();
            this.button����������������� = new System.Windows.Forms.Button();
            this.textBox������������ = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBox�������������������� = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox������������������ = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.ds11 = new RegKor.DS1();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).BeginInit();
            this.SuspendLayout();
            // 
            // btnAdd
            // 
            this.btnAdd.Location = new System.Drawing.Point(3, 3);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(106, 30);
            this.btnAdd.TabIndex = 1;
            this.btnAdd.Text = "��������";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(3, 39);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(106, 30);
            this.btnEdit.TabIndex = 2;
            this.btnEdit.Text = "��������";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnDel
            // 
            this.btnDel.Enabled = false;
            this.btnDel.Location = new System.Drawing.Point(3, 75);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(106, 30);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "�������";
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnClose.Location = new System.Drawing.Point(316, 115);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(130, 30);
            this.btnClose.TabIndex = 10;
            this.btnClose.Text = "�������";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // List
            // 
            this.List.Dock = System.Windows.Forms.DockStyle.Left;
            this.List.Location = new System.Drawing.Point(0, 0);
            this.List.Name = "List";
            this.List.Size = new System.Drawing.Size(289, 407);
            this.List.TabIndex = 0;
            this.List.SelectedIndexChanged += new System.EventHandler(this.List_SelectedIndexChanged);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(51, 221);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(106, 30);
            this.btnSave.TabIndex = 8;
            this.btnSave.Text = "OK";
            this.btnSave.Visible = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(255, 221);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(106, 30);
            this.btnCancel.TabIndex = 9;
            this.btnCancel.Text = "������";
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txt���������������������);
            this.panel1.Controls.Add(this.label������������);
            this.panel1.Controls.Add(this.button�����������������);
            this.panel1.Controls.Add(this.textBox������������);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.textBox��������������������);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.textBox������������������);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.btnSave);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Enabled = false;
            this.panel1.Location = new System.Drawing.Point(289, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(449, 258);
            this.panel1.TabIndex = 10;
            // 
            // txt���������������������
            // 
            this.txt���������������������.Location = new System.Drawing.Point(117, 23);
            this.txt���������������������.Name = "txt���������������������";
            this.txt���������������������.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.txt���������������������.Size = new System.Drawing.Size(323, 85);
            this.txt���������������������.TabIndex = 4;
            this.txt���������������������.Text = "";
            // 
            // label������������
            // 
            this.label������������.AutoSize = true;
            this.label������������.Location = new System.Drawing.Point(266, 159);
            this.label������������.Name = "label������������";
            this.label������������.Size = new System.Drawing.Size(0, 13);
            this.label������������.TabIndex = 18;
            // 
            // button�����������������
            // 
            this.button�����������������.Image = global::RegKor.Properties.Resources.add;
            this.button�����������������.Location = new System.Drawing.Point(416, 114);
            this.button�����������������.Name = "button�����������������";
            this.button�����������������.Size = new System.Drawing.Size(23, 23);
            this.button�����������������.TabIndex = 5;
            this.toolTip1.SetToolTip(this.button�����������������, "������� ������ �����������");
            this.button�����������������.UseVisualStyleBackColor = true;
            this.button�����������������.Click += new System.EventHandler(this.button�����������������_Click);
            // 
            // textBox������������
            // 
            this.textBox������������.BackColor = System.Drawing.SystemColors.Control;
            this.textBox������������.Location = new System.Drawing.Point(117, 116);
            this.textBox������������.Name = "textBox������������";
            this.textBox������������.ReadOnly = true;
            this.textBox������������.Size = new System.Drawing.Size(294, 20);
            this.textBox������������.TabIndex = 16;
            this.textBox������������.TabStop = false;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(16, 116);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(97, 17);
            this.label4.TabIndex = 15;
            this.label4.Text = "������������";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBox��������������������
            // 
            this.textBox��������������������.Location = new System.Drawing.Point(117, 179);
            this.textBox��������������������.MaxLength = 1;
            this.textBox��������������������.Name = "textBox��������������������";
            this.textBox��������������������.Size = new System.Drawing.Size(40, 20);
            this.textBox��������������������.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(1, 179);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(110, 17);
            this.label3.TabIndex = 13;
            this.label3.Text = "���������� ������";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBox������������������
            // 
            this.textBox������������������.Location = new System.Drawing.Point(117, 148);
            this.textBox������������������.MaxLength = 2;
            this.textBox������������������.Name = "textBox������������������";
            this.textBox������������������.Size = new System.Drawing.Size(61, 20);
            this.textBox������������������.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(11, 148);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 17);
            this.label2.TabIndex = 11;
            this.label2.Text = "������� ���";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(16, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 37);
            this.label1.TabIndex = 10;
            this.label1.Text = "�������� �������������";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnAdd);
            this.panel2.Controls.Add(this.btnEdit);
            this.panel2.Controls.Add(this.btnClose);
            this.panel2.Controls.Add(this.btnDel);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(289, 260);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(449, 148);
            this.panel2.TabIndex = 11;
            // 
            // ds11
            // 
            this.ds11.DataSetName = "DS1";
            this.ds11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // Form�������������
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.CancelButton = this.btnClose;
            this.ClientSize = new System.Drawing.Size(738, 408);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.List);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(560, 290);
            this.Name = "Form�������������";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "���������� \"�������������\"";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.ResumeLayout(false);

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
			panel1.Enabled = val;
		}
	
		/// <summary>
		/// ������� Click ������ �������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnClose_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		/// <summary>
		/// ������� ��������� ��������� � ������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void List_SelectedIndexChanged(object sender, System.EventArgs e)
		{
            try
            {
                DataRow [ ] row = ds11.���������������������.Select( "id_�������������=" + List.SelectedValue );
                txt���������������������.Text = Convert.ToString( row [0] ["���������������������"] );
                textBox������������������.Text = Convert.ToString( row [0] ["������������������"] );
                textBox��������������������.Text = Convert.ToString( row [0] ["��������������������"] );
                DataRow [ ] row2 = ds11.����������.Select( "id_����������=" + row [0] ["id_�������������������������"] );
                if ( row2.Length > 0 )
                {
                    id_������������ = Convert.ToInt32( row2 [0] ["id_����������"] );
                    textBox������������.Text = Convert.ToString( row2 [0] ["������������������"] );
                }
                else
                {
                    id_������������ = null;
                    textBox������������.Text = "";
                }
            }
            catch
            {
            }
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
            txt���������������������.Text = "";
            textBox������������.Text = "";
            textBox��������������������.Text = "";
            textBox������������������.Text = "";
			txt���������������������.Focus();
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
            if ( txt���������������������.Text.Trim( ) == "" || textBox������������.Text.Trim( ) == "" || textBox������������������.Text.Trim( ) == "")
			{
                if ( flag == 1 )
                {
                    txt���������������������.Text = "";
                    textBox������������.Text = "";
                    textBox��������������������.Text = "";
                    textBox������������������.Text = "";
                }
                if ( flag == 2 )
                {
                    try
                    {
                        DataRow [ ] row = ds11.���������������������.Select( "id_�������������=" + List.SelectedValue );
                        txt���������������������.Text = Convert.ToString( row [0] ["���������������������"] );
                        textBox������������������.Text = Convert.ToString( row [0] ["������������������"] );
                        textBox��������������������.Text = Convert.ToString( row [0] ["��������������������"] );
                        DataRow [ ] row2 = ds11.����������.Select( "id_����������=" + row [0] ["id_�������������������������"] );
                        if ( row2.Length > 0 )
                        {
                            id_������������ = Convert.ToInt32( row2 [0] ["id_����������"] );
                            textBox������������.Text = Convert.ToString( row2 [0] ["������������������"] );
                        }
                        else
                        {
                            id_������������ = null;
                            textBox������������.Text = "";
                        }
                    }
                    catch
                    {
                    }
                }
				flag = 0;
				LockConponent(false);
				return;
			}

			if(flag == 2)// ���������: 
			{
                DataRow [ ] rows = ds11.���������������������.Select( "id_�������������=" + ( int ) this.List.SelectedValue );

                rows [0] ["���������������������"] = txt���������������������.Text.Trim( );
                rows [0] ["id_�������������������������"] = id_������������;
                rows [0] ["������������������"] = textBox������������������.Text.Trim( );
                rows [0] ["��������������������"] = textBox��������������������.Text.Trim( );

                �������������TableAdapter.Update( rows [0] );

                ��������������������������( );
			}

			if(flag == 1)// ����������:
			{
                DataRow [ ] row = ds11.���������������������.Select( "��������������������� = '" + txt���������������������.Text.Trim( ) + "'");
                if ( row.Length > 0 )
                {
                    MessageBox.Show( this, "������������� � ����� ��������� ��� ���� � �����������.", "���������� �������������", MessageBoxButtons.OK, MessageBoxIcon.Information );
                    flag = 0;
                    LockConponent( false );
                    return;
                }

                �������������TableAdapter.Insert( txt���������������������.Text.Trim( ), ( int ) id_������������, textBox������������������.Text.Trim( ), textBox��������������������.Text.Trim( ), false );
                
                ��������������������������( );
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
            if ( flag == 1 )
            {
                txt���������������������.Text = "";
                textBox������������.Text = "";
                textBox��������������������.Text = "";
                textBox������������������.Text = "";
            }
            if ( flag == 2 )
            {
                try
                {
                    DataRow [ ] row = ds11.���������������������.Select( "id_�������������=" + List.SelectedValue );
                    txt���������������������.Text = Convert.ToString( row [0] ["���������������������"] );
                    textBox������������������.Text = Convert.ToString( row [0] ["������������������"] );
                    textBox��������������������.Text = Convert.ToString( row [0] ["��������������������"] );
                    DataRow [ ] row2 = ds11.����������.Select( "id_����������=" + row [0] ["id_�������������������������"] );
                    if ( row2.Length > 0 )
                    {
                        id_������������ = Convert.ToInt32( row2 [0] ["id_����������"] );
                        textBox������������.Text = Convert.ToString( row2 [0] ["������������������"] );
                    }
                    else
                    {
                        id_������������ = null;
                        textBox������������.Text = "";
                    }
                }
                catch
                {
                }
            }
			flag = 0;
			LockConponent(false);
			this.List_SelectedIndexChanged(null, null);
		}

        private void button�����������������_Click ( object sender, EventArgs e )
        {
            Form����������������� form = new Form�����������������( );
            DialogResult result = form.ShowDialog( this );

            if ( result == DialogResult.OK )
            {
                //id_������������ = (int)form.������������;
                DataRow [ ] row2 = ds11.����������.Select( "id_����������=" + form.������������ );
                if ( row2.Length > 0 )
                {
                    id_������������ = Convert.ToInt32( row2 [0] ["id_����������"] );
                    textBox������������.Text = Convert.ToString( row2 [0] ["������������������"] );
                }
                //textBox������������.Text = 
                //����������������������� ["id_������������������"] = form.��������������������;
                //�������������������� = ( int ) form.��������������������;
                //textBox���������������.Text = ���������������( form.�������������������� );
            }
        }

	}
}
