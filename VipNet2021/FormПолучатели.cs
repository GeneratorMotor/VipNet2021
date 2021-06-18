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
	public class FormПолучатели : System.Windows.Forms.Form
	{
		/// <summary>
		/// Состояние формы.
		/// 0 - ожидание, 
		/// 1 - добавление, 
		/// 2 - изменение, 
		/// 3 - удаление.
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
        /// Датаадаптер для взаимодействия с источником данных
        /// </summary>
        DS1TableAdapters.ПолучателиTableAdapter получателиTableAdapter;

        /// <summary>
        /// Датасет - локальная копия данных
        /// </summary>
        private DS1 ds11;

		public FormПолучатели()
		{
			InitializeComponent();
            получателиTableAdapter = new RegKor.DS1TableAdapters.ПолучателиTableAdapter();
			ПодключитьсяПолучитьДанные();
		}

		private void ПодключитьсяПолучитьДанные()
		{
			ds11.Получатели.Clear();

            ПолучателиАдаптер(ds11.Получатели);

            //получателиTableAdapter.Fill(ds11.Получатели);

			// привязка к источнику данных
			this.List.DataSource = ds11.Получатели.Select("", "ОписаниеПолучателя");
			this.List.ValueMember = "id_получателя";
			this.List.DisplayMember = "ОписаниеПолучателя";
			this.List.SelectedValue = "id_получателя";
			// выделяем первый элемент в списке:
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormПолучатели));
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
            this.btnAdd.Text = "Добавить";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(268, 130);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(106, 30);
            this.btnEdit.TabIndex = 3;
            this.btnEdit.Text = "Изменить";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnDel
            // 
            this.btnDel.Location = new System.Drawing.Point(268, 168);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(106, 30);
            this.btnDel.TabIndex = 4;
            this.btnDel.Text = "Удалить";
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
            this.btnClose.Text = "Закрыть";
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
            this.btnCancel.Text = "Отмена";
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // ds11
            // 
            this.ds11.DataSetName = "DS1";
            this.ds11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // FormПолучатели
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
            this.Name = "FormПолучатели";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Справочник \"Сотрудники\"";
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Блокирует элементы управления
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
		/// Событие Click кнопки ЗАКРЫТЬ
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnClose_Click(object sender, System.EventArgs e)
		{
			Dispose(true);
		}

		/// <summary>
		/// Событие изменения выделения в списке
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void List_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			txt.Text = Convert.ToString(List.Text);
		}

		/// <summary>
		/// Событие Click кнопки ДОБАВИТЬ
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
		/// Событие Click кнопки ИЗМЕНИТЬ
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnEdit_Click(object sender, System.EventArgs e)
		{
			flag = 2;
			LockConponent(true);
		}

		/// <summary>
		/// Событие Click кнопки УДАЛИТЬ
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnDel_Click(object sender, System.EventArgs e)
		{
            flag = 3;
            LockConponent(true);
		}

		/// <summary>
		/// Событие Click кнопки СОХРАНИТЬ
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

			if(flag == 1)// Добавление:
			{
			    DataRow[] row = ds11.Получатели.Select("ОписаниеПолучателя = '" + txt.Text.Trim() + "'");
			    if(row.Length > 0)
			    {
				    MessageBox.Show(this, "Получатель с таким описанием уже есть в справочнике.", "Добавление получателя", MessageBoxButtons.OK, MessageBoxIcon.Information);
				    flag = 0;
				    LockConponent(false);
				    return;
			    }
                получателиTableAdapter.Insert(txt.Text.Trim(), null, null);
                ПодключитьсяПолучитьДанные();
			}
			if(flag == 2)// Изменение: 
			{
                DataRow[] rows = ds11.Получатели.Select("id_получателя=" + (int)this.List.SelectedValue);
                rows[0]["ОписаниеПолучателя"] = txt.Text.Trim();
                получателиTableAdapter.Update(rows[0]);
                ПодключитьсяПолучитьДанные();
			}

            // Удалить сотрудника.
            if (flag == 3)
            {
               string query = "update dbo.Получатели " +
                               "set Удален = 'True' " +
                               "where LOWER(RTRIM(LTRIM(ОписаниеПолучателя))) =  '"+ this.txt.Text.Trim().ToLower() +"' ";

               ExecuteQuery execQuery = new ExecuteQuery(query);
               execQuery.Excecute();

                //DataRow[] rows = ds11.Получатели.Select("id_получателя=" + (int)this.List.SelectedValue);
                //rows[0]["ОписаниеПолучателя"] = txt.Text.Trim();
                //получателиTableAdapter.Update(rows[0]);
                ПодключитьсяПолучитьДанные();
            }
			flag = 0;
			LockConponent(false);
		}

		/// <summary>
		/// Событие Click кнопки ОТМЕНА
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			flag = 0;
			LockConponent(false);
			this.List_SelectedIndexChanged(null, null);
		}

        //Заполняет список получателей
        public void ПолучателиАдаптер(DS1.ПолучателиDataTable dt)
        {
            string query = "SELECT [id_получателя] " +
                           ",[ОписаниеПолучателя] " +
                           ",[ИмяРегистрации] " +
                           ",[Удален] " +
                           "FROM [Получатели] " +
                           "where [Удален] is null";
            SqlConnection con = new SqlConnection();
            con.ConnectionString = ConfigurationSettings.AppSettings["строкаДокументооборот"].ToString();
            SqlCommand com = new SqlCommand(query, con);

            con.Open();
            SqlDataReader read = com.ExecuteReader();

            while (read.Read())
            {
                DataRow row = dt.NewRow();
                row["id_получателя"] = read["id_получателя"];
                row["ОписаниеПолучателя"] = read["ОписаниеПолучателя"];
                row["ИмяРегистрации"] = read["ИмяРегистрации"];
                row["Удален"] = read["Удален"];
                dt.Rows.Add(row);
            }

            //return dt;
        }

	}
}
