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
	public class FormКорреспонденты : System.Windows.Forms.Form
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
        DS1TableAdapters.КорреспондентыTableAdapter корреспондентыTableAdapter;
        private CheckBox checkBox1;
        private Button btnEditClose;
        private GroupBox groupBox1;
        private Button btnOpen;
        private CheckBox flagОткрытьСкрытое;

        /// <summary>
        /// Датасет - локальная копия данных
        /// </summary>
        private DS1 ds11;

        public FormКорреспонденты ( string новыйКорреспондент )
        {
            InitializeComponent( );
            корреспондентыTableAdapter = new RegKor.DS1TableAdapters.КорреспондентыTableAdapter( );
            ПодключитьсяПолучитьДанные( );
            flag = 1;
            LockConponent( true );
            txt.Text = новыйКорреспондент;
        }

		public FormКорреспонденты()
		{
            InitializeComponent();
            корреспондентыTableAdapter = new RegKor.DS1TableAdapters.КорреспондентыTableAdapter();
			ПодключитьсяПолучитьДанные();
		}

		private void ПодключитьсяПолучитьДанные()
		{
            //Обнуляем DataSet
            ds11.Корреспонденты.Clear();

            //=====================

            ////Определем ключ в конфигурационном файле если ГодДокумента = "истина" тогда 2012 год или больше
            ////если ГодДокумента = "ложно" тогда 2011 год или раньше
            //Classess.ГодДокументооборота год = new RegKor.Classess.ГодДокументооборота();
            //bool flag = год.ГодВКонфигурационномФайле();

            ////если 2011 год или раньше тогда ds11 типа DataSet заполняем как обычно
            //if (flag == false)
            //{
            //    //Заполняем DataSet ds11 данными
            //    корреспондентыTableAdapter.Fill(ds11.Корреспонденты);
            //}

            ////если 2012 год или больше тогда заполняем ds11 типа DataSet данными где в таблице
            ////Корреспонденты поле Удалён = NULL  

            //if (flag == true && this.flagОткрытьСкрытое.Checked == false)
            //{
            //    Classess.Корреспонденты корреспонденты = new RegKor.Classess.Корреспонденты();
            //    DataSet dsКорреспонденты = корреспонденты.ЗаполнитьКорреспонденты_DataSet();

            //    //foreach (DataRow rowКорреспонденты in dsКорреспонденты.Tables[0].Rows)
            //    //{
            //    //    DataRow row1 = ds11.Корреспонденты.NewRow();
            //    //    row1[0] = rowКорреспонденты[0];
            //    //    row1[1] = rowКорреспонденты[1];
            //    //    row1[2] = rowКорреспонденты[2];
            //    //    ds11.Корреспонденты.Rows.Add(row1);
            //    //}

            //}

            //////Открываем скрытые записи
            //if (flag == true && this.flagОткрытьСкрытое.Checked == true)
            //{

            //    Корреспонденты корреспонденты = new Корреспонденты();
            //    DataSet dsКорреспондентыСкрытые = корреспонденты.ПоказатьСкрытые();

            //    //foreach (DataRow rowКорреспонденты in dsКорреспондентыСкрытые.Tables[0].Rows)
            //    //{
            //    //    DataRow row1 = ds11.Корреспонденты.NewRow();
            //    //    row1[0] = rowКорреспонденты[0];
            //    //    row1[1] = rowКорреспонденты[1];
            //    //    row1[2] = rowКорреспонденты[2];
            //    //    ds11.Корреспонденты.Rows.Add(row1);
            //    //}

            //}


            //=======================

            ////Заполняем DataSet ds11 данными
            корреспондентыTableAdapter.Fill(ds11.Корреспонденты); //- Старый рабочий метод

			// привязка к источнику данных
			this.List.DataSource = ds11.Корреспонденты.Select("", "ОписаниеКорреспондента");
    		this.List.ValueMember = "id_корреспондента";
			this.List.DisplayMember = "ОписаниеКорреспондента";
			this.List.SelectedValue = "id_корреспондента";

            int iCount =   ds11.Корреспонденты.Rows.Count;

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormКорреспонденты));
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
            this.flagОткрытьСкрытое = new System.Windows.Forms.CheckBox();
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
            this.btnAdd.Text = "Добавить";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(268, 129);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(106, 30);
            this.btnEdit.TabIndex = 3;
            this.btnEdit.Text = "Изменить";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnDel
            // 
            this.btnDel.Enabled = false;
            this.btnDel.Location = new System.Drawing.Point(268, 167);
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
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Checked = true;
            this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBox1.Location = new System.Drawing.Point(271, 69);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(198, 17);
            this.checkBox1.TabIndex = 10;
            this.checkBox1.Text = "Для документооборота 2012 года";
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
            this.btnEditClose.Text = "Скрыть";
            this.btnEditClose.UseVisualStyleBackColor = true;
            this.btnEditClose.Visible = false;
            this.btnEditClose.Click += new System.EventHandler(this.btnEditClose_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnOpen);
            this.groupBox1.Controls.Add(this.flagОткрытьСкрытое);
            this.groupBox1.Location = new System.Drawing.Point(381, 93);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(169, 104);
            this.groupBox1.TabIndex = 12;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Корреспонденты";
            this.groupBox1.Visible = false;
            // 
            // btnOpen
            // 
            this.btnOpen.Enabled = false;
            this.btnOpen.Location = new System.Drawing.Point(7, 60);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(106, 30);
            this.btnOpen.TabIndex = 1;
            this.btnOpen.Text = "Открыть";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // flagОткрытьСкрытое
            // 
            this.flagОткрытьСкрытое.AutoSize = true;
            this.flagОткрытьСкрытое.Location = new System.Drawing.Point(6, 36);
            this.flagОткрытьСкрытое.Name = "flagОткрытьСкрытое";
            this.flagОткрытьСкрытое.Size = new System.Drawing.Size(118, 17);
            this.flagОткрытьСкрытое.TabIndex = 0;
            this.flagОткрытьСкрытое.Text = "Открыть скрытые";
            this.flagОткрытьСкрытое.UseVisualStyleBackColor = true;
            this.flagОткрытьСкрытое.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
            // 
            // FormКорреспонденты
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
            this.Name = "FormКорреспонденты";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Справочник \"Адресаты\"";
            this.Load += new System.EventHandler(this.FormКорреспонденты_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

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
			this.Close();
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
//			flag = 3;
//			LockConponent(true);
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
			    DataRow[] row = ds11.Корреспонденты.Select("ОписаниеКорреспондента = '" + txt.Text.Trim() + "'");
			    if(row.Length > 0)
			    {
				    MessageBox.Show(this, "Корреспондент с таким описанием уже есть в справочнике.", "Добавление корреспондента", MessageBoxButtons.OK, MessageBoxIcon.Information);
				    flag = 0;
				    LockConponent(false);
				    return;
			    }
                корреспондентыTableAdapter.Insert(txt.Text.Trim(), null);
                ПодключитьсяПолучитьДанные();
			}
			if(flag == 2)// Изменение: 
			{
				DataRow[] rows = ds11.Корреспонденты.Select("id_корреспондента=" + (int)this.List.SelectedValue);
                rows[0]["ОписаниеКорреспондента"] = txt.Text.Trim();
                корреспондентыTableAdapter.Update(rows[0]);
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
            Корреспонденты корреспонденты = new Корреспонденты();
            корреспонденты.Скрыть(this.txt.Text);

            ПодключитьсяПолучитьДанные();
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (this.flagОткрытьСкрытое.Checked == true)
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
            ПодключитьсяПолучитьДанные();


            //Classess.Корреспонденты корреспонденты = new RegKor.Classess.Корреспонденты();
            //DataSet dsКорреспонденты = корреспонденты.ЗаполнитьКорреспонденты_DataSet();

            //foreach (DataRow rowКорреспонденты in dsКорреспонденты.Tables[0].Rows)
            //{
            //    DataRow row1 = ds11.Корреспонденты.NewRow();
            //    row1[0] = rowКорреспонденты[0];
            //    row1[1] = rowКорреспонденты[1];
            //    ds11.Корреспонденты.Rows.Add(row1);
            //}
        }

        private void FormКорреспонденты_Load(object sender, EventArgs e)
        {
            RegKor.Classess.ГодДокументооборота годДокумента = new ГодДокументооборота();
            bool flag = годДокумента.ГодВКонфигурационномФайле();

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
                this.flagОткрытьСкрытое.Visible = true;
                this.btnOpen.Visible = true;
                this.btnEditClose.Visible = true;
                this.checkBox1.Visible = true;
            }
            else
            {
                this.groupBox1.Visible = false;
                this.flagОткрытьСкрытое.Visible = false;
                this.btnOpen.Visible = false;
                this.btnEditClose.Visible = false;
                this.checkBox1.Visible = false;
            }
            
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Корреспонденты корреспонденты = new Корреспонденты();
            корреспонденты.Открыть(this.txt.Text);
            this.flagОткрытьСкрытое.Checked = false;
            ПодключитьсяПолучитьДанные();
            this.btnOpen.Enabled = false;
        }

        

	}
}
