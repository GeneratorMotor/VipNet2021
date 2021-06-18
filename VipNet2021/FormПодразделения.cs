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
	public class FormПодразделения : System.Windows.Forms.Form
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
        private System.Windows.Forms.TextBox textBoxРуководитель;
        private System.Windows.Forms.TextBox textBoxНомерПодразделения;
        private System.Windows.Forms.TextBox textBoxБуквенноеОбозначение;
		private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnCancel;
        private IContainer components;

        /// <summary>
        /// Датаадаптер для взаимодействия с источником данных
        /// </summary>
        DS1TableAdapters.ПодразделенияКомитетаTableAdapter подразделенияTableAdapter;
        DS1TableAdapters.ПолучателиTableAdapter руководителиTableAdapter;
        private Panel panel1;
        private Label label1;
        private Panel panel2;
        private Label label2;
        private Label label3;
        private Button buttonСписокСотрудников;
        private Label label4;
        private ToolTip toolTip1;
        private Label labelРуководитель;
        private object id_руководителя;
        private DS1 ds11;
        private RichTextBox txtОписаниеПодразделения;


        public FormПодразделения ( DS1 dataset )
		{
			InitializeComponent();
            this.ds11 = dataset;

            подразделенияTableAdapter = new RegKor.DS1TableAdapters.ПодразделенияКомитетаTableAdapter( );
            руководителиTableAdapter = new RegKor.DS1TableAdapters.ПолучателиTableAdapter( );

			ПодключитьсяПолучитьДанные();
		}

		private void ПодключитьсяПолучитьДанные()
		{
            ds11.ПодразделенияКомитета.Clear( );
            подразделенияTableAdapter.Fill( ds11.ПодразделенияКомитета );

			// привязка к источнику данных
            this.List.DataSource = ds11.ПодразделенияКомитета.Select( "Удален=False", "НомерПодразделения, БуквенноеОбозначение" );
            this.List.ValueMember = "id_подразделения";
            this.List.DisplayMember = "ОписаниеПодразделения";
            this.List.SelectedValue = "id_подразделения";
			// выделяем первый элемент в списке:
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormПодразделения));
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.List = new System.Windows.Forms.ListBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtОписаниеПодразделения = new System.Windows.Forms.RichTextBox();
            this.labelРуководитель = new System.Windows.Forms.Label();
            this.buttonСписокСотрудников = new System.Windows.Forms.Button();
            this.textBoxРуководитель = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxБуквенноеОбозначение = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxНомерПодразделения = new System.Windows.Forms.TextBox();
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
            this.btnAdd.Text = "Добавить";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(3, 39);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(106, 30);
            this.btnEdit.TabIndex = 2;
            this.btnEdit.Text = "Изменить";
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnDel
            // 
            this.btnDel.Enabled = false;
            this.btnDel.Location = new System.Drawing.Point(3, 75);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(106, 30);
            this.btnDel.TabIndex = 3;
            this.btnDel.Text = "Удалить";
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
            this.btnClose.Text = "Закрыть";
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
            this.btnCancel.Text = "Отмена";
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.txtОписаниеПодразделения);
            this.panel1.Controls.Add(this.labelРуководитель);
            this.panel1.Controls.Add(this.buttonСписокСотрудников);
            this.panel1.Controls.Add(this.textBoxРуководитель);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.textBoxБуквенноеОбозначение);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.textBoxНомерПодразделения);
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
            // txtОписаниеПодразделения
            // 
            this.txtОписаниеПодразделения.Location = new System.Drawing.Point(117, 23);
            this.txtОписаниеПодразделения.Name = "txtОписаниеПодразделения";
            this.txtОписаниеПодразделения.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.txtОписаниеПодразделения.Size = new System.Drawing.Size(323, 85);
            this.txtОписаниеПодразделения.TabIndex = 4;
            this.txtОписаниеПодразделения.Text = "";
            // 
            // labelРуководитель
            // 
            this.labelРуководитель.AutoSize = true;
            this.labelРуководитель.Location = new System.Drawing.Point(266, 159);
            this.labelРуководитель.Name = "labelРуководитель";
            this.labelРуководитель.Size = new System.Drawing.Size(0, 13);
            this.labelРуководитель.TabIndex = 18;
            // 
            // buttonСписокСотрудников
            // 
            this.buttonСписокСотрудников.Image = global::RegKor.Properties.Resources.add;
            this.buttonСписокСотрудников.Location = new System.Drawing.Point(416, 114);
            this.buttonСписокСотрудников.Name = "buttonСписокСотрудников";
            this.buttonСписокСотрудников.Size = new System.Drawing.Size(23, 23);
            this.buttonСписокСотрудников.TabIndex = 5;
            this.toolTip1.SetToolTip(this.buttonСписокСотрудников, "Открыть список сотрудников");
            this.buttonСписокСотрудников.UseVisualStyleBackColor = true;
            this.buttonСписокСотрудников.Click += new System.EventHandler(this.buttonСписокСотрудников_Click);
            // 
            // textBoxРуководитель
            // 
            this.textBoxРуководитель.BackColor = System.Drawing.SystemColors.Control;
            this.textBoxРуководитель.Location = new System.Drawing.Point(117, 116);
            this.textBoxРуководитель.Name = "textBoxРуководитель";
            this.textBoxРуководитель.ReadOnly = true;
            this.textBoxРуководитель.Size = new System.Drawing.Size(294, 20);
            this.textBoxРуководитель.TabIndex = 16;
            this.textBoxРуководитель.TabStop = false;
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(16, 116);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(97, 17);
            this.label4.TabIndex = 15;
            this.label4.Text = "Руководитель";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBoxБуквенноеОбозначение
            // 
            this.textBoxБуквенноеОбозначение.Location = new System.Drawing.Point(117, 179);
            this.textBoxБуквенноеОбозначение.MaxLength = 1;
            this.textBoxБуквенноеОбозначение.Name = "textBoxБуквенноеОбозначение";
            this.textBoxБуквенноеОбозначение.Size = new System.Drawing.Size(40, 20);
            this.textBoxБуквенноеОбозначение.TabIndex = 7;
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(1, 179);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(110, 17);
            this.label3.TabIndex = 13;
            this.label3.Text = "Символьная литера";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // textBoxНомерПодразделения
            // 
            this.textBoxНомерПодразделения.Location = new System.Drawing.Point(117, 148);
            this.textBoxНомерПодразделения.MaxLength = 2;
            this.textBoxНомерПодразделения.Name = "textBoxНомерПодразделения";
            this.textBoxНомерПодразделения.Size = new System.Drawing.Size(61, 20);
            this.textBoxНомерПодразделения.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(11, 148);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 17);
            this.label2.TabIndex = 11;
            this.label2.Text = "Цифрвой код";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(16, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 37);
            this.label1.TabIndex = 10;
            this.label1.Text = "Описание подразделения";
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
            // FormПодразделения
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
            this.Name = "FormПодразделения";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Справочник \"Подразделения\"";
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
			panel1.Enabled = val;
		}
	
		/// <summary>
		/// Событие Click кнопки ЗАКРЫТЬ
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btnClose_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		/// <summary>
		/// Событие изменения выделения в списке
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void List_SelectedIndexChanged(object sender, System.EventArgs e)
		{
            try
            {
                DataRow [ ] row = ds11.ПодразделенияКомитета.Select( "id_подразделения=" + List.SelectedValue );
                txtОписаниеПодразделения.Text = Convert.ToString( row [0] ["ОписаниеПодразделения"] );
                textBoxНомерПодразделения.Text = Convert.ToString( row [0] ["НомерПодразделения"] );
                textBoxБуквенноеОбозначение.Text = Convert.ToString( row [0] ["БуквенноеОбозначение"] );
                DataRow [ ] row2 = ds11.Получатели.Select( "id_получателя=" + row [0] ["id_РуководителяПодразделения"] );
                if ( row2.Length > 0 )
                {
                    id_руководителя = Convert.ToInt32( row2 [0] ["id_получателя"] );
                    textBoxРуководитель.Text = Convert.ToString( row2 [0] ["ОписаниеПолучателя"] );
                }
                else
                {
                    id_руководителя = null;
                    textBoxРуководитель.Text = "";
                }
            }
            catch
            {
            }
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
            txtОписаниеПодразделения.Text = "";
            textBoxРуководитель.Text = "";
            textBoxБуквенноеОбозначение.Text = "";
            textBoxНомерПодразделения.Text = "";
			txtОписаниеПодразделения.Focus();
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
            if ( txtОписаниеПодразделения.Text.Trim( ) == "" || textBoxРуководитель.Text.Trim( ) == "" || textBoxНомерПодразделения.Text.Trim( ) == "")
			{
                if ( flag == 1 )
                {
                    txtОписаниеПодразделения.Text = "";
                    textBoxРуководитель.Text = "";
                    textBoxБуквенноеОбозначение.Text = "";
                    textBoxНомерПодразделения.Text = "";
                }
                if ( flag == 2 )
                {
                    try
                    {
                        DataRow [ ] row = ds11.ПодразделенияКомитета.Select( "id_подразделения=" + List.SelectedValue );
                        txtОписаниеПодразделения.Text = Convert.ToString( row [0] ["ОписаниеПодразделения"] );
                        textBoxНомерПодразделения.Text = Convert.ToString( row [0] ["НомерПодразделения"] );
                        textBoxБуквенноеОбозначение.Text = Convert.ToString( row [0] ["БуквенноеОбозначение"] );
                        DataRow [ ] row2 = ds11.Получатели.Select( "id_получателя=" + row [0] ["id_РуководителяПодразделения"] );
                        if ( row2.Length > 0 )
                        {
                            id_руководителя = Convert.ToInt32( row2 [0] ["id_получателя"] );
                            textBoxРуководитель.Text = Convert.ToString( row2 [0] ["ОписаниеПолучателя"] );
                        }
                        else
                        {
                            id_руководителя = null;
                            textBoxРуководитель.Text = "";
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

			if(flag == 2)// Изменение: 
			{
                DataRow [ ] rows = ds11.ПодразделенияКомитета.Select( "id_подразделения=" + ( int ) this.List.SelectedValue );

                rows [0] ["ОписаниеПодразделения"] = txtОписаниеПодразделения.Text.Trim( );
                rows [0] ["id_РуководителяПодразделения"] = id_руководителя;
                rows [0] ["НомерПодразделения"] = textBoxНомерПодразделения.Text.Trim( );
                rows [0] ["БуквенноеОбозначение"] = textBoxБуквенноеОбозначение.Text.Trim( );

                подразделенияTableAdapter.Update( rows [0] );

                ПодключитьсяПолучитьДанные( );
			}

			if(flag == 1)// Добавление:
			{
                DataRow [ ] row = ds11.ПодразделенияКомитета.Select( "ОписаниеПодразделения = '" + txtОписаниеПодразделения.Text.Trim( ) + "'");
                if ( row.Length > 0 )
                {
                    MessageBox.Show( this, "Подразделение с таким описанием уже есть в справочнике.", "Добавление подразделения", MessageBoxButtons.OK, MessageBoxIcon.Information );
                    flag = 0;
                    LockConponent( false );
                    return;
                }

                подразделенияTableAdapter.Insert( txtОписаниеПодразделения.Text.Trim( ), ( int ) id_руководителя, textBoxНомерПодразделения.Text.Trim( ), textBoxБуквенноеОбозначение.Text.Trim( ), false );
                
                ПодключитьсяПолучитьДанные( );
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
            if ( flag == 1 )
            {
                txtОписаниеПодразделения.Text = "";
                textBoxРуководитель.Text = "";
                textBoxБуквенноеОбозначение.Text = "";
                textBoxНомерПодразделения.Text = "";
            }
            if ( flag == 2 )
            {
                try
                {
                    DataRow [ ] row = ds11.ПодразделенияКомитета.Select( "id_подразделения=" + List.SelectedValue );
                    txtОписаниеПодразделения.Text = Convert.ToString( row [0] ["ОписаниеПодразделения"] );
                    textBoxНомерПодразделения.Text = Convert.ToString( row [0] ["НомерПодразделения"] );
                    textBoxБуквенноеОбозначение.Text = Convert.ToString( row [0] ["БуквенноеОбозначение"] );
                    DataRow [ ] row2 = ds11.Получатели.Select( "id_получателя=" + row [0] ["id_РуководителяПодразделения"] );
                    if ( row2.Length > 0 )
                    {
                        id_руководителя = Convert.ToInt32( row2 [0] ["id_получателя"] );
                        textBoxРуководитель.Text = Convert.ToString( row2 [0] ["ОписаниеПолучателя"] );
                    }
                    else
                    {
                        id_руководителя = null;
                        textBoxРуководитель.Text = "";
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

        private void buttonСписокСотрудников_Click ( object sender, EventArgs e )
        {
            FormСписокСотрудников form = new FormСписокСотрудников( );
            DialogResult result = form.ShowDialog( this );

            if ( result == DialogResult.OK )
            {
                //id_руководителя = (int)form.ИДСотрудника;
                DataRow [ ] row2 = ds11.Получатели.Select( "id_получателя=" + form.ИДСотрудника );
                if ( row2.Length > 0 )
                {
                    id_руководителя = Convert.ToInt32( row2 [0] ["id_получателя"] );
                    textBoxРуководитель.Text = Convert.ToString( row2 [0] ["ОписаниеПолучателя"] );
                }
                //textBoxРуководитель.Text = 
                //строкаИсходящейКарточки ["id_ВходящегоДокумента"] = form.ИДВходящегоДокумента;
                //ИДВходящегоДокумента = ( int ) form.ИДВходящегоДокумента;
                //textBoxОтветНаДокумент.Text = ОтветНаДокумент( form.ИДВходящегоДокумента );
            }
        }

	}
}
