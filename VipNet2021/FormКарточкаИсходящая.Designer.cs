/*
 * Created by SharpDevelop.
 * User: Денис Николаевич
 * Date: 25.03.2007
 * Time: 20:33
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace RegKor
{
	partial class FormКарточкаИсходящая : System.Windows.Forms.Form
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormКарточкаИсходящая));
            this.labelНомер = new System.Windows.Forms.Label();
            this.labelДата = new System.Windows.Forms.Label();
            this.dateTimeДата = new System.Windows.Forms.DateTimePicker();
            this.labelАдресат = new System.Windows.Forms.Label();
            this.comboBoxАдресат = new System.Windows.Forms.ComboBox();
            this.labelСодержание = new System.Windows.Forms.Label();
            this.textBoxСодержание = new System.Windows.Forms.TextBox();
            this.labelОтветНаДокумент = new System.Windows.Forms.Label();
            this.textBoxОтветНаДокумент = new System.Windows.Forms.TextBox();
            this.buttonОтмена = new System.Windows.Forms.Button();
            this.buttonСохранить = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.button1 = new System.Windows.Forms.Button();
            this.buttonОтветНаДокумент = new System.Windows.Forms.Button();
            this.maskedTextBox1 = new System.Windows.Forms.MaskedTextBox();
            this.errorProviderНомер = new System.Windows.Forms.ErrorProvider(this.components);
            this.errorProviderДата = new System.Windows.Forms.ErrorProvider(this.components);
            this.errorProviderАдресат = new System.Windows.Forms.ErrorProvider(this.components);
            this.errorProviderСодержание = new System.Windows.Forms.ErrorProvider(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.chkFlagPersonData = new System.Windows.Forms.CheckBox();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.linkLabel2 = new System.Windows.Forms.LinkLabel();
            this.chkBoxDsp = new System.Windows.Forms.CheckBox();
            this.ds11 = new RegKor.DS1();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderНомер)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderДата)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderАдресат)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderСодержание)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).BeginInit();
            this.SuspendLayout();
            // 
            // labelНомер
            // 
            this.labelНомер.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelНомер.Location = new System.Drawing.Point(3, 47);
            this.labelНомер.Name = "labelНомер";
            this.labelНомер.Size = new System.Drawing.Size(85, 32);
            this.labelНомер.TabIndex = 8;
            this.labelНомер.Text = "Номер:";
            this.labelНомер.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // labelДата
            // 
            this.labelДата.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelДата.Location = new System.Drawing.Point(-1, 9);
            this.labelДата.Name = "labelДата";
            this.labelДата.Size = new System.Drawing.Size(85, 29);
            this.labelДата.TabIndex = 9;
            this.labelДата.Text = "Дата:";
            this.labelДата.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // dateTimeДата
            // 
            this.dateTimeДата.CalendarTrailingForeColor = System.Drawing.SystemColors.Control;
            this.dateTimeДата.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dateTimeДата.Location = new System.Drawing.Point(89, 12);
            this.dateTimeДата.Name = "dateTimeДата";
            this.dateTimeДата.Size = new System.Drawing.Size(166, 22);
            this.dateTimeДата.TabIndex = 1;
            // 
            // labelАдресат
            // 
            this.labelАдресат.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelАдресат.Location = new System.Drawing.Point(1, 116);
            this.labelАдресат.Name = "labelАдресат";
            this.labelАдресат.Size = new System.Drawing.Size(87, 27);
            this.labelАдресат.TabIndex = 11;
            this.labelАдресат.Text = "Адресат:";
            this.labelАдресат.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // comboBoxАдресат
            // 
            this.comboBoxАдресат.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.comboBoxАдресат.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.comboBoxАдресат.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBoxАдресат.Location = new System.Drawing.Point(89, 117);
            this.comboBoxАдресат.MaxDropDownItems = 40;
            this.comboBoxАдресат.Name = "comboBoxАдресат";
            this.comboBoxАдресат.Size = new System.Drawing.Size(414, 24);
            this.comboBoxАдресат.TabIndex = 3;
            // 
            // labelСодержание
            // 
            this.labelСодержание.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelСодержание.Location = new System.Drawing.Point(7, 222);
            this.labelСодержание.Name = "labelСодержание";
            this.labelСодержание.Size = new System.Drawing.Size(105, 20);
            this.labelСодержание.TabIndex = 14;
            this.labelСодержание.Text = "Содержание:";
            this.labelСодержание.TextAlign = System.Drawing.ContentAlignment.BottomRight;
            // 
            // textBoxСодержание
            // 
            this.textBoxСодержание.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxСодержание.BackColor = System.Drawing.SystemColors.HighlightText;
            this.textBoxСодержание.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxСодержание.Location = new System.Drawing.Point(5, 245);
            this.textBoxСодержание.MaxLength = 250;
            this.textBoxСодержание.Multiline = true;
            this.textBoxСодержание.Name = "textBoxСодержание";
            this.textBoxСодержание.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxСодержание.Size = new System.Drawing.Size(532, 149);
            this.textBoxСодержание.TabIndex = 5;
            // 
            // labelОтветНаДокумент
            // 
            this.labelОтветНаДокумент.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelОтветНаДокумент.Location = new System.Drawing.Point(7, 397);
            this.labelОтветНаДокумент.Name = "labelОтветНаДокумент";
            this.labelОтветНаДокумент.Size = new System.Drawing.Size(220, 20);
            this.labelОтветНаДокумент.TabIndex = 18;
            this.labelОтветНаДокумент.Text = "Ответ на документ:";
            this.labelОтветНаДокумент.TextAlign = System.Drawing.ContentAlignment.BottomLeft;
            // 
            // textBoxОтветНаДокумент
            // 
            this.textBoxОтветНаДокумент.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxОтветНаДокумент.BackColor = System.Drawing.SystemColors.HighlightText;
            this.textBoxОтветНаДокумент.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.textBoxОтветНаДокумент.Location = new System.Drawing.Point(5, 422);
            this.textBoxОтветНаДокумент.MaxLength = 250;
            this.textBoxОтветНаДокумент.Multiline = true;
            this.textBoxОтветНаДокумент.Name = "textBoxОтветНаДокумент";
            this.textBoxОтветНаДокумент.ReadOnly = true;
            this.textBoxОтветНаДокумент.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxОтветНаДокумент.Size = new System.Drawing.Size(498, 42);
            this.textBoxОтветНаДокумент.TabIndex = 17;
            this.textBoxОтветНаДокумент.TabStop = false;
            // 
            // buttonОтмена
            // 
            this.buttonОтмена.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonОтмена.Dock = System.Windows.Forms.DockStyle.Right;
            this.buttonОтмена.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonОтмена.Location = new System.Drawing.Point(385, 0);
            this.buttonОтмена.Name = "buttonОтмена";
            this.buttonОтмена.Size = new System.Drawing.Size(178, 36);
            this.buttonОтмена.TabIndex = 8;
            this.buttonОтмена.Text = "Отмена";
            this.buttonОтмена.Click += new System.EventHandler(this.buttonОтмена_Click);
            // 
            // buttonСохранить
            // 
            this.buttonСохранить.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.buttonСохранить.Dock = System.Windows.Forms.DockStyle.Left;
            this.buttonСохранить.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonСохранить.Location = new System.Drawing.Point(0, 0);
            this.buttonСохранить.Name = "buttonСохранить";
            this.buttonСохранить.Size = new System.Drawing.Size(178, 36);
            this.buttonСохранить.TabIndex = 7;
            this.buttonСохранить.Text = "Сохранить";
            this.buttonСохранить.Click += new System.EventHandler(this.buttonСохранить_Click);
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.buttonОтмена);
            this.panel1.Controls.Add(this.buttonСохранить);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.panel1.Location = new System.Drawing.Point(0, 480);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(567, 40);
            this.panel1.TabIndex = 22;
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Location = new System.Drawing.Point(524, 117);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(22, 22);
            this.button1.TabIndex = 4;
            this.toolTip1.SetToolTip(this.button1, "Добавить адресата в справочник");
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // buttonОтветНаДокумент
            // 
            this.buttonОтветНаДокумент.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonОтветНаДокумент.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonОтветНаДокумент.Image = global::RegKor.Properties.Resources.add;
            this.buttonОтветНаДокумент.Location = new System.Drawing.Point(524, 422);
            this.buttonОтветНаДокумент.Name = "buttonОтветНаДокумент";
            this.buttonОтветНаДокумент.Size = new System.Drawing.Size(22, 22);
            this.buttonОтветНаДокумент.TabIndex = 6;
            this.toolTip1.SetToolTip(this.buttonОтветНаДокумент, "Открыть список входящих документов");
            this.buttonОтветНаДокумент.Click += new System.EventHandler(this.buttonОтветНаДокумент_Click);
            // 
            // maskedTextBox1
            // 
            this.maskedTextBox1.BeepOnError = true;
            this.maskedTextBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.maskedTextBox1.HidePromptOnLeave = true;
            this.maskedTextBox1.InsertKeyMode = System.Windows.Forms.InsertKeyMode.Overwrite;
            this.maskedTextBox1.Location = new System.Drawing.Point(89, 50);
            this.maskedTextBox1.Name = "maskedTextBox1";
            this.maskedTextBox1.PromptChar = ' ';
            this.maskedTextBox1.Size = new System.Drawing.Size(166, 22);
            this.maskedTextBox1.TabIndex = 2;
            this.maskedTextBox1.Text = "02-";
            this.maskedTextBox1.MouseUp += new System.Windows.Forms.MouseEventHandler(this.maskedTextBox1_MouseUp);
            this.maskedTextBox1.MaskInputRejected += new System.Windows.Forms.MaskInputRejectedEventHandler(this.maskedTextBox1_MaskInputRejected);
            this.maskedTextBox1.Leave += new System.EventHandler(this.maskedTextBox1_Leave);
            this.maskedTextBox1.Enter += new System.EventHandler(this.maskedTextBox1_Enter);
            this.maskedTextBox1.TextChanged += new System.EventHandler(this.maskedTextBox1_TextChanged);
            // 
            // errorProviderНомер
            // 
            this.errorProviderНомер.BlinkRate = 500;
            this.errorProviderНомер.ContainerControl = this;
            // 
            // errorProviderДата
            // 
            this.errorProviderДата.BlinkRate = 500;
            this.errorProviderДата.ContainerControl = this;
            // 
            // errorProviderАдресат
            // 
            this.errorProviderАдресат.BlinkRate = 500;
            this.errorProviderАдресат.ContainerControl = this;
            // 
            // errorProviderСодержание
            // 
            this.errorProviderСодержание.BlinkRate = 500;
            this.errorProviderСодержание.ContainerControl = this;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(273, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 16);
            this.label1.TabIndex = 24;
            // 
            // chkFlagPersonData
            // 
            this.chkFlagPersonData.AutoSize = true;
            this.chkFlagPersonData.Location = new System.Drawing.Point(276, 13);
            this.chkFlagPersonData.Name = "chkFlagPersonData";
            this.chkFlagPersonData.Size = new System.Drawing.Size(143, 17);
            this.chkFlagPersonData.TabIndex = 25;
            this.chkFlagPersonData.Text = "Персональные данные";
            this.chkFlagPersonData.UseVisualStyleBackColor = true;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(89, 148);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(184, 23);
            this.button2.TabIndex = 26;
            this.button2.Text = "Основание передачи";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Image = global::RegKor.Properties.Resources.add;
            this.button3.Location = new System.Drawing.Point(524, 145);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(22, 22);
            this.button3.TabIndex = 27;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.ForeColor = System.Drawing.Color.Green;
            this.label2.Location = new System.Drawing.Point(13, 163);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(0, 13);
            this.label2.TabIndex = 28;
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(157, 191);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(0, 13);
            this.linkLabel1.TabIndex = 29;
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // linkLabel2
            // 
            this.linkLabel2.AutoSize = true;
            this.linkLabel2.Location = new System.Drawing.Point(157, 209);
            this.linkLabel2.Name = "linkLabel2";
            this.linkLabel2.Size = new System.Drawing.Size(0, 13);
            this.linkLabel2.TabIndex = 30;
            this.linkLabel2.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel2_LinkClicked);
            // 
            // chkBoxDsp
            // 
            this.chkBoxDsp.AutoSize = true;
            this.chkBoxDsp.Location = new System.Drawing.Point(89, 90);
            this.chkBoxDsp.Name = "chkBoxDsp";
            this.chkBoxDsp.Size = new System.Drawing.Size(50, 17);
            this.chkBoxDsp.TabIndex = 31;
            this.chkBoxDsp.Text = "ДСП";
            this.chkBoxDsp.UseVisualStyleBackColor = true;
            this.chkBoxDsp.CheckedChanged += new System.EventHandler(this.chkBoxDsp_CheckedChanged);
            // 
            // ds11
            // 
            this.ds11.DataSetName = "DS1";
            this.ds11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // FormКарточкаИсходящая
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonОтмена;
            this.ClientSize = new System.Drawing.Size(567, 520);
            this.Controls.Add(this.chkBoxDsp);
            this.Controls.Add(this.linkLabel2);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.chkFlagPersonData);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.maskedTextBox1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.buttonОтветНаДокумент);
            this.Controls.Add(this.labelОтветНаДокумент);
            this.Controls.Add(this.textBoxОтветНаДокумент);
            this.Controls.Add(this.labelСодержание);
            this.Controls.Add(this.textBoxСодержание);
            this.Controls.Add(this.labelАдресат);
            this.Controls.Add(this.comboBoxАдресат);
            this.Controls.Add(this.labelДата);
            this.Controls.Add(this.dateTimeДата);
            this.Controls.Add(this.labelНомер);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormКарточкаИсходящая";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Карточка исходящего документа";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormКарточкаИсходящая_FormClosed);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormКарточкаИсходящая_FormClosing);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderНомер)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderДата)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderАдресат)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProviderСодержание)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        private System.Windows.Forms.Label labelНомер;
        private DS1 ds11;
        private System.Windows.Forms.Label labelДата;
        public System.Windows.Forms.DateTimePicker dateTimeДата;
        private System.Windows.Forms.Label labelАдресат;
        public System.Windows.Forms.ComboBox comboBoxАдресат;
        private System.Windows.Forms.Label labelСодержание;
        public System.Windows.Forms.TextBox textBoxСодержание;
        private System.Windows.Forms.Button buttonОтветНаДокумент;
        private System.Windows.Forms.Label labelОтветНаДокумент;
        public System.Windows.Forms.TextBox textBoxОтветНаДокумент;
        private System.Windows.Forms.Button buttonОтмена;
        private System.Windows.Forms.Button buttonСохранить;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.MaskedTextBox maskedTextBox1;
        private System.Windows.Forms.ErrorProvider errorProviderНомер;
        private System.Windows.Forms.ErrorProvider errorProviderДата;
        private System.Windows.Forms.ErrorProvider errorProviderАдресат;
        private System.Windows.Forms.ErrorProvider errorProviderСодержание;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox chkFlagPersonData;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.LinkLabel linkLabel2;
        private System.Windows.Forms.CheckBox chkBoxDsp;
	}
}
