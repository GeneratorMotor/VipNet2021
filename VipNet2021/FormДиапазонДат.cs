using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;

namespace RegKor
{
	/// <summary>
	/// Summary description for FormDateRange.
	/// </summary>
	public class FormДиапазонДат : System.Windows.Forms.Form
	{
		
		public string BeginDate = null;
		public string EndDate = null;
        RegKor.DS1 ds11;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.DateTimePicker dt1;
		private System.Windows.Forms.DateTimePicker dt2;
        private Panel panel1;
        private GroupBox groupBox2;
        private Label label4;
        private Label label5;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FormДиапазонДат(RegKor.DS1 ds)
		{
			InitializeComponent();
			this.ds11 = ds;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormДиапазонДат));
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.dt1 = new System.Windows.Forms.DateTimePicker();
            this.dt2 = new System.Windows.Forms.DateTimePicker();
            this.ds11 = new RegKor.DS1();
            this.panel1 = new System.Windows.Forms.Panel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).BeginInit();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Top;
            this.button1.Location = new System.Drawing.Point(4, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(114, 39);
            this.button1.TabIndex = 5;
            this.button1.Text = "Готово";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.button2.Location = new System.Drawing.Point(4, 54);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(114, 39);
            this.button2.TabIndex = 6;
            this.button2.Text = "Отмена";
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // dt1
            // 
            this.dt1.Location = new System.Drawing.Point(31, 54);
            this.dt1.Name = "dt1";
            this.dt1.Size = new System.Drawing.Size(142, 20);
            this.dt1.TabIndex = 7;
            this.dt1.Value = new System.DateTime(2006, 8, 28, 0, 0, 0, 0);
            // 
            // dt2
            // 
            this.dt2.Location = new System.Drawing.Point(212, 54);
            this.dt2.Name = "dt2";
            this.dt2.Size = new System.Drawing.Size(142, 20);
            this.dt2.TabIndex = 8;
            // 
            // ds11
            // 
            this.ds11.DataSetName = "DS1";
            this.ds11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel1.Location = new System.Drawing.Point(398, 4);
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding(4);
            this.panel1.Size = new System.Drawing.Size(122, 97);
            this.panel1.TabIndex = 9;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.dt2);
            this.groupBox2.Controls.Add(this.dt1);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point(4, 4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(394, 97);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "За какой период вывести статистику?";
            this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(31, 33);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(156, 18);
            this.label4.TabIndex = 3;
            this.label4.Text = "Начало отчетного периода:";
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(212, 33);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(156, 18);
            this.label5.TabIndex = 4;
            this.label5.Text = "Конец отчетного периода:";
            // 
            // FormДиапазонДат
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(524, 105);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormДиапазонДат";
            this.Padding = new System.Windows.Forms.Padding(4);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Статистика поступления корреспонденции";
            this.Load += new System.EventHandler(this.FormDateRange_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ds11)).EndInit();
            this.panel1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Событие Click кнопки ОТМЕНА
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void button2_Click(object sender, System.EventArgs e)
		{
			BeginDate = null;
			EndDate = null;
			Dispose(true);
		}

		/// <summary>
		/// Событие Click кнопки ГОТОВО
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void button1_Click(object sender, System.EventArgs e)
		{
			this.Enabled = false;
			FormView frmPrint = new FormView();
			BeginDate = dt1.Value.ToShortDateString();
			EndDate = dt2.Value.ToShortDateString();
			try
			{
				ReportDocument rptDoc = new ReportDocument();
				string fileName = @"..\report\Statistic.rpt";
				// загружает файл отчета:
				rptDoc.Load(fileName);   
				// источник данных:
				rptDoc.SetDataSource(ds11);
				// просмотрщику передали источник отчета:
				frmPrint.reportViewer.ReportSource = rptDoc;
				FormГлавная.ПараметрыДляОтчета("BeginDate", BeginDate, frmPrint.reportViewer.ParameterFieldInfo);
				FormГлавная.ПараметрыДляОтчета("EndDate", EndDate, frmPrint.reportViewer.ParameterFieldInfo);
				frmPrint.Text = "Статистика поступления";
				// показываем форму:
				frmPrint.reportViewer.ShowGroupTreeButton = false;
                this.Hide();
				frmPrint.ShowDialog( this );
			}
			catch(Exception exc)
			{
				MessageBox.Show(this, "Произошла ошибка при открытии файла отчета \"Статистика поступления\".\n" + exc.Message, "Ошибка открытия файла отчета");
				return;
			}
			finally
			{	
				this.Enabled = true;
				BeginDate = null;
				EndDate = null;
				Dispose(true);
			}
		}

		private void FormDateRange_Load(object sender, System.EventArgs e)
		{
			dt1.Value = DateTime.Now.AddMonths(-1);//начало отчетного периода на 1 месяц меньше сегодняшней даты
			dt2.Value = DateTime.Now;// конец отчетного период равен сегодняшняй дате
		}

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
	}
}
