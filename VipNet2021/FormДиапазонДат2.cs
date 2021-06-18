using System;
using System.Data;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;

namespace RegKor
{
	public class Form�����������2 : System.Windows.Forms.Form
	{
		
		public string BeginDate = null;
        public string EndDate = null;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.DateTimePicker dt1;
        private System.Windows.Forms.DateTimePicker dt2;
        private GroupBox groupBox2;
        private Panel panel1;
        private DS���������������������� ds����������������������1;
		private System.ComponentModel.Container components = null;
        RegKor.DS1 ds;
		public Form�����������2(RegKor.DS1 ds)
		{

			InitializeComponent();
			this.ds = ds;
            this.ds����������������������1 = new DS����������������������( );
		}

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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager( typeof( Form�����������2 ) );
            this.label2 = new System.Windows.Forms.Label( );
            this.label3 = new System.Windows.Forms.Label( );
            this.button1 = new System.Windows.Forms.Button( );
            this.button2 = new System.Windows.Forms.Button( );
            this.dt1 = new System.Windows.Forms.DateTimePicker( );
            this.dt2 = new System.Windows.Forms.DateTimePicker( );
            this.groupBox2 = new System.Windows.Forms.GroupBox( );
            this.panel1 = new System.Windows.Forms.Panel( );
            this.ds����������������������1 = new RegKor.DS����������������������( );
            this.groupBox2.SuspendLayout( );
            this.panel1.SuspendLayout( );
            ( ( System.ComponentModel.ISupportInitialize ) ( this.ds����������������������1 ) ).BeginInit( );
            this.SuspendLayout( );
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point( 31, 29 );
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size( 156, 18 );
            this.label2.TabIndex = 3;
            this.label2.Text = "������ ��������� �������:";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point( 212, 31 );
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size( 156, 17 );
            this.label3.TabIndex = 4;
            this.label3.Text = "����� ��������� �������:";
            // 
            // button1
            // 
            this.button1.Dock = System.Windows.Forms.DockStyle.Top;
            this.button1.Location = new System.Drawing.Point( 4, 4 );
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size( 100, 42 );
            this.button1.TabIndex = 5;
            this.button1.Text = "������";
            this.button1.Click += new System.EventHandler( this.button1_Click );
            // 
            // button2
            // 
            this.button2.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.button2.Location = new System.Drawing.Point( 4, 60 );
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size( 100, 42 );
            this.button2.TabIndex = 6;
            this.button2.Text = "������";
            this.button2.Click += new System.EventHandler( this.button2_Click );
            // 
            // dt1
            // 
            this.dt1.Location = new System.Drawing.Point( 31, 49 );
            this.dt1.Name = "dt1";
            this.dt1.Size = new System.Drawing.Size( 156, 20 );
            this.dt1.TabIndex = 7;
            this.dt1.Value = new System.DateTime( 2006, 8, 28, 0, 0, 0, 0 );
            // 
            // dt2
            // 
            this.dt2.Location = new System.Drawing.Point( 212, 49 );
            this.dt2.Name = "dt2";
            this.dt2.Size = new System.Drawing.Size( 156, 20 );
            this.dt2.TabIndex = 8;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add( this.label2 );
            this.groupBox2.Controls.Add( this.dt1 );
            this.groupBox2.Controls.Add( this.dt2 );
            this.groupBox2.Controls.Add( this.label3 );
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupBox2.Location = new System.Drawing.Point( 4, 4 );
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size( 408, 106 );
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "�� ����� ������ ������� ����������?";
            // 
            // panel1
            // 
            this.panel1.Controls.Add( this.button2 );
            this.panel1.Controls.Add( this.button1 );
            this.panel1.Dock = System.Windows.Forms.DockStyle.Right;
            this.panel1.Location = new System.Drawing.Point( 412, 4 );
            this.panel1.Name = "panel1";
            this.panel1.Padding = new System.Windows.Forms.Padding( 4 );
            this.panel1.Size = new System.Drawing.Size( 108, 106 );
            this.panel1.TabIndex = 12;
            // 
            // ds����������������������1
            // 
            this.ds����������������������1.DataSetName = "DS����������������������";
            this.ds����������������������1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // Form�����������2
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size( 5, 13 );
            this.ClientSize = new System.Drawing.Size( 524, 114 );
            this.Controls.Add( this.groupBox2 );
            this.Controls.Add( this.panel1 );
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ( ( System.Drawing.Icon ) ( resources.GetObject( "$this.Icon" ) ) );
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form�����������2";
            this.Padding = new System.Windows.Forms.Padding( 4 );
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "���������� �� ������������";
            this.Load += new System.EventHandler( this.FormDateRange_Load );
            this.groupBox2.ResumeLayout( false );
            this.panel1.ResumeLayout( false );
            ( ( System.ComponentModel.ISupportInitialize ) ( this.ds����������������������1 ) ).EndInit( );
            this.ResumeLayout( false );

		}
		#endregion

		/// <summary>
		/// ������� Click ������ ������
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
		/// ������� Click ������ ������
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void button1_Click(object sender, System.EventArgs e)
		{
                                  
            this.Enabled = false;
            �����������������( );

            FormView frmPrint = new FormView( );
            try
            {
                ReportDocument rptDoc = new ReportDocument( );
                string fileName = @"..\report\Statistic2.rpt";
                // ��������� ���� ������:
                rptDoc.Load( fileName );
                // �������� ������:
                rptDoc.SetDataSource( ds����������������������1 );
                // ������������ �������� �������� ������:
                frmPrint.reportViewer.ReportSource = rptDoc;
                Form�������.������������������( "BeginDate", BeginDate, frmPrint.reportViewer.ParameterFieldInfo );
                Form�������.������������������( "EndDate", EndDate, frmPrint.reportViewer.ParameterFieldInfo );
                frmPrint.Text = "���������� �� ������������";
                // ���������� �����:
                frmPrint.reportViewer.ShowGroupTreeButton = false;
                this.Hide( );
                frmPrint.ShowDialog( this );
            }
            catch ( Exception exc )
            {
                MessageBox.Show( this, "��������� ������ ��� �������� ����� ������ \"���������� �� ������������\".\n" + exc.Message, "������ �������� ����� ������" );
                return;
            }
            finally
            {
                this.Enabled = true;
                BeginDate = null;
                EndDate = null;
                Dispose( true );
            }
		}

		private void FormDateRange_Load(object sender, System.EventArgs e)
		{
			dt1.Value = DateTime.Now.AddMonths(-1);//������ ��������� ������� �� 1 ����� ������ ����������� ����
			dt2.Value = DateTime.Now;// ����� ��������� ������ ����� ����������� ����
		}

        private void ����������������� ( )
        {          

            BeginDate = dt1.Value.ToShortDateString( );
            EndDate = dt2.Value.ToShortDateString( );
            // �������� ������ ������������
            ArrayList arr = new ArrayList();     
            foreach ( DataRow row in ds.����������.Rows )
            {
                if (!arr.Contains(row["������������������"].ToString()))
                {
                    if ( DBNull.Value == row["������"])
                    {
                        arr.Add(row["������������������"].ToString());
                    }
                }
            }     
            arr.Sort( );

            // ���� �� ������ ������������ � ������ ������� ���������� ��������-����������
            foreach ( Object ����������� in arr )
            {
                // �������� �����    
                int ��������������� = 0;
                string filter = "��������� LIKE '%" + �����������.ToString( ) + "%' AND ����������>='" + BeginDate + "' AND ����������<='" + EndDate + "'";
                DataRow [ ] rows�������� = ds.��������.Select(filter);
                if ( rows��������.Length > 0 )
                {
                    ��������������� = rows��������.Length;
                }

                // ���������� �����
                filter = "������������������='" + �����������.ToString( ) + "' AND (NOT ������ OR ISNULL(������, True))";
                DataRow [ ] rows = ds.����������.Select(filter);
                int id����������� = ( int ) rows [0] ["id_����������"];
                filter = "id_�������������������������=" + id����������� + " AND (NOT ������ OR ISNULL(������, True))";
                rows = ds.���������������������.Select( filter );
                int ����������������� = 0; 
                if ( rows.Length > 0 )
                {
                    string ������������������� = "";
                    for ( int i = 0; i < rows.Length; i++ )
                    {
                        ������������������� += rows [i] ["id_�������������"].ToString( );
                        if ( i < rows.Length - 1 )
                        {
                            ������������������� += ", ";
                        }
                    }
                    filter = "id_������������� IN (" + ������������������� + ") AND ����>='" + BeginDate + "' AND ����<='" + EndDate + "'";
                    rows = ds.�����������������.Select( filter );
                    ����������������� = rows.Length;

                }
                this.ds����������������������1.�����������.Add�����������Row( �����������.ToString( ), ���������������, ����������������� );
            } 
        }
	}
}
