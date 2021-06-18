namespace RegKor
{
    partial class FormКонтрольныеУведомления
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager( typeof( FormКонтрольныеУведомления ) );
            this.dsКонтрольныеУведомления1 = new RegKor.DSКонтрольныеУведомления( );
            this.reportViewer = new CrystalDecisions.Windows.Forms.CrystalReportViewer( );
            ( ( System.ComponentModel.ISupportInitialize ) ( this.dsКонтрольныеУведомления1 ) ).BeginInit( );
            this.SuspendLayout( );
            // 
            // dsКонтрольныеУведомления1
            // 
            this.dsКонтрольныеУведомления1.DataSetName = "DSКонтрольныеУведомления";
            this.dsКонтрольныеУведомления1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // reportViewer
            // 
            this.reportViewer.ActiveViewIndex = -1;
            this.reportViewer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.reportViewer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.reportViewer.EnableToolTips = false;
            this.reportViewer.Location = new System.Drawing.Point( 0, 0 );
            this.reportViewer.Name = "reportViewer";
            this.reportViewer.SelectionFormula = "";
            this.reportViewer.ShowCloseButton = false;
            this.reportViewer.ShowGotoPageButton = false;
            this.reportViewer.ShowTextSearchButton = false;
            this.reportViewer.Size = new System.Drawing.Size( 606, 406 );
            this.reportViewer.TabIndex = 2;
            this.reportViewer.ViewTimeSelectionFormula = "";
            // 
            // FormКонтрольныеУведомления
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF( 6F, 13F );
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size( 606, 406 );
            this.Controls.Add( this.reportViewer );
            this.Icon = ( ( System.Drawing.Icon ) ( resources.GetObject( "$this.Icon" ) ) );
            this.Name = "FormКонтрольныеУведомления";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Контрольные уведомления";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            ( ( System.ComponentModel.ISupportInitialize ) ( this.dsКонтрольныеУведомления1 ) ).EndInit( );
            this.ResumeLayout( false );

        }

        #endregion

        private DSКонтрольныеУведомления dsКонтрольныеУведомления1;
        public CrystalDecisions.Windows.Forms.CrystalReportViewer reportViewer;
    }
}