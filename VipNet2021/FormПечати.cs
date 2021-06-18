using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace RegKor
{
	/// <summary>
	/// Summary description for FormView.
	/// </summary>
	public class FormView : System.Windows.Forms.Form
	{
		public CrystalDecisions.Windows.Forms.CrystalReportViewer reportViewer;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public FormView()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				this.reportViewer.Dispose();
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormView));
            this.reportViewer = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.SuspendLayout();
            // 
            // reportViewer
            // 
            this.reportViewer.ActiveViewIndex = -1;
            this.reportViewer.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.reportViewer.AutoScroll = true;
            this.reportViewer.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.reportViewer.Cursor = System.Windows.Forms.Cursors.Hand;
            this.reportViewer.DisplayGroupTree = false;
            this.reportViewer.EnableDrillDown = false;
            this.reportViewer.Location = new System.Drawing.Point(0, 0);
            this.reportViewer.Name = "reportViewer";
            this.reportViewer.SelectionFormula = "";
            this.reportViewer.ShowCloseButton = false;
            this.reportViewer.ShowGotoPageButton = false;
            this.reportViewer.ShowGroupTreeButton = false;
            this.reportViewer.ShowRefreshButton = false;
            this.reportViewer.ShowTextSearchButton = false;
            this.reportViewer.Size = new System.Drawing.Size(670, 456);
            this.reportViewer.TabIndex = 0;
            this.reportViewer.ViewTimeSelectionFormula = "";
            this.reportViewer.QueryContinueDrag += new System.Windows.Forms.QueryContinueDragEventHandler(this.reportViewer_QueryContinueDrag);
            this.reportViewer.Load += new System.EventHandler(this.crystalReportViewer1_Load);
            // 
            // FormView
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(672, 455);
            this.Controls.Add(this.reportViewer);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimizeBox = false;
            this.Name = "FormView";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Печать";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.ResumeLayout(false);

		}
		#endregion

		private void crystalReportViewer1_Load(object sender, System.EventArgs e)
		{
			reportViewer.Zoom(100);
		}

		private void reportDocument1_InitReport(object sender, System.EventArgs e)
		{
		
		}

        private void reportViewer_QueryContinueDrag(object sender, QueryContinueDragEventArgs e)
        {
            MessageBox.Show("");
        }
	}
}
