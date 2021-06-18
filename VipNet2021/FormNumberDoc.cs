using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RegKor
{
    public partial class FormNumberDoc : Form
    {

        private string nummberDoc = string.Empty;

        /// <summary>
        /// Номер тдокумента.
        /// </summary>
        public string NumberDoc
        {
            get
            {
                return nummberDoc;
            }
            set
            {
                nummberDoc = value;
            }
        }

        public FormNumberDoc()
        {
            InitializeComponent();
        }

               

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.NumberDoc = this.txtNumberDoc.Text.Trim();
        }

        private void FormNumberDoc_Load(object sender, EventArgs e)
        {
            this.txtNumberDoc.Text = this.NumberDoc;
        }
    }
}