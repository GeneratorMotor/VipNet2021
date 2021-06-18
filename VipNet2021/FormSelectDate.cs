using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;

namespace RegKor
{
    public partial class FormSelectDate : Form
    {

        private RangeDate _rd;

        /// <summary>
        /// Диапазон дат.
        /// </summary>
        public RangeDate ДиапазоДат
        {
            get
            {
                return _rd;
            }
            set
            {
                _rd = value;
            }
            

        }

        public FormSelectDate()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            RangeDate rd = new RangeDate();
            rd.DataStart = this.dtStart.Value.Date;
            rd.DataEnd = this.dtEnd.Value.Date;

            this.ДиапазоДат = rd;

            this.Close();
        }
    }
}