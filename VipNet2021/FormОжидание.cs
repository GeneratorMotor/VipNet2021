using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RegKor
{
    public partial class FormОжидание : Form
    {
        public FormОжидание()
        {
            InitializeComponent();
            this.timer1.Interval = 10;
        }

        private void FormОжидание_Shown(object sender, EventArgs e)
        {
            this.timer1.Enabled = true;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Refresh();
        }
    }
}