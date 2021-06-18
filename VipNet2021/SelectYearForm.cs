using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;

namespace RegKor
{
    public partial class SelectYearForm : Form
    {
        private int selectedYear;

        /// <summary>
        /// Хранит выбранный год
        /// </summary>
        public int SelectedYear
        {
            get
            {
                return selectedYear;
            }
            set
            {
                selectedYear = value;
            }
        }

        public SelectYearForm()
        {
            InitializeComponent();


        }

        private void SelectYearForm_Load(object sender, EventArgs e)
        {
            int iТекущийГод = Convert.ToInt32(ConfigurationSettings.AppSettings["выбранныйГод"]);
            int iпервыйГод = Convert.ToInt32(ConfigurationSettings.AppSettings["первыйГод"]);

            for (int i = iТекущийГод; i >= iпервыйГод; i--)
            {
                this.comboBox1.Items.Add(i);
            }
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            //если пользователь не выбрал год то кнопка остаётся закрытой
            if (this.comboBox1.Text != "")
            {
                this.button1.Enabled = true;
            }
            else
            {
                this.button1.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.Text != "")
            {
                this.SelectedYear = Convert.ToInt32(this.comboBox1.Text);
            }
        }
    }
}