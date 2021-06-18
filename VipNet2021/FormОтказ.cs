using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RegKor
{
    public partial class FormОтказ : Form
    {
        private string text = string.Empty;

        /// <summary>
        /// Хранит текст отказа.
        /// </summary>
        public string ТекстОтказа
        {
            get
            {
                return text;
            }
            set
            {
                text = value;
            }
        }

        public FormОтказ()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            // Передадим в свойстов текст отказа.
            ТекстОтказа = this.textBox1.Text.Trim();

            // Закроем форму.
            this.Close();
        }

        
    }
}