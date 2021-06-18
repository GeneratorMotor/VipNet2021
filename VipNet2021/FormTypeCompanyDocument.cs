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
    public partial class FormTypeCompanyDocument : Form
    {
        private ItemСпособПоступленияДокумента item;

        /// <summary>
        /// Возвращает выбранный спосб поступления документа.
        /// </summary>
        public ItemСпособПоступленияДокумента СпособПоступления
        {
            get
            {
                return item;
            }
            set
            {
                item = value;
            }
        }
            
            

        public FormTypeCompanyDocument()
        {
            InitializeComponent();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            // Экземпляр который хранит выбранный способ получения документа.
            item = new ItemСпособПоступленияДокумента();

            ПодключитьБД connectString = new ПодключитьБД();

            ВидПоступленияДокументаDBContext context = new ВидПоступленияДокументаDBContext(connectString.СтрокаПодключения());

            foreach (Control contrl in this.Controls)
            {
                if (contrl is RadioButton)
                {
                    RadioButton rb = (RadioButton)contrl;

                    if (rb.Checked == true)
                    {
                        if (context.Select(rb.Text.Trim()).Count > 0)
                        {
                            item = context.Select(rb.Text.Trim())[0];
                        }
                        else
                        {
                            MessageBox.Show("Вы не выбрали способ поступления документа");
                            return;
                        }
                    }
                }
            }

            СпособПоступления = item;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}