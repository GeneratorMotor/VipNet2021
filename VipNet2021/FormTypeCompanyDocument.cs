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
        private Item�������������������������� item;

        /// <summary>
        /// ���������� ��������� ����� ����������� ���������.
        /// </summary>
        public Item�������������������������� �����������������
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
            // ��������� ������� ������ ��������� ������ ��������� ���������.
            item = new Item��������������������������();

            ������������ connectString = new ������������();

            �����������������������DBContext context = new �����������������������DBContext(connectString.�����������������());

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
                            MessageBox.Show("�� �� ������� ������ ����������� ���������");
                            return;
                        }
                    }
                }
            }

            ����������������� = item;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}