using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RegKor
{
    public partial class FormSelectMonth : Form
    {
        string ������������;
        string ���;
        string �������;
        string ���������������;

        /// <summary>
        ///  ���������� ������ ���� ���������� ���������
        /// </summary>
        public string Get����������
        {
            get
            {
                return ������������;
            }
            set
            {
                ������������ = value;
            }
        }

        /// <summary>
        /// ���������� ��������� ���� ������.
        /// </summary>
        public string Get�����������
        {
            get
            {
                return ���������������;
            }
            set
            {
                ��������������� = value;
            }
        }

        /// <summary>
        /// ���� ���.
        /// </summary>
        public string �������
        {
            get
            {
                return �������;
            }
            set
            {
                ������� = value;
            }
        }

        /// <summary>
        /// ���������� ���.
        /// </summary>
        public string ������������
        {
            get
            {
                return ���;
            }
            set
            {
                ��� = value;
            }
        }

        public FormSelectMonth()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.Text != "")
            {

                // ������� ���.
                int _��� = Convert.ToInt32(this.������������);

                // ������� ����� ������.
                string �������������� = this.comboBox1.Text;
                string ������ = string.Empty;

                switch (��������������)
                {
                    case "������":
                        ������ = "01";
                        break;
                    case "�������":
                        ������ = "02";
                        break;
                    case "����":
                        ������ = "03";
                        break;
                    case "������":
                        ������ = "04";
                        break;
                    case "���":
                        ������ = "05";
                        break;
                    case "����":
                        ������ = "06";
                        break;
                    case "����":
                        ������ = "07";
                        break;
                    case "������":
                        ������ = "08";
                        break;
                    case "��������":
                        ������ = "09";
                        break;
                    case "�������":
                        ������ = "10";
                        break;
                    case "������":
                        ������ = "11";
                        break;
                    case "�������":
                        ������ = "12";
                        break;
                }

                if (�������������� != "���� ���")
                {

                    // ������� ���������� ���� � ������.
                    int �������������� = Convert.ToInt32(������);

                    int num��� = _��� + 1;

                    int countDay = DateTime.DaysInMonth(num���, ��������������);

                    // ��������� ���� ������ ������.
                    this.Get���������� = "01." + ������ + "." + num���.ToString().Trim();

                    this.Get����������� = countDay.ToString().Trim() + "." + ������ + "." + num���.ToString().Trim();
                }
                else
                {
                    int num��� = _��� + 1;

                    this.Get���������� = "01.01." + num���.ToString().Trim();

                    this.Get����������� = "31.12." + num���.ToString().Trim();

                }
            }
            else
            {
                MessageBox.Show("�� ������ �����");
                this.Close();
            }

        }
    }
}