using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using RegKor.Classess;

namespace RegKor
{
    public partial class FormMessage : Form
    {
        private string ��������������;

        private Item�������������������������� ������;

        private string ������������� = string.Empty;

        public bool flag���������VipNet = false;

        // ��������� ��� ����� ���������� ����� ��� ���������� ��������� VipNet
        public bool Flag���������VipNet
        {
            get
            {
                return flag���������VipNet;
            }
            set
            {
                flag���������VipNet = value;
            }
            
        }
                

        /// <summary>
        /// ������ id ��������.
        /// </summary>
        public string NumCardDoc
        {
            get
            {
                return �������������;
            }
            set
            {
                ������������� = value;
            }
        }

        /// <summary>
        /// �������� ������ ����������� ���������.
        /// </summary>
        public Item�������������������������� ��������������������������
        {
            get
            {
                return ������;
            }
            set
            {
                ������ = value;
            }
        }

        /// <summary>
        /// ����� ���������
        /// </summary>
        public string ��������������
        {
            get
            {
                return ��������������;
            }
            set
            {
                �������������� = value;
            }

        }

        public FormMessage(string numDoc)
        {
            InitializeComponent();

            this.label2.Text = numDoc.Trim();
            //this.label2.TextAlign = ContentAlignment.MiddleCenter;
        }

        private void dtnClose_Click(object sender, EventArgs e)
        {
            //if (������ != null)
            //{
            //    if (������.ProcessName.ToLower().Trim() == "ViPNet".ToLower().Trim() || ������.ProcessName.ToLower().Trim() == "e-mail".ToLower().Trim())
            //    {
            //        // ������� ���� � ����� ������ ������� ����� ������� ����� � ������� ����������.
            //        string patchDir = ConfigurationSettings.AppSettings["����������������������������"].Trim();

            //        // �������� ���������� ��� �������� ���������.
            //        string nameDir = this.��������������.Trim().Replace("/", "-") + "-id" + this.NumCardDoc.Trim();

            //        // ������� ���������� � �������� ��������.
            //        DirectoryInfo dirInfo = new DirectoryInfo(patchDir);

            //        // �������� ��������������.
            //        dirInfo.CreateSubdirectory(nameDir);

            //        // �������� ����� ������ ����� patchDir.
            //        //System.IO.Directory.CreateDirectory(patchDir + "/" + );

            //    }
            //}

            this.Close();
        }
    }
}