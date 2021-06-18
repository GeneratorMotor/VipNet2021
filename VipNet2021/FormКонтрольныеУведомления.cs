using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RegKor
{
    public partial class FormКонтрольныеУведомления : Form
    {
        public FormКонтрольныеУведомления(DSКонтрольныеУведомления dataset)
        {
            InitializeComponent();
            this.dsКонтрольныеУведомления1 = dataset;
            //DataRow[] строкиПолучателей = ds11.Получатели.Select("", "ОписаниеПолучателя");
            //string[] массивПОлучателей = new string[строкиПолучателей.Length];
            //for(int i = 0; i < строкиПолучателей.Length; i++)
            //{
            //    массивПОлучателей[i] = (string)строкиПолучателей[i]["ОписаниеПолучателя"];
            //}
            //listBoxПолучатели.Items.AddRange(массивПОлучателей);
        }
    }
}