using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Configuration;
using System.Configuration.Assemblies;

namespace RegKor
{
    public partial class FormГодКонфигурации : Form
    {
        public FormГодКонфигурации()
        {
            InitializeComponent();
            
            //string год = ConfigurationSettings.AppSettings["ГодДокумента"].ToString();
            RegKor.Classess.ГодДокументооборота годДокумента = new RegKor.Classess.ГодДокументооборота();
            bool режимРаботы = годДокумента.ГодВКонфигурационномФайле();


            if (режимРаботы == true)
            {
                this.flagГод.Checked = true;
                //this.radioButton2.Checked = true;
                //this.radioButton1.Checked = false;
            }
            if (режимРаботы == false)
            {
                this.flagГод.Checked = false;
                //this.radioButton1.Checked = true;
                //this.radioButton2.Checked = false;
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Создаём файл записываем в него булево значение с сериализуем

            bool версияБазы;
            if (flagГод.Checked == true)
            {
                версияБазы = true;
                Classess.ВерсияРаботы версияРаботыПрограммы = new RegKor.Classess.ВерсияРаботы();
                версияРаботыПрограммы.ВерсияБазы(версияБазы);
                this.Close();
            }
            else
            {
                версияБазы = false;
                Classess.ВерсияРаботы версияРаботыПрограммы = new RegKor.Classess.ВерсияРаботы();
                версияРаботыПрограммы.ВерсияБазы(версияБазы);
                this.Close();
            }
        }
    }
}