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
        private string номерДокумента;

        private ItemСпособПоступленияДокумента способ;

        private string номерКарточки = string.Empty;

        public bool flagИсходящийVipNet = false;

        // Указывает что форма гинерирует папку для исходящего документа VipNet
        public bool FlagИсходящийVipNet
        {
            get
            {
                return flagИсходящийVipNet;
            }
            set
            {
                flagИсходящийVipNet = value;
            }
            
        }
                

        /// <summary>
        /// Хранит id карточки.
        /// </summary>
        public string NumCardDoc
        {
            get
            {
                return номерКарточки;
            }
            set
            {
                номерКарточки = value;
            }
        }

        /// <summary>
        /// Получает способ поступления документа.
        /// </summary>
        public ItemСпособПоступленияДокумента СпособПоступленияДокумента
        {
            get
            {
                return способ;
            }
            set
            {
                способ = value;
            }
        }

        /// <summary>
        /// Номер документа
        /// </summary>
        public string НомерДокумента
        {
            get
            {
                return номерДокумента;
            }
            set
            {
                номерДокумента = value;
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
            //if (способ != null)
            //{
            //    if (способ.ProcessName.ToLower().Trim() == "ViPNet".ToLower().Trim() || способ.ProcessName.ToLower().Trim() == "e-mail".ToLower().Trim())
            //    {
            //        // Получим путь к папке внутри которой нужно создать папку с номером документов.
            //        string patchDir = ConfigurationSettings.AppSettings["локальнаПапкаДокументооборот"].Trim();

            //        // Название директории для хранения документа.
            //        string nameDir = this.НомерДокумента.Trim().Replace("/", "-") + "-id" + this.NumCardDoc.Trim();

            //        // Получим информацию о каталоге хранения.
            //        DirectoryInfo dirInfo = new DirectoryInfo(patchDir);

            //        // Создадим поддирректорию.
            //        dirInfo.CreateSubdirectory(nameDir);

            //        // Создадим папку внутри папки patchDir.
            //        //System.IO.Directory.CreateDirectory(patchDir + "/" + );

            //    }
            //}

            this.Close();
        }
    }
}