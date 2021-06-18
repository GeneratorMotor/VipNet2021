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
    public partial class FormСписокСотрудников : Form
    {
        public int ИДСотрудника = -1;
        DS1TableAdapters.ПолучателиTableAdapter получателиTableAdapter = new RegKor.DS1TableAdapters.ПолучателиTableAdapter( );
        private System.Data.DataView viewСотрудники;

        public FormСписокСотрудников ( )
        {
            InitializeComponent( );


            string query = " select id_получателя,ОписаниеПолучателя from dbo.Получатели " +
                          " where Удален is null ";

            DataTable tabPers = DataTableSql.GetDataTable(query);

            //получателиTableAdapter.Fill( ds11.Получатели );
            //this.viewСотрудники = new System.Data.DataView( ds11.Получатели);

            this.viewСотрудники = new System.Data.DataView(tabPers);

            this.viewСотрудники.Sort = "ОписаниеПолучателя";
            listBoxСотрудники.DataSource = viewСотрудники;
            listBoxСотрудники.DisplayMember = "ОписаниеПолучателя";
            listBoxСотрудники.ValueMember = "id_получателя";
        }

        private void buttonОтмена_Click ( object sender, EventArgs e )
        {
            this.Close( );
        }

        private void textBoxПоиск_TextChanged ( object sender, EventArgs e )
        {
            this.viewСотрудники.RowFilter = "ОписаниеПолучателя LIKE '%" + textBoxПоиск.Text + "%'";
            this.viewСотрудники.Sort = "ОписаниеПолучателя";
            listBoxСотрудники.DataSource = viewСотрудники;
            listBoxСотрудники.DisplayMember = "ОписаниеПолучателя";
            listBoxСотрудники.ValueMember = "id_получателя";
        }

        private void buttonСохранить_Click ( object sender, EventArgs e )
        {

            ИДСотрудника = ( int ) listBoxСотрудники.SelectedValue;
            Close( );
        }

        private void listBoxСотрудники_DoubleClick ( object sender, EventArgs e )
        {
            if ( listBoxСотрудники.SelectedItem != null )
            {
                ИДСотрудника = ( int ) listBoxСотрудники.SelectedValue;
                this.DialogResult = DialogResult.OK;
                Close( );
            }
        }

    }
}