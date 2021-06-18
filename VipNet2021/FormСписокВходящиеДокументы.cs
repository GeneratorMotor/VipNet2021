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
    public partial class FormСписокВходящиеДокументы : Form
    {
        public int ИДВходящегоДокумента = -1;
        private RegKor.DS1.ВыборкаDataTable tableДокументы;
        private System.Data.DataView viewДокументы;
        public FormСписокВходящиеДокументы ( RegKor.DS1.ВыборкаDataTable table )
        {
            InitializeComponent( );
            tableДокументы = new DS1.ВыборкаDataTable( );
            this.tableДокументы = table;
            this.viewДокументы = new System.Data.DataView(tableДокументы, "ВДело=False AND ДатаПоступ >='01.11.2011' AND НомерВход<>''", "НомерВход", DataViewRowState.CurrentRows);
            
            // Получим данные 
            string date = DateTime.Now.ToShortDateString();
            string query = "select НомерВход,id_карточкиВходящей as 'id_карточки' from ВыборкаПовтор " +
                           "where FlagControl = 'True' and СрокВыполнения < '"+ ДатаSQL.Дата(date) +"'";

            GetDataTable tab = new GetDataTable(query);
            DataTable tabПовтор = tab.DataTable("КарточкаПовтор");

            List<ItemDoc> list = new List<ItemDoc>();

            // Пробежимся по первой таблице и запишем в неё все номера
            foreach (DataRowView row in viewДокументы)
            {
                ItemDoc item = new ItemDoc();
                item.id_карточки = Convert.ToInt32(row["id_карточки"]);
                item.НомерВход = row["НомерВход"].ToString().Trim();

                list.Add(item);
            }

            //int i1 = list.Count;

            // Теперь пробежимся по второй таблице.
            foreach (DataRow row in tabПовтор.Rows)
            {
                ItemDoc item = new ItemDoc();
                item.id_карточки = Convert.ToInt32(row["id_карточки"]);
                item.НомерВход = row["НомерВход"].ToString().Trim();

                list.Add(item);
            }

            //int i2 = list.Count;
            
            //listBoxДокументы.DataSource = viewДокументы;
            listBoxДокументы.DataSource = list;
            listBoxДокументы.DisplayMember = "НомерВход";
            listBoxДокументы.ValueMember = "id_карточки";
        }

        private void buttonОтмена_Click ( object sender, EventArgs e )
        {
            Close( );
        }

        private void textBoxПоиск_TextChanged ( object sender, EventArgs e )
        {
            this.viewДокументы.RowFilter = "ВДело=False AND НомерВход<>'' AND НомерВход LIKE '%" + textBoxПоиск.Text + "%'";

            // Получим данные 
            string date = DateTime.Now.ToShortDateString();
            string query = "select НомерВход,id_карточкиВходящей as 'id_карточки' from ВыборкаПовтор " +
                           "where FlagControl = 'True' and СрокВыполнения < '" + ДатаSQL.Дата(date) + "' AND НомерВход LIKE '%" + textBoxПоиск.Text + "%'";

            GetDataTable tab = new GetDataTable(query);
            DataTable tabПовтор = tab.DataTable("КарточкаПовтор");

            List<ItemDoc> list = new List<ItemDoc>();

            // Пробежимся по первой таблице и запишем в неё все номера
            foreach (DataRowView row in viewДокументы)
            {
                ItemDoc item = new ItemDoc();
                item.id_карточки = Convert.ToInt32(row["id_карточки"]);
                item.НомерВход = row["НомерВход"].ToString().Trim();

                list.Add(item);
            }

            int i1 = list.Count;

            // Теперь пробежимся по второй таблице.
            foreach (DataRow row in tabПовтор.Rows)
            {
                ItemDoc item = new ItemDoc();
                item.id_карточки = Convert.ToInt32(row["id_карточки"]);
                item.НомерВход = row["НомерВход"].ToString().Trim();

                list.Add(item);
            }

            int i2 = list.Count;

            //listBoxДокументы.DataSource = viewДокументы;
            listBoxДокументы.DataSource = list;
            listBoxДокументы.DisplayMember = "НомерВход";
            listBoxДокументы.ValueMember = "id_карточки";
        }

        private void listBoxДокументы_MouseDown ( object sender, MouseEventArgs e )
        {
            if ( listBoxДокументы.SelectedItem != null )
            {
                // получаем данные отображаемые в выделенной строке:
                BindingManagerBase bmb = this.BindingContext [viewДокументы];
                bmb.Position = listBoxДокументы.SelectedIndex;
                DataRowView drv = ( DataRowView ) bmb.Current;
                // выводим полученные данные на информационный лэйбл:РезультатВыполнения
                string подсказка = "Документ: " + drv ["ОписаниеДокумента"].ToString( ) + Environment.NewLine +
                    "Дата пост.: " + Convert.ToDateTime( drv ["ДатаПоступ"] ).ToShortDateString( ) + Environment.NewLine +
                    "Корреспондент: " + drv ["ОписаниеКорреспондента"].ToString( ) + Environment.NewLine +
                    "Содержание: " + drv ["КраткоеСодержание"].ToString( );

                toolTip1.UseAnimation = true;
                toolTip1.Show( подсказка, listBoxДокументы, e.X + 15, e.Y + 15, 5000 );
            }
        }

        private void listBoxДокументы_Leave ( object sender, EventArgs e )
        {
            if ( toolTip1.Active )
            {
                toolTip1.Hide( listBoxДокументы );
            }
        }

        private void buttonСохранить_Click ( object sender, EventArgs e )
        {
            ИДВходящегоДокумента = (int)listBoxДокументы.SelectedValue;
            Close( );
        }

        private void listBoxДокументы_MouseDoubleClick ( object sender, MouseEventArgs e )
        {
            if ( listBoxДокументы.SelectedItem != null )
            {
                ИДВходящегоДокумента = (int)listBoxДокументы.SelectedValue;
                this.DialogResult = DialogResult.OK;
                Close( );
            }
        }
    }
}