using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Windows.Forms;

namespace RegKor.Classess
{
    public class ОтчетОВходДокументах:IОтчет
    {
        private DataTable tab;

        /// <summary>
        /// Таблица с данными.
        /// </summary>
        public DataTable TableData
        {
            get
            {
                return tab;
            }
            set
            {
                tab = value;
            }
        }


        // DataGridView
        private DataGridView dataGridView1;

        /// <summary>
        /// Дата GridView.
        /// </summary>
        public DataGridView DataGridView1
        {
            get
            {
                return dataGridView1;
            }
            set
            {
                dataGridView1 = value;
            }
        }
    }
}
