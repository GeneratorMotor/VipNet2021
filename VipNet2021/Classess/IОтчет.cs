using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Windows.Forms;

namespace RegKor.Classess
{
    public class IОтчет
    {
        DataTable tab;

        /// <summary>
        /// Таблица с данными.
        /// </summary>
        DataTable TableData
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
        DataGridView dataGridView1;

        /// <summary>
        /// Дата GridView.
        /// </summary>
        DataGridView DataGridView1
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
