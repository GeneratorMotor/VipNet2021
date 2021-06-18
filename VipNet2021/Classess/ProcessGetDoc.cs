using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    public class ProcessGetDoc:IProcessGetDoc
    {
        #region IProcessGetDoc Members

        public List<string> GetDoc()
        {
            List<string> listDoc = new List<string>();

            string query = "SELECT [ВидПоступленияДокумента] " +
                           "FROM [ВидПоступленияДокумента] ";

            DataTable tabDoc = DataTableSql.GetDataTable(query);

            if (tabDoc.Rows.Count > 0)
            {
                foreach(DataRow row in tabDoc.Rows)
                {
                    listDoc.Add(row["ВидПоступленияДокумента"].ToString().Trim());
                }
            }

            return listDoc;
        }

        #endregion
    }
}
