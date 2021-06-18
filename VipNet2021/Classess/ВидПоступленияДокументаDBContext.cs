using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;


namespace RegKor.Classess
{
    class ВидПоступленияДокументаDBContext : DataBaseContext
    {
        private string connectionString = string.Empty;

        public ВидПоступленияДокументаDBContext(string sConnect)
        {
            if (sConnect.Length > 0)
            {
                connectionString = sConnect;
            }
        }

        /// <summary>
        /// Добавляет запись в таблицу ВидПоступленияДокумента.
        /// </summary>
        /// <param name="value"></param>
        public override void Insert(string value)
        {
            string query = "INSERT INTO [ВидПоступленияДокумента] ([ВидПоступленияДокумента]) " +
                           "VALUES " +
                           "("+ value +") ";
            using(SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                SqlCommand com = new SqlCommand(query, con);
                com.ExecuteNonQuery();
            }
        }

        public override void Update(int idKey)
        {
            
        }

        public override void Delete(int idKey)
        {
            //System.Windows.Forms.MessageBox.Show("Удалить вид документа");
        }

        private void ExecuteQuery()
        {

        }

        public List<ItemСпособПоступленияДокумента> Select(string nameItem)
        {

            List<ItemСпособПоступленияДокумента> list = new List<ItemСпособПоступленияДокумента>();

            string query = "SELECT  [id] " +
                           ",[ВидПоступленияДокумента] " +
                          "FROM [ВидПоступленияДокумента] where LOWER(LTRIM(RTRIM(ВидПоступленияДокумента))) = '"+ nameItem.Trim().ToLower() +"' ";

            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                SqlCommand com = new SqlCommand(query, con);

                SqlDataReader reader = com.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        ItemСпособПоступленияДокумента item = new ItemСпособПоступленияДокумента();
                        item.Id = Convert.ToInt32(reader["id"]);
                        item.ProcessName = reader["ВидПоступленияДокумента"].ToString().Trim();

                        list.Add(item);
                    }
                }
               
                reader.Close();
            }

            return list;
        }

    }
}
