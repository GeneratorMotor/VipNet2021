using System;
using System.Collections.Generic;
using System.Text;

namespace RegKor.Classess
{
    public abstract class DataBaseContext
    {
        //private string connectString = string.Empty;

        //public DataBaseContext(string connectionString)
        //{
        //    if(connectionString.Length > 0)
        //    connectString = connectionString;
        //}

        public abstract void Insert(string value);

        public abstract void Update(int idKey);

        public abstract void Delete(int idKey);
    }
}
