using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace RegKor.Classess
{
    class PersonDataDefault
    {
        private int id_����������������� = 0;
        public PersonDataDefault(int id��������)
        {
            id_����������������� = id��������;
        }

        /// <summary>
        /// ���������� ������ ��������� �������� ������ ������������� ������ �� ���������.
        /// </summary>
        /// <returns></returns>
        public List<�����������������> GetList()
        {
            string quer = "select * from ����������������� " +
                          "where id_����������������� in ( " +
                          "SELECT [id_�����������������] " +
                          "FROM [���������������������������������������] " +
                          "where id_�������� = " + id_����������������� + " )";

            GetDataTable getTabC = new GetDataTable(quer);
            DataTable tabContr = getTabC.DataTable("����������������������������");

            List<�����������������> list = new List<�����������������>();
            foreach (DataRow row in tabContr.Rows)
            {
                ����������������� item = new �����������������();
                item.Id_����������������� = Convert.ToInt32(row["id_�����������������"]);
                item.��������� = row["�����������������"].ToString().Trim();
                item.FlagSelect = true;
               
                list.Add(item);
            }

            return list;
        }
                
    }
}
