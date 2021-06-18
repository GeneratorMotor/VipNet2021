using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    /// <summary>
    /// ����� ������������ �����.
    /// </summary>
    public  class GenerateStaticDocInput
    {
        private List<StatisticDocInput> list;
        private StatisticDocInput itemReportCount;

        public GenerateStaticDocInput(List<StatisticDocInput> listReport)
        {
            list = listReport;
            itemReportCount = new StatisticDocInput();
        }

        public StatisticDocInput ItemCount
        {
            get
            {
                return itemReportCount;
            }

        }

        public void Generate(DateTime dataStart, DateTime dataEnd, string fio)
        {
            DataTable dTab = new DataTable();
            ������������ strConnect = new ������������();

            StatisticDocInput itemCount = new StatisticDocInput();
            
            // ������� ���������� �������� �����.
            using (SqlConnection con = new SqlConnection(strConnect.�����������������()))
            {
                SqlCommand com = new SqlCommand("ReportStatInputLetter", con);
                com.CommandType = CommandType.StoredProcedure;

                com.Parameters.Add(new SqlParameter("@dateStart", SqlDbType.DateTime));
                com.Parameters["@dateStart"].Value = dataStart;// ����SQL.����(this.�����������.DataStart.ToShortDateString());

                com.Parameters.Add(new SqlParameter("@dateEnd", SqlDbType.DateTime));
                com.Parameters["@dateEnd"].Value = dataEnd;// ����SQL.����(this.�����������.DataEnd.ToShortDateString());

                com.Parameters.Add(new SqlParameter("@fio", SqlDbType.VarChar,100));
                com.Parameters["@fio"].Value = fio;

                SqlDataAdapter da = new SqlDataAdapter(com);
                da.Fill(dTab);
            }

            int iNum = 1;
            string person = string.Empty;

            // �������� ��������� ������� ����������� ����� �������.
            foreach (DataRow row in dTab.Rows)
            {
                StatisticDocInput item = new StatisticDocInput();


                item.Num = iNum.ToString();
                item.�������������������������� = row["����������������������"].ToString().Trim();
                item.������������������������ = Convert.ToInt32(row["����������������������������"]);
                itemCount.������������������������ += Convert.ToInt32(row["����������������������������"]);

                if (row["����������������"] != DBNull.Value)
                {
                    item.���������������� = Convert.ToInt32(row["����������������"]);
                    itemCount.���������������� += Convert.ToInt32(row["����������������"]);
                }
                else
                {
                    item.���������������� = null;
                }

                if (row["e-mail"] != DBNull.Value)
                {
                    item.Email = Convert.ToInt32(row["e-mail"]);
                    itemCount.Email += Convert.ToInt32(row["e-mail"]);
                }
                else
                {
                    item.Email = null;
                }

                if (row["VipNet"] != DBNull.Value)
                {
                    item.VipNet = Convert.ToInt32(row["VipNet"]);
                    itemCount.VipNet += Convert.ToInt32(row["VipNet"]);
                }
                else
                {
                    item.VipNet = null;
                }

                if (row["Fax"] != DBNull.Value)
                {
                    item.Fax = Convert.ToInt32(row["Fax"]);
                    itemCount.Fax += Convert.ToInt32(row["Fax"]);
                }
                else
                {
                    item.Fax = null;
                }
                item.����������� = fio.Trim();
                itemCount.����������� = fio.Trim();

                list.Add(item);

                iNum++;

                //itemReportCount.���������������� += itemCount.����������������;
                //itemReportCount.������������������������ += itemCount.������������������������;
                //itemReportCount.VipNet += itemCount.VipNet;
                //itemReportCount.Fax += itemCount.Fax;
                //itemReportCount.Email += itemCount.Email;

            }

            itemCount.Num = "����� �� ����������� " + itemCount.�����������.Trim();

            itemReportCount = itemCount;


            list.Add(itemCount);
        }

    }
}
