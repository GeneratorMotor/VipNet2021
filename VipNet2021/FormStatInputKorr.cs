using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace RegKor
{
    public partial class FormStatInputKorr : Form
    {
        private RangeDate rd;

        private DataTable rez;

        private List<StatisticDocInput> listCount;

        private List<StatisticDocInput> listPrint;

        /// <summary>
        /// ������ �������� ���.
        /// </summary>
        public RangeDate �����������
        {
            get
            {
                return rd;
            }
            set
            {
                rd = value;
            }
        }

        public FormStatInputKorr()
        {
            InitializeComponent();

            listCount = new List<StatisticDocInput>();

            listPrint = new List<StatisticDocInput>();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormStatInputKorr_Load(object sender, EventArgs e)
        {
            List<StatisticDocInput> list = SelectCorr();

            // ������ �������� �������� �����.
            StatisticDocInput count = new StatisticDocInput();

            foreach (StatisticDocInput it in this.listCount)
            {
                count.Num = "����� � ����� �� ��������.";
                count.������������������������ += it.������������������������;
                count.���������������� += it.����������������;
                count.VipNet += it.VipNet;
                count.Email += it.Email;
                count.Fax += it.Fax;
            }

            list.Add(count);

            this.listPrint = list;

            this.dataGridView1.DataSource = list;

            //this.dataGridView1.Columns["����"].HeaderText = "e-mail";
        }

        /// <summary>
        /// ���������� ���������� ��������� �� ����������� ������.
        /// </summary>
        /// <returns></returns>
        private List<StatisticDocInput> SelectCorr()
        {
            DataTable dTab��������� = new DataTable();

            ������������ strConnect = new ������������();

            List<StatisticDocInput> list = new List<StatisticDocInput>();

            using(SqlConnection con = new SqlConnection(strConnect.�����������������()))
            {

                // ������� ���� ����������� �������.
                SqlCommand com1 = new SqlCommand("ReportStatInputLetterPerson", con);
                com1.CommandType = CommandType.StoredProcedure;
                com1.Parameters.Add(new SqlParameter("@DataStart", SqlDbType.DateTime));
                com1.Parameters["@DataStart"].Value = this.�����������.DataStart;// ����SQL.����(this.�����������.DataStart.ToShortDateString());
                com1.Parameters.Add(new SqlParameter("@DateEnd", SqlDbType.DateTime));
                com1.Parameters["@DateEnd"].Value = this.�����������.DataEnd;// ����SQL.����(this.�����������.DataEnd.ToShortDateString());

                SqlDataAdapter da = new SqlDataAdapter(com1);
                da.Fill(dTab���������);

                foreach (DataRow r in dTab���������.Rows)
                {
                    GenerateStaticDocInput generatorReport = new GenerateStaticDocInput(list);
                    generatorReport.Generate(this.�����������.DataStart, this.�����������.DataEnd, r["������������������"].ToString().Trim());

                    StatisticDocInput itmCount = new StatisticDocInput();
                    itmCount = generatorReport.ItemCount;

                    listCount.Add(itmCount);

                }

               
            }

            #region
            //string query = "SELECT     derivedtbl_1.id_��������������, derivedtbl_1.����������������������, derivedtbl_1.������������������������, " +
            //              "derivedtbl_1.������������������, derivedtbl_2.������������������������ AS ������, derivedtbl_3.������������������������ AS ����, " +
            //              "derivedtbl_4.������������������������ AS ���, derivedtbl_5.������������������������ AS ���� " +
            //              " FROM         (SELECT     dbo.��������.id_��������������, dbo.��������������.����������������������, COUNT(dbo.����������.������������������) " +
            //              " AS ������������������������, dbo.����������.������������������ " +
            //          "FROM          dbo.�������� LEFT OUTER JOIN " +
            //                                 "dbo.�������������� ON dbo.��������.id_�������������� = dbo.��������������.id_�������������� LEFT OUTER JOIN " +
            //                                 "dbo.����������������������������� ON dbo.�����������������������������.id�������� = dbo.��������.id_�������� INNER JOIN " +
            //                                 "dbo.���������� ON dbo.����������.id_���������� = dbo.�����������������������������.id���������� " +
            //          "WHERE      (dbo.��������.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and dbo.��������.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
            //          "GROUP BY dbo.��������������.����������������������, dbo.��������.id_��������������, dbo.����������.������������������)  " +
            //         "AS derivedtbl_1 LEFT OUTER JOIN " +
            //             "(SELECT     ��������_4.id_��������������, ��������������_4.����������������������, COUNT(����������_4.������������������) " +
            //                                      "AS ������������������������, ����������_4.������������������ " +
            //               "FROM          dbo.�������� AS ��������_4 LEFT OUTER JOIN " +
            //                                      "dbo.�������������� AS ��������������_4 ON  " +
            //                                      "��������_4.id_�������������� = ��������������_4.id_�������������� LEFT OUTER JOIN " +
            //                                      "dbo.����������������������������� AS �����������������������������_4 ON  " +
            //                                      "�����������������������������_4.id�������� = ��������_4.id_�������� INNER JOIN " +
            //                                      "dbo.���������� AS ����������_4 ON ����������_4.id_���������� = �����������������������������_4.id���������� " +
            //               "WHERE      (��������_4.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_4.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
            //               "GROUP BY ��������������_4.����������������������, ��������_4.id_��������������, ����������_4.������������������, " +
            //                                      " ��������_4.id����������������������� " +
            //               "HAVING      (��������_4.id����������������������� = 4)) AS derivedtbl_5 ON  " +
            //         " derivedtbl_1.id_�������������� = derivedtbl_5.id_�������������� AND " +
            //         " derivedtbl_1.������������������ = derivedtbl_5.������������������ LEFT OUTER JOIN " +
            //             " (SELECT     ��������_3.id_��������������, ��������������_3.����������������������, COUNT(����������_3.������������������) " +
            //                                      " AS ������������������������, ����������_3.������������������ " +
            //               " FROM          dbo.�������� AS ��������_3 LEFT OUTER JOIN " +
            //                                      " dbo.�������������� AS ��������������_3 ON  " +
            //                                      " ��������_3.id_�������������� = ��������������_3.id_�������������� LEFT OUTER JOIN " +
            //                                      " dbo.����������������������������� AS �����������������������������_3 ON  " +
            //                                      " �����������������������������_3.id�������� = ��������_3.id_�������� INNER JOIN " +
            //                                      " dbo.���������� AS ����������_3 ON ����������_3.id_���������� = �����������������������������_3.id���������� " +
            //               "WHERE      (��������_3.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_3.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
            //               " GROUP BY ��������������_3.����������������������, ��������_3.id_��������������, ����������_3.������������������, " +
            //                                     " ��������_3.id����������������������� " +
            //               " HAVING      (��������_3.id����������������������� = 3)) AS derivedtbl_4 ON  " +
            //         " derivedtbl_1.id_�������������� = derivedtbl_4.id_�������������� AND " +
            //         " derivedtbl_1.������������������ = derivedtbl_4.������������������ LEFT OUTER JOIN " +
            //             " (SELECT     ��������_2.id_��������������, ��������������_2.����������������������, COUNT(����������_2.������������������) " +
            //                                      " AS ������������������������, ����������_2.������������������ " +
            //               " FROM          dbo.�������� AS ��������_2 LEFT OUTER JOIN " +
            //                                      " dbo.�������������� AS ��������������_2 ON " +
            //                                      " ��������_2.id_�������������� = ��������������_2.id_�������������� LEFT OUTER JOIN " +
            //                                      " dbo.����������������������������� AS �����������������������������_2 ON " +
            //                                      " �����������������������������_2.id�������� = ��������_2.id_�������� INNER JOIN " +
            //                                      " dbo.���������� AS ����������_2 ON ����������_2.id_���������� = �����������������������������_2.id����������" +
            //               " WHERE      (��������_2.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_2.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
            //              " GROUP BY ��������������_2.����������������������, ��������_2.id_��������������, ����������_2.������������������, " +
            //                                      " ��������_2.id����������������������� " +
            //               " HAVING      (��������_2.id����������������������� = 2)) AS derivedtbl_3 ON  " +
            //         " derivedtbl_1.id_�������������� = derivedtbl_3.id_�������������� AND " +
            //         "derivedtbl_1.������������������ = derivedtbl_3.������������������ LEFT OUTER JOIN " +
            //             " (SELECT     ��������_1.id_��������������, ��������������_1.����������������������, COUNT(����������_1.������������������)  " +
            //                                      " AS ������������������������, ����������_1.������������������ " +
            //               "FROM          dbo.�������� AS ��������_1 LEFT OUTER JOIN " +
            //                                      " dbo.�������������� AS ��������������_1 ON " +
            //                                      "��������_1.id_�������������� = ��������������_1.id_�������������� LEFT OUTER JOIN " +
            //                                      "dbo.����������������������������� AS �����������������������������_1 ON  " +
            //                                      " �����������������������������_1.id�������� = ��������_1.id_�������� INNER JOIN " +
            //                                     " dbo.���������� AS ����������_1 ON ����������_1.id_���������� = �����������������������������_1.id���������� " +
            //               " WHERE      (��������_1.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_1.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "' )" +
            //               " GROUP BY ��������������_1.����������������������, ��������_1.id_��������������, ����������_1.������������������,  " +
            //                                      " ��������_1.id����������������������� " +
            //               " HAVING      (��������_1.id����������������������� = 1)) AS derivedtbl_2 ON  " +
            //         " derivedtbl_1.������������������ = derivedtbl_2.������������������ AND derivedtbl_1.id_�������������� = derivedtbl_2.id_�������������� " +
            //         " order by derivedtbl_1.������������������ asc ";
            #endregion

            //GetDataTable getData = new GetDataTable(query);
            //return getData.DataTable();

            return list;
        }


        /// <summary>
        /// ���������� ������������ ������� ���� �������� ��������� �� ��������� ������.
        /// </summary>
        /// <returns></returns>
        private DataTable SelectPerson()
        {
            DataTable dTab = new DataTable();

            ������������ strConnect = new ������������();

            using (SqlConnection con = new SqlConnection(strConnect.�����������������()))
            {
                SqlCommand com = new SqlCommand("ReportStatInputLetterPerson", con);
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.Add(new SqlParameter("@DataStart", SqlDbType.DateTime));
                com.Parameters["@DataStart"].Value = this.�����������.DataStart;// ����SQL.����(this.�����������.DataStart.ToShortDateString());
                com.Parameters.Add(new SqlParameter("@DateEnd", SqlDbType.DateTime));
                com.Parameters["@DateEnd"].Value = this.�����������.DataEnd;// ����SQL.����(this.�����������.DataEnd.ToShortDateString());



                SqlDataAdapter da = new SqlDataAdapter(com);
                da.Fill(dTab);
            }

            return dTab;

            #region
            //// ������� ���� ����������� �� ��������� ������.
            //string queryCOunt = "SELECT    " +
            //               "derivedtbl_1.������������������ " +
            //               " FROM         (SELECT     dbo.��������.id_��������������, dbo.��������������.����������������������, COUNT(dbo.����������.������������������) " +
            //               " AS ������������������������, dbo.����������.������������������ " +
            //           "FROM          dbo.�������� LEFT OUTER JOIN " +
            //                                  "dbo.�������������� ON dbo.��������.id_�������������� = dbo.��������������.id_�������������� LEFT OUTER JOIN " +
            //                                  "dbo.����������������������������� ON dbo.�����������������������������.id�������� = dbo.��������.id_�������� INNER JOIN " +
            //                                  "dbo.���������� ON dbo.����������.id_���������� = dbo.�����������������������������.id���������� " +
            //           "WHERE      (dbo.��������.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and dbo.��������.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
            //           "GROUP BY dbo.��������������.����������������������, dbo.��������.id_��������������, dbo.����������.������������������)  " +
            //          "AS derivedtbl_1 LEFT OUTER JOIN " +
            //              "(SELECT     ��������_4.id_��������������, ��������������_4.����������������������, COUNT(����������_4.������������������) " +
            //                                       "AS ������������������������, ����������_4.������������������ " +
            //                "FROM          dbo.�������� AS ��������_4 LEFT OUTER JOIN " +
            //                                       "dbo.�������������� AS ��������������_4 ON  " +
            //                                       "��������_4.id_�������������� = ��������������_4.id_�������������� LEFT OUTER JOIN " +
            //                                       "dbo.����������������������������� AS �����������������������������_4 ON  " +
            //                                       "�����������������������������_4.id�������� = ��������_4.id_�������� INNER JOIN " +
            //                                       "dbo.���������� AS ����������_4 ON ����������_4.id_���������� = �����������������������������_4.id���������� " +
            //                "WHERE      (��������_4.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_4.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
            //                "GROUP BY ��������������_4.����������������������, ��������_4.id_��������������, ����������_4.������������������, " +
            //                                       " ��������_4.id����������������������� " +
            //                "HAVING      (��������_4.id����������������������� = 4)) AS derivedtbl_5 ON  " +
            //          " derivedtbl_1.id_�������������� = derivedtbl_5.id_�������������� AND " +
            //          " derivedtbl_1.������������������ = derivedtbl_5.������������������ LEFT OUTER JOIN " +
            //              " (SELECT     ��������_3.id_��������������, ��������������_3.����������������������, COUNT(����������_3.������������������) " +
            //                                       " AS ������������������������, ����������_3.������������������ " +
            //                " FROM          dbo.�������� AS ��������_3 LEFT OUTER JOIN " +
            //                                       " dbo.�������������� AS ��������������_3 ON  " +
            //                                       " ��������_3.id_�������������� = ��������������_3.id_�������������� LEFT OUTER JOIN " +
            //                                       " dbo.����������������������������� AS �����������������������������_3 ON  " +
            //                                       " �����������������������������_3.id�������� = ��������_3.id_�������� INNER JOIN " +
            //                                       " dbo.���������� AS ����������_3 ON ����������_3.id_���������� = �����������������������������_3.id���������� " +
            //                "WHERE      (��������_3.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_3.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
            //                " GROUP BY ��������������_3.����������������������, ��������_3.id_��������������, ����������_3.������������������, " +
            //                                      " ��������_3.id����������������������� " +
            //                " HAVING      (��������_3.id����������������������� = 3)) AS derivedtbl_4 ON  " +
            //          " derivedtbl_1.id_�������������� = derivedtbl_4.id_�������������� AND " +
            //          " derivedtbl_1.������������������ = derivedtbl_4.������������������ LEFT OUTER JOIN " +
            //              " (SELECT     ��������_2.id_��������������, ��������������_2.����������������������, COUNT(����������_2.������������������) " +
            //                                       " AS ������������������������, ����������_2.������������������ " +
            //                " FROM          dbo.�������� AS ��������_2 LEFT OUTER JOIN " +
            //                                       " dbo.�������������� AS ��������������_2 ON " +
            //                                       " ��������_2.id_�������������� = ��������������_2.id_�������������� LEFT OUTER JOIN " +
            //                                       " dbo.����������������������������� AS �����������������������������_2 ON " +
            //                                       " �����������������������������_2.id�������� = ��������_2.id_�������� INNER JOIN " +
            //                                       " dbo.���������� AS ����������_2 ON ����������_2.id_���������� = �����������������������������_2.id����������" +
            //                " WHERE      (��������_2.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_2.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
            //               " GROUP BY ��������������_2.����������������������, ��������_2.id_��������������, ����������_2.������������������, " +
            //                                       " ��������_2.id����������������������� " +
            //                " HAVING      (��������_2.id����������������������� = 2)) AS derivedtbl_3 ON  " +
            //          " derivedtbl_1.id_�������������� = derivedtbl_3.id_�������������� AND " +
            //          "derivedtbl_1.������������������ = derivedtbl_3.������������������ LEFT OUTER JOIN " +
            //              " (SELECT     ��������_1.id_��������������, ��������������_1.����������������������, COUNT(����������_1.������������������)  " +
            //                                       " AS ������������������������, ����������_1.������������������ " +
            //                "FROM          dbo.�������� AS ��������_1 LEFT OUTER JOIN " +
            //                                       " dbo.�������������� AS ��������������_1 ON " +
            //                                       "��������_1.id_�������������� = ��������������_1.id_�������������� LEFT OUTER JOIN " +
            //                                       "dbo.����������������������������� AS �����������������������������_1 ON  " +
            //                                       " �����������������������������_1.id�������� = ��������_1.id_�������� INNER JOIN " +
            //                                      " dbo.���������� AS ����������_1 ON ����������_1.id_���������� = �����������������������������_1.id���������� " +
            //                " WHERE      (��������_1.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_1.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "' )" +
            //                " GROUP BY ��������������_1.����������������������, ��������_1.id_��������������, ����������_1.������������������,  " +
            //                                       " ��������_1.id����������������������� " +
            //                " HAVING      (��������_1.id����������������������� = 1)) AS derivedtbl_2 ON  " +
            //          " derivedtbl_1.������������������ = derivedtbl_2.������������������ AND derivedtbl_1.id_�������������� = derivedtbl_2.id_�������������� " +
            //          "  group by derivedtbl_1.������������������ " +
            //          "  order by derivedtbl_1.������������������ ";

            //GetDataTable getData = new GetDataTable(queryCOunt);
            //return getData.DataTable();
            #endregion
        }

        /// <summary>
        /// ���������� �������� �������� ��� ������� ���������.
        /// </summary>
        /// <returns></returns>
        private DataTable SelectCountPerson(string fio)
        {
             DataTable dTab = new DataTable();

            ������������ strConnect = new ������������();

            using (SqlConnection con = new SqlConnection(strConnect.�����������������()))
            {
                SqlCommand com = new SqlCommand("ReportStatInputForPerson", con);
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.Add(new SqlParameter("@DataStart", SqlDbType.DateTime));
                com.Parameters["@DataStart"].Value = this.�����������.DataStart;// ����SQL.����(this.�����������.DataStart.ToShortDateString());
                com.Parameters.Add(new SqlParameter("@DateEnd", SqlDbType.DateTime));
                com.Parameters["@DateEnd"].Value = this.�����������.DataEnd;// ����SQL.����(this.�����������.DataEnd.ToShortDateString());
                com.Parameters.Add(new SqlParameter("@fio", SqlDbType.NVarChar, 255));
                com.Parameters["@fio"].Value = fio;

                SqlDataAdapter da = new SqlDataAdapter(com);
                da.Fill(dTab);
            }
            #region
            //   string query = "SELECT      SUM(derivedtbl_1.������������������������) as '�����������������', " +
         //             " derivedtbl_1.������������������, SUM(derivedtbl_2.������������������������) AS ������, SUM(derivedtbl_3.������������������������) AS ����, " +
         //             " SUM(derivedtbl_4.������������������������) AS ���, SUM(derivedtbl_5.������������������������) AS ���� " +
         //     " FROM         (SELECT     dbo.��������.id_��������������, dbo.��������������.����������������������, COUNT(dbo.����������.������������������) " +
         //     " AS ������������������������, dbo.����������.������������������ " +
         // "FROM          dbo.�������� LEFT OUTER JOIN " +
         //                        "dbo.�������������� ON dbo.��������.id_�������������� = dbo.��������������.id_�������������� LEFT OUTER JOIN " +
         //                        "dbo.����������������������������� ON dbo.�����������������������������.id�������� = dbo.��������.id_�������� INNER JOIN " +
         //                        "dbo.���������� ON dbo.����������.id_���������� = dbo.�����������������������������.id���������� " +
         // "WHERE      (dbo.��������.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and dbo.��������.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
         // "GROUP BY dbo.��������������.����������������������, dbo.��������.id_��������������, dbo.����������.������������������)  " +
         //"AS derivedtbl_1 LEFT OUTER JOIN " +
         //    "(SELECT     ��������_4.id_��������������, ��������������_4.����������������������, COUNT(����������_4.������������������) " +
         //                             "AS ������������������������, ����������_4.������������������ " +
         //      "FROM          dbo.�������� AS ��������_4 LEFT OUTER JOIN " +
         //                             "dbo.�������������� AS ��������������_4 ON  " +
         //                             "��������_4.id_�������������� = ��������������_4.id_�������������� LEFT OUTER JOIN " +
         //                             "dbo.����������������������������� AS �����������������������������_4 ON  " +
         //                             "�����������������������������_4.id�������� = ��������_4.id_�������� INNER JOIN " +
         //                             "dbo.���������� AS ����������_4 ON ����������_4.id_���������� = �����������������������������_4.id���������� " +
         //      "WHERE      (��������_4.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_4.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
         //      "GROUP BY ��������������_4.����������������������, ��������_4.id_��������������, ����������_4.������������������, " +
         //                             " ��������_4.id����������������������� " +
         //      "HAVING      (��������_4.id����������������������� = 4)) AS derivedtbl_5 ON  " +
         //" derivedtbl_1.id_�������������� = derivedtbl_5.id_�������������� AND " +
         //" derivedtbl_1.������������������ = derivedtbl_5.������������������ LEFT OUTER JOIN " +
         //    " (SELECT     ��������_3.id_��������������, ��������������_3.����������������������, COUNT(����������_3.������������������) " +
         //                             " AS ������������������������, ����������_3.������������������ " +
         //      " FROM          dbo.�������� AS ��������_3 LEFT OUTER JOIN " +
         //                             " dbo.�������������� AS ��������������_3 ON  " +
         //                             " ��������_3.id_�������������� = ��������������_3.id_�������������� LEFT OUTER JOIN " +
         //                             " dbo.����������������������������� AS �����������������������������_3 ON  " +
         //                             " �����������������������������_3.id�������� = ��������_3.id_�������� INNER JOIN " +
         //                             " dbo.���������� AS ����������_3 ON ����������_3.id_���������� = �����������������������������_3.id���������� " +
         //      "WHERE      (��������_3.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_3.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
         //      " GROUP BY ��������������_3.����������������������, ��������_3.id_��������������, ����������_3.������������������, " +
         //                            " ��������_3.id����������������������� " +
         //      " HAVING      (��������_3.id����������������������� = 3)) AS derivedtbl_4 ON  " +
         //" derivedtbl_1.id_�������������� = derivedtbl_4.id_�������������� AND " +
         //" derivedtbl_1.������������������ = derivedtbl_4.������������������ LEFT OUTER JOIN " +
         //    " (SELECT     ��������_2.id_��������������, ��������������_2.����������������������, COUNT(����������_2.������������������) " +
         //                             " AS ������������������������, ����������_2.������������������ " +
         //      " FROM          dbo.�������� AS ��������_2 LEFT OUTER JOIN " +
         //                             " dbo.�������������� AS ��������������_2 ON " +
         //                             " ��������_2.id_�������������� = ��������������_2.id_�������������� LEFT OUTER JOIN " +
         //                             " dbo.����������������������������� AS �����������������������������_2 ON " +
         //                             " �����������������������������_2.id�������� = ��������_2.id_�������� INNER JOIN " +
         //                             " dbo.���������� AS ����������_2 ON ����������_2.id_���������� = �����������������������������_2.id����������" +
         //      " WHERE      (��������_2.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_2.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "') " +
         //     " GROUP BY ��������������_2.����������������������, ��������_2.id_��������������, ����������_2.������������������, " +
         //                             " ��������_2.id����������������������� " +
         //      " HAVING      (��������_2.id����������������������� = 2)) AS derivedtbl_3 ON  " +
         //" derivedtbl_1.id_�������������� = derivedtbl_3.id_�������������� AND " +
         //"derivedtbl_1.������������������ = derivedtbl_3.������������������ LEFT OUTER JOIN " +
         //    " (SELECT     ��������_1.id_��������������, ��������������_1.����������������������, COUNT(����������_1.������������������)  " +
         //                             " AS ������������������������, ����������_1.������������������ " +
         //      "FROM          dbo.�������� AS ��������_1 LEFT OUTER JOIN " +
         //                             " dbo.�������������� AS ��������������_1 ON " +
         //                             "��������_1.id_�������������� = ��������������_1.id_�������������� LEFT OUTER JOIN " +
         //                             "dbo.����������������������������� AS �����������������������������_1 ON  " +
         //                             " �����������������������������_1.id�������� = ��������_1.id_�������� INNER JOIN " +
         //                            " dbo.���������� AS ����������_1 ON ����������_1.id_���������� = �����������������������������_1.id���������� " +
         //      " WHERE      (��������_1.���������� >= '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "' and ��������_1.���������� <= '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "' )" +
         //      " GROUP BY ��������������_1.����������������������, ��������_1.id_��������������, ����������_1.������������������,  " +
         //                             " ��������_1.id����������������������� " +
         //      " HAVING      (��������_1.id����������������������� = 1)) AS derivedtbl_2 ON  " +
         //" derivedtbl_1.������������������ = derivedtbl_2.������������������ AND derivedtbl_1.id_�������������� = derivedtbl_2.id_�������������� " +
         //"  where LOWER(LTRIM(RTRIM(derivedtbl_1.������������������))) = '"+ fio.Trim().ToLower() +"' " +
         //" group by derivedtbl_1.[������������������] " +
         //" order by derivedtbl_1.������������������ asc ";

         //   GetDataTable getData = new GetDataTable(query);
            //   return getData.DataTable();
            #endregion

            return dTab;
        }

        /// <summary>
        /// ���������� ����� �� ���������������.
        /// </summary>
        /// <returns></returns>
        private DataTable SelectCountCorrespondent()
        {
            string query = " SELECT     TOP (100) PERCENT ��������������_2.id_��������������, ��������������_2.����������������������, derivedtbl_1.���, derivedtbl_1.������, " +
                      "derivedtbl_1.����, derivedtbl_1.���, derivedtbl_1.���� " +
                        "FROM         (SELECT     derivedtbl_1_6.Expr1 AS ���, derivedtbl_1_6.id_��������������, derivedtbl_2.Expr1 AS ������, derivedtbl_3.Expr1 AS ����,  " +
                                              "derivedtbl_4.Expr1 AS ���, derivedtbl_5.Expr1 AS ���� " +
                       "FROM          (SELECT     SUM(Expr1) AS Expr1, id_�������������� " +
                                               "FROM          (SELECT     TOP (100) PERCENT COUNT(dbo.��������.id_��������) AS Expr1, dbo.��������������.id_�������������� " +
                                                                       "FROM          dbo.�������� LEFT OUTER JOIN " +
                                                                                              "dbo.�������������� ON dbo.��������.id_�������������� = dbo.��������������.id_�������������� " +
                                                                       "GROUP BY dbo.��������.����������, dbo.��������.id�����������������������,  "+
                                                                                              "dbo.��������������.����������������������, dbo.��������������.id_�������������� " +
                                                                       "HAVING      (dbo.��������.���������� >= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                              "(dbo.��������.���������� <= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                       "ORDER BY dbo.��������������.����������������������) AS derivedtbl_1_1 " +
                                               "GROUP BY id_��������������) AS derivedtbl_1_6 LEFT OUTER JOIN " +
                                                  "(SELECT     id_�������������� AS Expr2, SUM(Expr1) AS Expr1 " +
                                                    "FROM          (SELECT     TOP (100) PERCENT COUNT(��������_1.id_��������) AS Expr1, ��������������_1.id_�������������� "+
                                                                            "FROM          dbo.�������� AS ��������_1 LEFT OUTER JOIN " +
                                                                                                   "dbo.�������������� AS ��������������_1 ON  " +
                                                                                                   "��������_1.id_�������������� = ��������������_1.id_�������������� " +
                                                                            "WHERE      (��������_1.id����������������������� = 1) " +
                                                                            "GROUP BY ��������_1.����������, ��������_1.id�����������������������,  " +
                                                                                                   "��������������_1.����������������������, ��������������_1.id_�������������� " +
                                                                            "HAVING      (��������_1.���������� >= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                                   "(��������_1.���������� <= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                            "ORDER BY ��������������_1.����������������������) AS derivedtbl_1_5 " +
                                                    "GROUP BY id_��������������) AS derivedtbl_2 ON derivedtbl_1_6.id_�������������� = derivedtbl_2.Expr2 LEFT OUTER JOIN " +
                                                  "(SELECT     id_�������������� AS Expr2, SUM(Expr1) AS Expr1 " +
                                                    "FROM          (SELECT     TOP (100) PERCENT COUNT(��������_1.id_��������) AS Expr1, ��������������_1.id_�������������� " +
                                                                           " FROM          dbo.�������� AS ��������_1 LEFT OUTER JOIN " +
                                                                                                   "dbo.�������������� AS ��������������_1 ON  " +
                                                                                                   "��������_1.id_�������������� = ��������������_1.id_�������������� " +
                                                                            "WHERE      (��������_1.id����������������������� = 4) " +
                                                                            "GROUP BY ��������_1.����������, ��������_1.id�����������������������,  " +
                                                                                                   "��������������_1.����������������������, ��������������_1.id_�������������� " +
                                                                            "HAVING      (��������_1.���������� >= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                                   "(��������_1.���������� <= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                            " ORDER BY ��������������_1.����������������������) AS derivedtbl_1_4 " +
                                                    " GROUP BY id_��������������) AS derivedtbl_5 ON derivedtbl_1_6.id_�������������� = derivedtbl_5.Expr2 LEFT OUTER JOIN " +
                                                  "(SELECT     id_�������������� AS Expr2, SUM(Expr1) AS Expr1 " +
                                                    "FROM          (SELECT     TOP (100) PERCENT COUNT(��������_1.id_��������) AS Expr1, ��������������_1.id_�������������� " +
                                                                            "FROM          dbo.�������� AS ��������_1 LEFT OUTER JOIN " +
                                                                                                   "dbo.�������������� AS ��������������_1 ON  " +
                                                                                                   "��������_1.id_�������������� = ��������������_1.id_�������������� " +
                                                                            "WHERE      (��������_1.id����������������������� = 3) " +
                                                                            "GROUP BY ��������_1.����������, ��������_1.id�����������������������,  " +
                                                                                                   "��������������_1.����������������������, ��������������_1.id_�������������� " +
                                                                            "HAVING      (��������_1.���������� >= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                                   "(��������_1.���������� <= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                            "ORDER BY ��������������_1.����������������������) AS derivedtbl_1_3 " +
                                                    "GROUP BY id_��������������) AS derivedtbl_4 ON derivedtbl_1_6.id_�������������� = derivedtbl_4.Expr2 LEFT OUTER JOIN " +
                                                  " (SELECT     id_�������������� AS Expr2, SUM(Expr1) AS Expr1 " +
                                                    " FROM          (SELECT     TOP (100) PERCENT COUNT(��������_1.id_��������) AS Expr1, ��������������_1.id_�������������� " +
                                                                            "FROM          dbo.�������� AS ��������_1 LEFT OUTER JOIN " +
                                                                                                   "dbo.�������������� AS ��������������_1 ON  " +
                                                                                                   "��������_1.id_�������������� = ��������������_1.id_�������������� " +
                                                                            "WHERE      (��������_1.id����������������������� = 2) " +
                                                                            "GROUP BY ��������_1.����������, ��������_1.id�����������������������,  " +
                                                                                                   "��������������_1.����������������������, ��������������_1.id_�������������� " +
                                                                            "HAVING      (��������_1.���������� >= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                                   "(��������_1.���������� <= CONVERT(DATETIME, '" + ����SQL.����(this.�����������.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                            "ORDER BY ��������������_1.����������������������) AS derivedtbl_1_2 " +
                                                    "GROUP BY id_��������������) AS derivedtbl_3 ON derivedtbl_1_6.id_�������������� = derivedtbl_3.Expr2)  " +
                      " AS derivedtbl_1 INNER JOIN " + 
                      " dbo.�������������� AS ��������������_2 ON derivedtbl_1.id_�������������� = ��������������_2.id_�������������� " +
"ORDER BY ��������������_2.id_�������������� ";

            GetDataTable getData = new GetDataTable(query);
            return getData.DataTable();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

            //List<StatisticDocInput> list = new List<StatisticDocInput>();

            //StatisticDocInput itemHead = new StatisticDocInput();
            //// ����� ������.
            ////�������������������������� itemHead = new ��������������������������();
            //itemHead.Num = "� �.�.";
            //itemHead.�������������������������� = "������������ ��������������";
            //itemHead.������������������������ = "���������� ��������� ����������";
            //itemHead.���������������� = "�������� ��������";
            //itemHead.VipNet = "VipNet";
            //itemHead.Fax = "����";
            //itemHead.Email = "e-mail";
            //itemHead.����������� = "�����������";

            //list.Add(itemHead);

            //list.AddRange(this.listPrint);

            List<StatisticDocInput> list = this.listPrint;

            string caption = "���������� �� �������� ��������������� �� ������ � " + rd.DataStart.ToShortDateString() + " �� " + rd.DataEnd.ToShortDateString();
            Report���������������������������� printDate = new Report����������������������������(caption);
            printDate.SetDate = list;

            PrintReport printPaper = new PrintReport();
            printPaper.SetCommand(printDate);
            printPaper.Execute();

            //// ����� �����.
            //StatisticDocInput head = new StatisticDocInput();
            //head.Num = "� �.�.";
            //head.�������������������������� = "������������ ��������������";
            //head.������������������������ = "���������� �������� ����������";
            //head.���������������� = "�������� �������� ��.";
            //head.Email = "����������� ����� ��.";
            //head.VipNet = "VipNet ��.";
            //head.Fax = "���� ��.";
            //head.����������� = "�����������";

            //list.Add(head);

            ////// ������� �������� �� �������.
            //////�������� �������� WORD
            ////string fName = "���������������������������������";

            ////// �������� ����� ���������.
            ////DirectoryInfo dirInfo = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + @"\���������\");

            ////foreach (FileInfo file in dirInfo.GetFiles())
            ////{
            ////    file.Delete();
            ////}

            //////��������� ������ � ����� ���������
            //////try
            //////{
            ////    FileInfo fn = new FileInfo(System.Windows.Forms.Application.StartupPath + @"\������\���������� �� �������� ���������������.doc");
            ////    fn.CopyTo(System.Windows.Forms.Application.StartupPath + @"\���������\" + fName + ".doc", true);
            //////}
            //////catch (Exception ex)
            //////{
            //////    MessageBox.Show("�������� ��� ������.\n" + ex.Message, "������");
            //////    return;
            //////}

            ////string filName = System.Windows.Forms.Application.StartupPath + @"\���������\" + fName + ".doc";

            //////������ ����� Word.Application
            ////Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //////��������� ��������
            ////Microsoft.Office.Interop.Word.Document doc = null;

            ////object fileName = filName;
            ////object falseValue = false;
            ////object trueValue = true;
            ////object missing = Type.Missing;

            ////doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ////ref missing, ref missing, ref missing, ref missing, ref missing,
            ////ref missing, ref missing, ref missing, ref missing, ref missing,
            ////ref missing, ref missing, ref missing);

            //////���� ���������.
            ////object wdrepl2 = WdReplace.wdReplaceAll;
            //////object searchtxt = "GreetingLine";
            ////object searchtxt2 = "datestart";
            ////object newtxt2 = (object)this.�����������.DataStart.ToShortDateString();
            //////object frwd = true;
            ////object frwd2 = false;
            ////doc.Content.Find.Execute(ref searchtxt2, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd2, ref missing, ref missing, ref newtxt2, ref wdrepl2, ref missing, ref missing,
            ////ref missing, ref missing);

            //////���� ��������.
            ////object wdrepl3 = WdReplace.wdReplaceAll;
            //////object searchtxt = "GreetingLine";
            ////object searchtxt3 = "dateend";
            ////object newtxt3 = (object)this.�����������.DataEnd.ToShortDateString();
            //////object frwd = true;
            ////object frwd3 = false;
            ////doc.Content.Find.Execute(ref searchtxt3, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd3, ref missing, ref missing, ref newtxt3, ref wdrepl3, ref missing, ref missing,
            ////ref missing, ref missing);

            //////�������� �������
            ////object bookNaziv = "�������";
            ////Range wrdRng = doc.Bookmarks.get_Item(ref  bookNaziv).Range;

            ////object behavior = Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord8TableBehavior;
            ////object autobehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow;


            ////Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(wrdRng, 1, 8, ref behavior, ref autobehavior);
            ////table.Range.ParagraphFormat.SpaceAfter = 8;

            //////�������� ������ ��������
            ////table.Columns[1].Width = 80;
            ////table.Columns[2].Width = 260;
            ////table.Columns[3].Width = 80;
            ////table.Columns[4].Width = 60;
            ////table.Columns[5].Width = 60;
            ////table.Columns[6].Width = 60;
            ////table.Columns[7].Width = 60;
            ////table.Columns[8].Width = 80;
            ////table.Borders.Enable = 1; // ����� - �������� �����
            ////table.Range.Font.Name = "Times New Roman";
            ////table.Range.Font.Size = 10;
            //////������� �����
            ////int i = 1;

                      
            //////// ������� ����� � �������.
            //////table.Cell(i, 1).Range.Text = "� �.�.";
            //////table.Cell(i, 2).Range.Text = "������������ ��������������";

            //////table.Cell(i, 3).Range.Text = "���������� �������� ����������";
            //////table.Cell(i, 4).Range.Text = "�������� �������� ��.";

            //////table.Cell(i, 5).Range.Text = "����������� ����� ��.";
            //////table.Cell(i, 6).Range.Text = "VipNet ��.";

            //////table.Cell(i, 7).Range.Text = "���� ��.";
            //////table.Cell(i, 8).Range.Text = "�����������";



            ////// ������� ����� � �������.
            ////table.Cell(1, 1).Range.Text = "� �.�.";
            ////table.Cell(1, 2).Range.Text = "������������ ��������������";

            ////table.Cell(1, 3).Range.Text = "���������� �������� ����������";
            ////table.Cell(1, 4).Range.Text = "�������� �������� ��.";

            ////table.Cell(1, 5).Range.Text = "����������� ����� ��.";
            ////table.Cell(1, 6).Range.Text = "VipNet ��.";

            ////table.Cell(1, 7).Range.Text = "���� ��.";
            ////table.Cell(1, 8).Range.Text = "�����������";

            //////doc.Words.Count.ToString();
            ////Object beforeRow1 = Type.Missing;
            ////table.Rows.Add(ref beforeRow1);

            //////i++;
            ////i = 2;

            //////������� ������ � �������
            ////foreach (StatisticDocInput item in list)
            ////{
            ////    table.Cell(i, 1).Range.Text = item.Num;
            ////    table.Cell(i, 2).Range.Text = item.��������������������������.Trim();

            ////    if (item.������������������������ > 0)
            ////    {
            ////        table.Cell(i, 3).Range.Text = item.������������������������.ToString().Trim();
            ////    }
            ////    else
            ////    {
            ////        table.Cell(i, 3).Range.Text = "";
            ////    }

            ////    if (item.���������������� > 0)
            ////    {
            ////        table.Cell(i, 4).Range.Text = item.����������������.ToString().Trim();
            ////    }
            ////    else
            ////    {
            ////        table.Cell(i, 4).Range.Text = "";
            ////    }

            ////    if (item.Email > 0)
            ////    {
            ////        table.Cell(i, 5).Range.Text = item.Email.ToString().Trim();
            ////    }
            ////    else
            ////    {
            ////        table.Cell(i, 5).Range.Text = "";
            ////    }

            ////    if (item.VipNet > 0)
            ////    {
            ////        table.Cell(i, 6).Range.Text = item.VipNet.ToString().Trim();
            ////    }
            ////    else
            ////    {
            ////        table.Cell(i, 6).Range.Text = "";
            ////    }

            ////    if (item.Fax > 0)
            ////    {
            ////        table.Cell(i, 7).Range.Text = item.Fax.ToString().Trim();
            ////    }
            ////    else
            ////    {
            ////        table.Cell(i, 7).Range.Text = "";
            ////    }
            ////    table.Cell(i, 8).Range.Text = item.�����������.Trim();

            ////    //doc.Words.Count.ToString();
            ////    Object beforeRow2 = Type.Missing;
            ////    table.Rows.Add(ref beforeRow2);

            ////    i++;
            ////}
            ////table.Rows[i].Delete();


            ////// ��������� ��������.
            ////app.Visible = true;
            
            //DataTable tPerson =  SelectPerson();

            //List<��������������������������> list = new List<��������������������������>();

            //foreach (DataRow row in tPerson.Rows)
            //{
            //    // ������� �������� ����� ��� �������� ������������.
            //    DataRow[] rows = rez.Select("������������������ = '" + row["������������������"].ToString().Trim() + "' ");

            //    int iCount = 1;
            //    foreach (DataRow r in rows)
            //    {
            //        �������������������������� itm = new ��������������������������();
            //        itm.������� = iCount.ToString();
            //        itm.�������������������������� = r["����������������������"].ToString().Trim();

            //        itm.������������������������ = r["����������������������������"].ToString().Trim();
            //        itm.����������������� = r["������"].ToString().Trim();
            //        itm.EMail = r["e-mail"].ToString().Trim();
            //        itm.VipNet = r["VipNet"].ToString().Trim();
            //        itm.Fax = r["����"].ToString().Trim();
            //        itm.����������� = r["������������������"].ToString().Trim();

            //        list.Add(itm);

            //        iCount++;
            //    }

            //    DataRow rCount = SelectCountPerson(row["������������������"].ToString().Trim()).Rows[0];

            //    // �������� ������ ����� ��� �����������.
            //    �������������������������� itCount = new ��������������������������();
            //    itCount.������� = "����� �� ����������� " + rCount["������������������"].ToString().Trim();
            //    itCount.����������� = "----------";
            //    itCount.������������������������ = rCount["�����������������"].ToString().Trim();
            //    itCount.����������������� = rCount["������"].ToString().Trim();
            //    itCount.EMail = rCount["����"].ToString().Trim();
            //    itCount.VipNet = rCount["���"].ToString().Trim();
            //    itCount.Fax = rCount["����"].ToString().Trim();
            //    itCount.����������� = rCount["������������������"].ToString().Trim();

            //    list.Add(itCount);
            //}

            //int iCount2 = 1;

            //int count������������ = 0;
            //int count��������� = 0;
            //int count������� = 0;
            //int count������ = 0;
            //int count������� = 0;

            //DataTable dtRowsCorrespondent = SelectCountCorrespondent();
            //// ������� ����� �� �� �������� �� ���������������.
            //foreach (DataRow r in dtRowsCorrespondent.Rows)
            //{
            //    �������������������������� itm = new ��������������������������();
            //    itm.������� = iCount2.ToString();
            //    itm.�������������������������� = r["����������������������"].ToString().Trim();

            //    itm.������������������������ = r["���"].ToString().Trim();

            //    if (DBNull.Value != r["���"])
            //    {
            //        count������������ += Convert.ToInt32(r["���"]);
            //    }

            //    itm.����������������� = r["������"].ToString().Trim();

            //    if (DBNull.Value != r["������"])
            //    {
            //        count��������� += Convert.ToInt32(r["������"]);
            //    }

            //    itm.EMail = r["����"].ToString().Trim();

            //    if (DBNull.Value != r["����"])
            //    {
            //        count������� += Convert.ToInt32(r["����"]);
            //    }

            //    itm.VipNet = r["���"].ToString().Trim();

            //    if (DBNull.Value != r["���"])
            //    {
            //        count������ += Convert.ToInt32(r["���"]);
            //    }

            //    itm.Fax = r["����"].ToString().Trim();

            //    if (DBNull.Value != r["����"])
            //    {
            //        count������� += Convert.ToInt32(r["����"]);
            //    }

            //    itm.����������� = "-----";

            //    list.Add(itm);

            //    iCount2++;
            //}

            //�������������������������� count = new ��������������������������();
            //count.������� = "����� � ����� �� ��������";
            //count.�������������������������� = "-----";
            //count.������������������������ = count������������.ToString();
            //count.����������������� = count���������.ToString();
            //count.EMail = count�������.ToString();
            //count.VipNet = count������.ToString();
            //count.Fax = count�������.ToString();

            //list.Add(count);
            

            //List<��������������������������> listTest = list;



            //ExcelPrint excel = new ExcelPrint(" ���������� �� �������� ��������������� c " + this.�����������.DataStart.ToShortDateString() + " �� " + this.�����������.DataEnd.ToShortDateString() + " �� ���� �. ��������");
            //excel.Print���������������������������������(list);

            // ��������� �� ��� � Excel.
            this.Close();

        }
    }
}