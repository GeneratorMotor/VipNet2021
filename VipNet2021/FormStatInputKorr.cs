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
        /// Хранит диапазон дат.
        /// </summary>
        public RangeDate ДиапазонДат
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

            // Строка содержит значение ИТОГО.
            StatisticDocInput count = new StatisticDocInput();

            foreach (StatisticDocInput it in this.listCount)
            {
                count.Num = "Итого в целом по комитету.";
                count.КолвоВходКорреспонденции += it.КолвоВходКорреспонденции;
                count.БумажныйНоситель += it.БумажныйНоситель;
                count.VipNet += it.VipNet;
                count.Email += it.Email;
                count.Fax += it.Fax;
            }

            list.Add(count);

            this.listPrint = list;

            this.dataGridView1.DataSource = list;

            //this.dataGridView1.Columns["мыло"].HeaderText = "e-mail";
        }

        /// <summary>
        /// Возвращает полученные документы за определённый период.
        /// </summary>
        /// <returns></returns>
        private List<StatisticDocInput> SelectCorr()
        {
            DataTable dTabНачОтдела = new DataTable();

            ПодключитьБД strConnect = new ПодключитьБД();

            List<StatisticDocInput> list = new List<StatisticDocInput>();

            using(SqlConnection con = new SqlConnection(strConnect.СтрокаПодключения()))
            {

                // Получим всех начальников отделов.
                SqlCommand com1 = new SqlCommand("ReportStatInputLetterPerson", con);
                com1.CommandType = CommandType.StoredProcedure;
                com1.Parameters.Add(new SqlParameter("@DataStart", SqlDbType.DateTime));
                com1.Parameters["@DataStart"].Value = this.ДиапазонДат.DataStart;// ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString());
                com1.Parameters.Add(new SqlParameter("@DateEnd", SqlDbType.DateTime));
                com1.Parameters["@DateEnd"].Value = this.ДиапазонДат.DataEnd;// ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString());

                SqlDataAdapter da = new SqlDataAdapter(com1);
                da.Fill(dTabНачОтдела);

                foreach (DataRow r in dTabНачОтдела.Rows)
                {
                    GenerateStaticDocInput generatorReport = new GenerateStaticDocInput(list);
                    generatorReport.Generate(this.ДиапазонДат.DataStart, this.ДиапазонДат.DataEnd, r["ОписаниеПолучателя"].ToString().Trim());

                    StatisticDocInput itmCount = new StatisticDocInput();
                    itmCount = generatorReport.ItemCount;

                    listCount.Add(itmCount);

                }

               
            }

            #region
            //string query = "SELECT     derivedtbl_1.id_корреспондента, derivedtbl_1.ОписаниеКорреспондента, derivedtbl_1.КоличествоВходДокументов, " +
            //              "derivedtbl_1.ОписаниеПолучателя, derivedtbl_2.КоличествоВходДокументов AS бумага, derivedtbl_3.КоличествоВходДокументов AS мыло, " +
            //              "derivedtbl_4.КоличествоВходДокументов AS вип, derivedtbl_5.КоличествоВходДокументов AS факс " +
            //              " FROM         (SELECT     dbo.Карточка.id_корреспондента, dbo.Корреспонденты.ОписаниеКорреспондента, COUNT(dbo.Получатели.ОписаниеПолучателя) " +
            //              " AS КоличествоВходДокументов, dbo.Получатели.ОписаниеПолучателя " +
            //          "FROM          dbo.Карточка LEFT OUTER JOIN " +
            //                                 "dbo.Корреспонденты ON dbo.Карточка.id_корреспондента = dbo.Корреспонденты.id_корреспондента LEFT OUTER JOIN " +
            //                                 "dbo.ПолучателДокументовУправление ON dbo.ПолучателДокументовУправление.idКарточки = dbo.Карточка.id_карточки INNER JOIN " +
            //                                 "dbo.Получатели ON dbo.Получатели.id_получателя = dbo.ПолучателДокументовУправление.idПолучатель " +
            //          "WHERE      (dbo.Карточка.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and dbo.Карточка.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
            //          "GROUP BY dbo.Корреспонденты.ОписаниеКорреспондента, dbo.Карточка.id_корреспондента, dbo.Получатели.ОписаниеПолучателя)  " +
            //         "AS derivedtbl_1 LEFT OUTER JOIN " +
            //             "(SELECT     Карточка_4.id_корреспондента, Корреспонденты_4.ОписаниеКорреспондента, COUNT(Получатели_4.ОписаниеПолучателя) " +
            //                                      "AS КоличествоВходДокументов, Получатели_4.ОписаниеПолучателя " +
            //               "FROM          dbo.Карточка AS Карточка_4 LEFT OUTER JOIN " +
            //                                      "dbo.Корреспонденты AS Корреспонденты_4 ON  " +
            //                                      "Карточка_4.id_корреспондента = Корреспонденты_4.id_корреспондента LEFT OUTER JOIN " +
            //                                      "dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_4 ON  " +
            //                                      "ПолучателДокументовУправление_4.idКарточки = Карточка_4.id_карточки INNER JOIN " +
            //                                      "dbo.Получатели AS Получатели_4 ON Получатели_4.id_получателя = ПолучателДокументовУправление_4.idПолучатель " +
            //               "WHERE      (Карточка_4.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_4.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
            //               "GROUP BY Корреспонденты_4.ОписаниеКорреспондента, Карточка_4.id_корреспондента, Получатели_4.ОписаниеПолучателя, " +
            //                                      " Карточка_4.idВидПоступленияДокумента " +
            //               "HAVING      (Карточка_4.idВидПоступленияДокумента = 4)) AS derivedtbl_5 ON  " +
            //         " derivedtbl_1.id_корреспондента = derivedtbl_5.id_корреспондента AND " +
            //         " derivedtbl_1.ОписаниеПолучателя = derivedtbl_5.ОписаниеПолучателя LEFT OUTER JOIN " +
            //             " (SELECT     Карточка_3.id_корреспондента, Корреспонденты_3.ОписаниеКорреспондента, COUNT(Получатели_3.ОписаниеПолучателя) " +
            //                                      " AS КоличествоВходДокументов, Получатели_3.ОписаниеПолучателя " +
            //               " FROM          dbo.Карточка AS Карточка_3 LEFT OUTER JOIN " +
            //                                      " dbo.Корреспонденты AS Корреспонденты_3 ON  " +
            //                                      " Карточка_3.id_корреспондента = Корреспонденты_3.id_корреспондента LEFT OUTER JOIN " +
            //                                      " dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_3 ON  " +
            //                                      " ПолучателДокументовУправление_3.idКарточки = Карточка_3.id_карточки INNER JOIN " +
            //                                      " dbo.Получатели AS Получатели_3 ON Получатели_3.id_получателя = ПолучателДокументовУправление_3.idПолучатель " +
            //               "WHERE      (Карточка_3.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_3.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
            //               " GROUP BY Корреспонденты_3.ОписаниеКорреспондента, Карточка_3.id_корреспондента, Получатели_3.ОписаниеПолучателя, " +
            //                                     " Карточка_3.idВидПоступленияДокумента " +
            //               " HAVING      (Карточка_3.idВидПоступленияДокумента = 3)) AS derivedtbl_4 ON  " +
            //         " derivedtbl_1.id_корреспондента = derivedtbl_4.id_корреспондента AND " +
            //         " derivedtbl_1.ОписаниеПолучателя = derivedtbl_4.ОписаниеПолучателя LEFT OUTER JOIN " +
            //             " (SELECT     Карточка_2.id_корреспондента, Корреспонденты_2.ОписаниеКорреспондента, COUNT(Получатели_2.ОписаниеПолучателя) " +
            //                                      " AS КоличествоВходДокументов, Получатели_2.ОписаниеПолучателя " +
            //               " FROM          dbo.Карточка AS Карточка_2 LEFT OUTER JOIN " +
            //                                      " dbo.Корреспонденты AS Корреспонденты_2 ON " +
            //                                      " Карточка_2.id_корреспондента = Корреспонденты_2.id_корреспондента LEFT OUTER JOIN " +
            //                                      " dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_2 ON " +
            //                                      " ПолучателДокументовУправление_2.idКарточки = Карточка_2.id_карточки INNER JOIN " +
            //                                      " dbo.Получатели AS Получатели_2 ON Получатели_2.id_получателя = ПолучателДокументовУправление_2.idПолучатель" +
            //               " WHERE      (Карточка_2.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_2.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
            //              " GROUP BY Корреспонденты_2.ОписаниеКорреспондента, Карточка_2.id_корреспондента, Получатели_2.ОписаниеПолучателя, " +
            //                                      " Карточка_2.idВидПоступленияДокумента " +
            //               " HAVING      (Карточка_2.idВидПоступленияДокумента = 2)) AS derivedtbl_3 ON  " +
            //         " derivedtbl_1.id_корреспондента = derivedtbl_3.id_корреспондента AND " +
            //         "derivedtbl_1.ОписаниеПолучателя = derivedtbl_3.ОписаниеПолучателя LEFT OUTER JOIN " +
            //             " (SELECT     Карточка_1.id_корреспондента, Корреспонденты_1.ОписаниеКорреспондента, COUNT(Получатели_1.ОписаниеПолучателя)  " +
            //                                      " AS КоличествоВходДокументов, Получатели_1.ОписаниеПолучателя " +
            //               "FROM          dbo.Карточка AS Карточка_1 LEFT OUTER JOIN " +
            //                                      " dbo.Корреспонденты AS Корреспонденты_1 ON " +
            //                                      "Карточка_1.id_корреспондента = Корреспонденты_1.id_корреспондента LEFT OUTER JOIN " +
            //                                      "dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_1 ON  " +
            //                                      " ПолучателДокументовУправление_1.idКарточки = Карточка_1.id_карточки INNER JOIN " +
            //                                     " dbo.Получатели AS Получатели_1 ON Получатели_1.id_получателя = ПолучателДокументовУправление_1.idПолучатель " +
            //               " WHERE      (Карточка_1.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_1.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "' )" +
            //               " GROUP BY Корреспонденты_1.ОписаниеКорреспондента, Карточка_1.id_корреспондента, Получатели_1.ОписаниеПолучателя,  " +
            //                                      " Карточка_1.idВидПоступленияДокумента " +
            //               " HAVING      (Карточка_1.idВидПоступленияДокумента = 1)) AS derivedtbl_2 ON  " +
            //         " derivedtbl_1.ОписаниеПолучателя = derivedtbl_2.ОписаниеПолучателя AND derivedtbl_1.id_корреспондента = derivedtbl_2.id_корреспондента " +
            //         " order by derivedtbl_1.ОписаниеПолучателя asc ";
            #endregion

            //GetDataTable getData = new GetDataTable(query);
            //return getData.DataTable();

            return list;
        }


        /// <summary>
        /// Возвращает Исполнителей которым были отписаны документы за указанный период.
        /// </summary>
        /// <returns></returns>
        private DataTable SelectPerson()
        {
            DataTable dTab = new DataTable();

            ПодключитьБД strConnect = new ПодключитьБД();

            using (SqlConnection con = new SqlConnection(strConnect.СтрокаПодключения()))
            {
                SqlCommand com = new SqlCommand("ReportStatInputLetterPerson", con);
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.Add(new SqlParameter("@DataStart", SqlDbType.DateTime));
                com.Parameters["@DataStart"].Value = this.ДиапазонДат.DataStart;// ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString());
                com.Parameters.Add(new SqlParameter("@DateEnd", SqlDbType.DateTime));
                com.Parameters["@DateEnd"].Value = this.ДиапазонДат.DataEnd;// ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString());



                SqlDataAdapter da = new SqlDataAdapter(com);
                da.Fill(dTab);
            }

            return dTab;

            #region
            //// Получим всех получателей за указанный период.
            //string queryCOunt = "SELECT    " +
            //               "derivedtbl_1.ОписаниеПолучателя " +
            //               " FROM         (SELECT     dbo.Карточка.id_корреспондента, dbo.Корреспонденты.ОписаниеКорреспондента, COUNT(dbo.Получатели.ОписаниеПолучателя) " +
            //               " AS КоличествоВходДокументов, dbo.Получатели.ОписаниеПолучателя " +
            //           "FROM          dbo.Карточка LEFT OUTER JOIN " +
            //                                  "dbo.Корреспонденты ON dbo.Карточка.id_корреспондента = dbo.Корреспонденты.id_корреспондента LEFT OUTER JOIN " +
            //                                  "dbo.ПолучателДокументовУправление ON dbo.ПолучателДокументовУправление.idКарточки = dbo.Карточка.id_карточки INNER JOIN " +
            //                                  "dbo.Получатели ON dbo.Получатели.id_получателя = dbo.ПолучателДокументовУправление.idПолучатель " +
            //           "WHERE      (dbo.Карточка.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and dbo.Карточка.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
            //           "GROUP BY dbo.Корреспонденты.ОписаниеКорреспондента, dbo.Карточка.id_корреспондента, dbo.Получатели.ОписаниеПолучателя)  " +
            //          "AS derivedtbl_1 LEFT OUTER JOIN " +
            //              "(SELECT     Карточка_4.id_корреспондента, Корреспонденты_4.ОписаниеКорреспондента, COUNT(Получатели_4.ОписаниеПолучателя) " +
            //                                       "AS КоличествоВходДокументов, Получатели_4.ОписаниеПолучателя " +
            //                "FROM          dbo.Карточка AS Карточка_4 LEFT OUTER JOIN " +
            //                                       "dbo.Корреспонденты AS Корреспонденты_4 ON  " +
            //                                       "Карточка_4.id_корреспондента = Корреспонденты_4.id_корреспондента LEFT OUTER JOIN " +
            //                                       "dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_4 ON  " +
            //                                       "ПолучателДокументовУправление_4.idКарточки = Карточка_4.id_карточки INNER JOIN " +
            //                                       "dbo.Получатели AS Получатели_4 ON Получатели_4.id_получателя = ПолучателДокументовУправление_4.idПолучатель " +
            //                "WHERE      (Карточка_4.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_4.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
            //                "GROUP BY Корреспонденты_4.ОписаниеКорреспондента, Карточка_4.id_корреспондента, Получатели_4.ОписаниеПолучателя, " +
            //                                       " Карточка_4.idВидПоступленияДокумента " +
            //                "HAVING      (Карточка_4.idВидПоступленияДокумента = 4)) AS derivedtbl_5 ON  " +
            //          " derivedtbl_1.id_корреспондента = derivedtbl_5.id_корреспондента AND " +
            //          " derivedtbl_1.ОписаниеПолучателя = derivedtbl_5.ОписаниеПолучателя LEFT OUTER JOIN " +
            //              " (SELECT     Карточка_3.id_корреспондента, Корреспонденты_3.ОписаниеКорреспондента, COUNT(Получатели_3.ОписаниеПолучателя) " +
            //                                       " AS КоличествоВходДокументов, Получатели_3.ОписаниеПолучателя " +
            //                " FROM          dbo.Карточка AS Карточка_3 LEFT OUTER JOIN " +
            //                                       " dbo.Корреспонденты AS Корреспонденты_3 ON  " +
            //                                       " Карточка_3.id_корреспондента = Корреспонденты_3.id_корреспондента LEFT OUTER JOIN " +
            //                                       " dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_3 ON  " +
            //                                       " ПолучателДокументовУправление_3.idКарточки = Карточка_3.id_карточки INNER JOIN " +
            //                                       " dbo.Получатели AS Получатели_3 ON Получатели_3.id_получателя = ПолучателДокументовУправление_3.idПолучатель " +
            //                "WHERE      (Карточка_3.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_3.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
            //                " GROUP BY Корреспонденты_3.ОписаниеКорреспондента, Карточка_3.id_корреспондента, Получатели_3.ОписаниеПолучателя, " +
            //                                      " Карточка_3.idВидПоступленияДокумента " +
            //                " HAVING      (Карточка_3.idВидПоступленияДокумента = 3)) AS derivedtbl_4 ON  " +
            //          " derivedtbl_1.id_корреспондента = derivedtbl_4.id_корреспондента AND " +
            //          " derivedtbl_1.ОписаниеПолучателя = derivedtbl_4.ОписаниеПолучателя LEFT OUTER JOIN " +
            //              " (SELECT     Карточка_2.id_корреспондента, Корреспонденты_2.ОписаниеКорреспондента, COUNT(Получатели_2.ОписаниеПолучателя) " +
            //                                       " AS КоличествоВходДокументов, Получатели_2.ОписаниеПолучателя " +
            //                " FROM          dbo.Карточка AS Карточка_2 LEFT OUTER JOIN " +
            //                                       " dbo.Корреспонденты AS Корреспонденты_2 ON " +
            //                                       " Карточка_2.id_корреспондента = Корреспонденты_2.id_корреспондента LEFT OUTER JOIN " +
            //                                       " dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_2 ON " +
            //                                       " ПолучателДокументовУправление_2.idКарточки = Карточка_2.id_карточки INNER JOIN " +
            //                                       " dbo.Получатели AS Получатели_2 ON Получатели_2.id_получателя = ПолучателДокументовУправление_2.idПолучатель" +
            //                " WHERE      (Карточка_2.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_2.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
            //               " GROUP BY Корреспонденты_2.ОписаниеКорреспондента, Карточка_2.id_корреспондента, Получатели_2.ОписаниеПолучателя, " +
            //                                       " Карточка_2.idВидПоступленияДокумента " +
            //                " HAVING      (Карточка_2.idВидПоступленияДокумента = 2)) AS derivedtbl_3 ON  " +
            //          " derivedtbl_1.id_корреспондента = derivedtbl_3.id_корреспондента AND " +
            //          "derivedtbl_1.ОписаниеПолучателя = derivedtbl_3.ОписаниеПолучателя LEFT OUTER JOIN " +
            //              " (SELECT     Карточка_1.id_корреспондента, Корреспонденты_1.ОписаниеКорреспондента, COUNT(Получатели_1.ОписаниеПолучателя)  " +
            //                                       " AS КоличествоВходДокументов, Получатели_1.ОписаниеПолучателя " +
            //                "FROM          dbo.Карточка AS Карточка_1 LEFT OUTER JOIN " +
            //                                       " dbo.Корреспонденты AS Корреспонденты_1 ON " +
            //                                       "Карточка_1.id_корреспондента = Корреспонденты_1.id_корреспондента LEFT OUTER JOIN " +
            //                                       "dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_1 ON  " +
            //                                       " ПолучателДокументовУправление_1.idКарточки = Карточка_1.id_карточки INNER JOIN " +
            //                                      " dbo.Получатели AS Получатели_1 ON Получатели_1.id_получателя = ПолучателДокументовУправление_1.idПолучатель " +
            //                " WHERE      (Карточка_1.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_1.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "' )" +
            //                " GROUP BY Корреспонденты_1.ОписаниеКорреспондента, Карточка_1.id_корреспондента, Получатели_1.ОписаниеПолучателя,  " +
            //                                       " Карточка_1.idВидПоступленияДокумента " +
            //                " HAVING      (Карточка_1.idВидПоступленияДокумента = 1)) AS derivedtbl_2 ON  " +
            //          " derivedtbl_1.ОписаниеПолучателя = derivedtbl_2.ОписаниеПолучателя AND derivedtbl_1.id_корреспондента = derivedtbl_2.id_корреспондента " +
            //          "  group by derivedtbl_1.ОписаниеПолучателя " +
            //          "  order by derivedtbl_1.ОписаниеПолучателя ";

            //GetDataTable getData = new GetDataTable(queryCOunt);
            //return getData.DataTable();
            #endregion
        }

        /// <summary>
        /// Возвращает итоговое значение для каждого льготника.
        /// </summary>
        /// <returns></returns>
        private DataTable SelectCountPerson(string fio)
        {
             DataTable dTab = new DataTable();

            ПодключитьБД strConnect = new ПодключитьБД();

            using (SqlConnection con = new SqlConnection(strConnect.СтрокаПодключения()))
            {
                SqlCommand com = new SqlCommand("ReportStatInputForPerson", con);
                com.CommandType = CommandType.StoredProcedure;
                com.Parameters.Add(new SqlParameter("@DataStart", SqlDbType.DateTime));
                com.Parameters["@DataStart"].Value = this.ДиапазонДат.DataStart;// ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString());
                com.Parameters.Add(new SqlParameter("@DateEnd", SqlDbType.DateTime));
                com.Parameters["@DateEnd"].Value = this.ДиапазонДат.DataEnd;// ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString());
                com.Parameters.Add(new SqlParameter("@fio", SqlDbType.NVarChar, 255));
                com.Parameters["@fio"].Value = fio;

                SqlDataAdapter da = new SqlDataAdapter(com);
                da.Fill(dTab);
            }
            #region
            //   string query = "SELECT      SUM(derivedtbl_1.КоличествоВходДокументов) as 'КолВходДокументов', " +
         //             " derivedtbl_1.ОписаниеПолучателя, SUM(derivedtbl_2.КоличествоВходДокументов) AS бумага, SUM(derivedtbl_3.КоличествоВходДокументов) AS мыло, " +
         //             " SUM(derivedtbl_4.КоличествоВходДокументов) AS вип, SUM(derivedtbl_5.КоличествоВходДокументов) AS факс " +
         //     " FROM         (SELECT     dbo.Карточка.id_корреспондента, dbo.Корреспонденты.ОписаниеКорреспондента, COUNT(dbo.Получатели.ОписаниеПолучателя) " +
         //     " AS КоличествоВходДокументов, dbo.Получатели.ОписаниеПолучателя " +
         // "FROM          dbo.Карточка LEFT OUTER JOIN " +
         //                        "dbo.Корреспонденты ON dbo.Карточка.id_корреспондента = dbo.Корреспонденты.id_корреспондента LEFT OUTER JOIN " +
         //                        "dbo.ПолучателДокументовУправление ON dbo.ПолучателДокументовУправление.idКарточки = dbo.Карточка.id_карточки INNER JOIN " +
         //                        "dbo.Получатели ON dbo.Получатели.id_получателя = dbo.ПолучателДокументовУправление.idПолучатель " +
         // "WHERE      (dbo.Карточка.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and dbo.Карточка.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
         // "GROUP BY dbo.Корреспонденты.ОписаниеКорреспондента, dbo.Карточка.id_корреспондента, dbo.Получатели.ОписаниеПолучателя)  " +
         //"AS derivedtbl_1 LEFT OUTER JOIN " +
         //    "(SELECT     Карточка_4.id_корреспондента, Корреспонденты_4.ОписаниеКорреспондента, COUNT(Получатели_4.ОписаниеПолучателя) " +
         //                             "AS КоличествоВходДокументов, Получатели_4.ОписаниеПолучателя " +
         //      "FROM          dbo.Карточка AS Карточка_4 LEFT OUTER JOIN " +
         //                             "dbo.Корреспонденты AS Корреспонденты_4 ON  " +
         //                             "Карточка_4.id_корреспондента = Корреспонденты_4.id_корреспондента LEFT OUTER JOIN " +
         //                             "dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_4 ON  " +
         //                             "ПолучателДокументовУправление_4.idКарточки = Карточка_4.id_карточки INNER JOIN " +
         //                             "dbo.Получатели AS Получатели_4 ON Получатели_4.id_получателя = ПолучателДокументовУправление_4.idПолучатель " +
         //      "WHERE      (Карточка_4.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_4.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
         //      "GROUP BY Корреспонденты_4.ОписаниеКорреспондента, Карточка_4.id_корреспондента, Получатели_4.ОписаниеПолучателя, " +
         //                             " Карточка_4.idВидПоступленияДокумента " +
         //      "HAVING      (Карточка_4.idВидПоступленияДокумента = 4)) AS derivedtbl_5 ON  " +
         //" derivedtbl_1.id_корреспондента = derivedtbl_5.id_корреспондента AND " +
         //" derivedtbl_1.ОписаниеПолучателя = derivedtbl_5.ОписаниеПолучателя LEFT OUTER JOIN " +
         //    " (SELECT     Карточка_3.id_корреспондента, Корреспонденты_3.ОписаниеКорреспондента, COUNT(Получатели_3.ОписаниеПолучателя) " +
         //                             " AS КоличествоВходДокументов, Получатели_3.ОписаниеПолучателя " +
         //      " FROM          dbo.Карточка AS Карточка_3 LEFT OUTER JOIN " +
         //                             " dbo.Корреспонденты AS Корреспонденты_3 ON  " +
         //                             " Карточка_3.id_корреспондента = Корреспонденты_3.id_корреспондента LEFT OUTER JOIN " +
         //                             " dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_3 ON  " +
         //                             " ПолучателДокументовУправление_3.idКарточки = Карточка_3.id_карточки INNER JOIN " +
         //                             " dbo.Получатели AS Получатели_3 ON Получатели_3.id_получателя = ПолучателДокументовУправление_3.idПолучатель " +
         //      "WHERE      (Карточка_3.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_3.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
         //      " GROUP BY Корреспонденты_3.ОписаниеКорреспондента, Карточка_3.id_корреспондента, Получатели_3.ОписаниеПолучателя, " +
         //                            " Карточка_3.idВидПоступленияДокумента " +
         //      " HAVING      (Карточка_3.idВидПоступленияДокумента = 3)) AS derivedtbl_4 ON  " +
         //" derivedtbl_1.id_корреспондента = derivedtbl_4.id_корреспондента AND " +
         //" derivedtbl_1.ОписаниеПолучателя = derivedtbl_4.ОписаниеПолучателя LEFT OUTER JOIN " +
         //    " (SELECT     Карточка_2.id_корреспондента, Корреспонденты_2.ОписаниеКорреспондента, COUNT(Получатели_2.ОписаниеПолучателя) " +
         //                             " AS КоличествоВходДокументов, Получатели_2.ОписаниеПолучателя " +
         //      " FROM          dbo.Карточка AS Карточка_2 LEFT OUTER JOIN " +
         //                             " dbo.Корреспонденты AS Корреспонденты_2 ON " +
         //                             " Карточка_2.id_корреспондента = Корреспонденты_2.id_корреспондента LEFT OUTER JOIN " +
         //                             " dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_2 ON " +
         //                             " ПолучателДокументовУправление_2.idКарточки = Карточка_2.id_карточки INNER JOIN " +
         //                             " dbo.Получатели AS Получатели_2 ON Получатели_2.id_получателя = ПолучателДокументовУправление_2.idПолучатель" +
         //      " WHERE      (Карточка_2.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_2.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "') " +
         //     " GROUP BY Корреспонденты_2.ОписаниеКорреспондента, Карточка_2.id_корреспондента, Получатели_2.ОписаниеПолучателя, " +
         //                             " Карточка_2.idВидПоступленияДокумента " +
         //      " HAVING      (Карточка_2.idВидПоступленияДокумента = 2)) AS derivedtbl_3 ON  " +
         //" derivedtbl_1.id_корреспондента = derivedtbl_3.id_корреспондента AND " +
         //"derivedtbl_1.ОписаниеПолучателя = derivedtbl_3.ОписаниеПолучателя LEFT OUTER JOIN " +
         //    " (SELECT     Карточка_1.id_корреспондента, Корреспонденты_1.ОписаниеКорреспондента, COUNT(Получатели_1.ОписаниеПолучателя)  " +
         //                             " AS КоличествоВходДокументов, Получатели_1.ОписаниеПолучателя " +
         //      "FROM          dbo.Карточка AS Карточка_1 LEFT OUTER JOIN " +
         //                             " dbo.Корреспонденты AS Корреспонденты_1 ON " +
         //                             "Карточка_1.id_корреспондента = Корреспонденты_1.id_корреспондента LEFT OUTER JOIN " +
         //                             "dbo.ПолучателДокументовУправление AS ПолучателДокументовУправление_1 ON  " +
         //                             " ПолучателДокументовУправление_1.idКарточки = Карточка_1.id_карточки INNER JOIN " +
         //                            " dbo.Получатели AS Получатели_1 ON Получатели_1.id_получателя = ПолучателДокументовУправление_1.idПолучатель " +
         //      " WHERE      (Карточка_1.ДатаПоступ >= '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "' and Карточка_1.ДатаПоступ <= '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "' )" +
         //      " GROUP BY Корреспонденты_1.ОписаниеКорреспондента, Карточка_1.id_корреспондента, Получатели_1.ОписаниеПолучателя,  " +
         //                             " Карточка_1.idВидПоступленияДокумента " +
         //      " HAVING      (Карточка_1.idВидПоступленияДокумента = 1)) AS derivedtbl_2 ON  " +
         //" derivedtbl_1.ОписаниеПолучателя = derivedtbl_2.ОписаниеПолучателя AND derivedtbl_1.id_корреспондента = derivedtbl_2.id_корреспондента " +
         //"  where LOWER(LTRIM(RTRIM(derivedtbl_1.ОписаниеПолучателя))) = '"+ fio.Trim().ToLower() +"' " +
         //" group by derivedtbl_1.[ОписаниеПолучателя] " +
         //" order by derivedtbl_1.ОписаниеПолучателя asc ";

         //   GetDataTable getData = new GetDataTable(query);
            //   return getData.DataTable();
            #endregion

            return dTab;
        }

        /// <summary>
        /// Возвращает итого по корреспондентам.
        /// </summary>
        /// <returns></returns>
        private DataTable SelectCountCorrespondent()
        {
            string query = " SELECT     TOP (100) PERCENT Корреспонденты_2.id_корреспондента, Корреспонденты_2.ОписаниеКорреспондента, derivedtbl_1.все, derivedtbl_1.бумага, " +
                      "derivedtbl_1.мыло, derivedtbl_1.вип, derivedtbl_1.факс " +
                        "FROM         (SELECT     derivedtbl_1_6.Expr1 AS все, derivedtbl_1_6.id_корреспондента, derivedtbl_2.Expr1 AS бумага, derivedtbl_3.Expr1 AS мыло,  " +
                                              "derivedtbl_4.Expr1 AS вип, derivedtbl_5.Expr1 AS факс " +
                       "FROM          (SELECT     SUM(Expr1) AS Expr1, id_корреспондента " +
                                               "FROM          (SELECT     TOP (100) PERCENT COUNT(dbo.Карточка.id_карточки) AS Expr1, dbo.Корреспонденты.id_корреспондента " +
                                                                       "FROM          dbo.Карточка LEFT OUTER JOIN " +
                                                                                              "dbo.Корреспонденты ON dbo.Карточка.id_корреспондента = dbo.Корреспонденты.id_корреспондента " +
                                                                       "GROUP BY dbo.Карточка.ДатаПоступ, dbo.Карточка.idВидПоступленияДокумента,  "+
                                                                                              "dbo.Корреспонденты.ОписаниеКорреспондента, dbo.Корреспонденты.id_корреспондента " +
                                                                       "HAVING      (dbo.Карточка.ДатаПоступ >= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                              "(dbo.Карточка.ДатаПоступ <= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                       "ORDER BY dbo.Корреспонденты.ОписаниеКорреспондента) AS derivedtbl_1_1 " +
                                               "GROUP BY id_корреспондента) AS derivedtbl_1_6 LEFT OUTER JOIN " +
                                                  "(SELECT     id_корреспондента AS Expr2, SUM(Expr1) AS Expr1 " +
                                                    "FROM          (SELECT     TOP (100) PERCENT COUNT(Карточка_1.id_карточки) AS Expr1, Корреспонденты_1.id_корреспондента "+
                                                                            "FROM          dbo.Карточка AS Карточка_1 LEFT OUTER JOIN " +
                                                                                                   "dbo.Корреспонденты AS Корреспонденты_1 ON  " +
                                                                                                   "Карточка_1.id_корреспондента = Корреспонденты_1.id_корреспондента " +
                                                                            "WHERE      (Карточка_1.idВидПоступленияДокумента = 1) " +
                                                                            "GROUP BY Карточка_1.ДатаПоступ, Карточка_1.idВидПоступленияДокумента,  " +
                                                                                                   "Корреспонденты_1.ОписаниеКорреспондента, Корреспонденты_1.id_корреспондента " +
                                                                            "HAVING      (Карточка_1.ДатаПоступ >= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                                   "(Карточка_1.ДатаПоступ <= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                            "ORDER BY Корреспонденты_1.ОписаниеКорреспондента) AS derivedtbl_1_5 " +
                                                    "GROUP BY id_корреспондента) AS derivedtbl_2 ON derivedtbl_1_6.id_корреспондента = derivedtbl_2.Expr2 LEFT OUTER JOIN " +
                                                  "(SELECT     id_корреспондента AS Expr2, SUM(Expr1) AS Expr1 " +
                                                    "FROM          (SELECT     TOP (100) PERCENT COUNT(Карточка_1.id_карточки) AS Expr1, Корреспонденты_1.id_корреспондента " +
                                                                           " FROM          dbo.Карточка AS Карточка_1 LEFT OUTER JOIN " +
                                                                                                   "dbo.Корреспонденты AS Корреспонденты_1 ON  " +
                                                                                                   "Карточка_1.id_корреспондента = Корреспонденты_1.id_корреспондента " +
                                                                            "WHERE      (Карточка_1.idВидПоступленияДокумента = 4) " +
                                                                            "GROUP BY Карточка_1.ДатаПоступ, Карточка_1.idВидПоступленияДокумента,  " +
                                                                                                   "Корреспонденты_1.ОписаниеКорреспондента, Корреспонденты_1.id_корреспондента " +
                                                                            "HAVING      (Карточка_1.ДатаПоступ >= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                                   "(Карточка_1.ДатаПоступ <= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                            " ORDER BY Корреспонденты_1.ОписаниеКорреспондента) AS derivedtbl_1_4 " +
                                                    " GROUP BY id_корреспондента) AS derivedtbl_5 ON derivedtbl_1_6.id_корреспондента = derivedtbl_5.Expr2 LEFT OUTER JOIN " +
                                                  "(SELECT     id_корреспондента AS Expr2, SUM(Expr1) AS Expr1 " +
                                                    "FROM          (SELECT     TOP (100) PERCENT COUNT(Карточка_1.id_карточки) AS Expr1, Корреспонденты_1.id_корреспондента " +
                                                                            "FROM          dbo.Карточка AS Карточка_1 LEFT OUTER JOIN " +
                                                                                                   "dbo.Корреспонденты AS Корреспонденты_1 ON  " +
                                                                                                   "Карточка_1.id_корреспондента = Корреспонденты_1.id_корреспондента " +
                                                                            "WHERE      (Карточка_1.idВидПоступленияДокумента = 3) " +
                                                                            "GROUP BY Карточка_1.ДатаПоступ, Карточка_1.idВидПоступленияДокумента,  " +
                                                                                                   "Корреспонденты_1.ОписаниеКорреспондента, Корреспонденты_1.id_корреспондента " +
                                                                            "HAVING      (Карточка_1.ДатаПоступ >= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                                   "(Карточка_1.ДатаПоступ <= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                            "ORDER BY Корреспонденты_1.ОписаниеКорреспондента) AS derivedtbl_1_3 " +
                                                    "GROUP BY id_корреспондента) AS derivedtbl_4 ON derivedtbl_1_6.id_корреспондента = derivedtbl_4.Expr2 LEFT OUTER JOIN " +
                                                  " (SELECT     id_корреспондента AS Expr2, SUM(Expr1) AS Expr1 " +
                                                    " FROM          (SELECT     TOP (100) PERCENT COUNT(Карточка_1.id_карточки) AS Expr1, Корреспонденты_1.id_корреспондента " +
                                                                            "FROM          dbo.Карточка AS Карточка_1 LEFT OUTER JOIN " +
                                                                                                   "dbo.Корреспонденты AS Корреспонденты_1 ON  " +
                                                                                                   "Карточка_1.id_корреспондента = Корреспонденты_1.id_корреспондента " +
                                                                            "WHERE      (Карточка_1.idВидПоступленияДокумента = 2) " +
                                                                            "GROUP BY Карточка_1.ДатаПоступ, Карточка_1.idВидПоступленияДокумента,  " +
                                                                                                   "Корреспонденты_1.ОписаниеКорреспондента, Корреспонденты_1.id_корреспондента " +
                                                                            "HAVING      (Карточка_1.ДатаПоступ >= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataStart.ToShortDateString()) + "', 102)) AND  " +
                                                                                                   "(Карточка_1.ДатаПоступ <= CONVERT(DATETIME, '" + ДатаSQL.Дата(this.ДиапазонДат.DataEnd.ToShortDateString()) + "', 102)) " +
                                                                            "ORDER BY Корреспонденты_1.ОписаниеКорреспондента) AS derivedtbl_1_2 " +
                                                    "GROUP BY id_корреспондента) AS derivedtbl_3 ON derivedtbl_1_6.id_корреспондента = derivedtbl_3.Expr2)  " +
                      " AS derivedtbl_1 INNER JOIN " + 
                      " dbo.Корреспонденты AS Корреспонденты_2 ON derivedtbl_1.id_корреспондента = Корреспонденты_2.id_корреспондента " +
"ORDER BY Корреспонденты_2.id_корреспондента ";

            GetDataTable getData = new GetDataTable(query);
            return getData.DataTable();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {

            //List<StatisticDocInput> list = new List<StatisticDocInput>();

            //StatisticDocInput itemHead = new StatisticDocInput();
            //// Шапка отчета.
            ////СтатистикаВходИсполнителей itemHead = new СтатистикаВходИсполнителей();
            //itemHead.Num = "№ п.п.";
            //itemHead.НаименованиеКорреспондента = "Наименование корреспондента";
            //itemHead.КолвоВходКорреспонденции = "Количество исходящих документов";
            //itemHead.БумажныйНоситель = "Бумажный носитель";
            //itemHead.VipNet = "VipNet";
            //itemHead.Fax = "факс";
            //itemHead.Email = "e-mail";
            //itemHead.Исполнитель = "Исполнитель";

            //list.Add(itemHead);

            //list.AddRange(this.listPrint);

            List<StatisticDocInput> list = this.listPrint;

            string caption = "Статистика по входящей корреспонденции за период с " + rd.DataStart.ToShortDateString() + " по " + rd.DataEnd.ToShortDateString();
            ReportСтатистикаВходящихДокументов printDate = new ReportСтатистикаВходящихДокументов(caption);
            printDate.SetDate = list;

            PrintReport printPaper = new PrintReport();
            printPaper.SetCommand(printDate);
            printPaper.Execute();

            //// Пишем шапку.
            //StatisticDocInput head = new StatisticDocInput();
            //head.Num = "№ п.п.";
            //head.НаименованиеКорреспондента = "Наименование корреспондента";
            //head.КолвоВходКорреспонденции = "Количество входящих документов";
            //head.БумажныйНоситель = "Бумажный носитель шт.";
            //head.Email = "Электронная почта шт.";
            //head.VipNet = "VipNet шт.";
            //head.Fax = "Факс шт.";
            //head.Исполнитель = "Исполнитель";

            //list.Add(head);

            ////// Откроем документ из шаблона.
            //////создадим документ WORD
            ////string fName = "СтатистикаВходящейКорреспонденции";

            ////// Очисмтим папку Документы.
            ////DirectoryInfo dirInfo = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + @"\Документы\");

            ////foreach (FileInfo file in dirInfo.GetFiles())
            ////{
            ////    file.Delete();
            ////}

            //////Скопируем шаблон в папку Документы
            //////try
            //////{
            ////    FileInfo fn = new FileInfo(System.Windows.Forms.Application.StartupPath + @"\Шаблон\Статистика по входящей корреспонденции.doc");
            ////    fn.CopyTo(System.Windows.Forms.Application.StartupPath + @"\Документы\" + fName + ".doc", true);
            //////}
            //////catch (Exception ex)
            //////{
            //////    MessageBox.Show("Документ уже открыт.\n" + ex.Message, "Ошибка");
            //////    return;
            //////}

            ////string filName = System.Windows.Forms.Application.StartupPath + @"\Документы\" + fName + ".doc";

            //////Создаём новый Word.Application
            ////Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //////Загружаем документ
            ////Microsoft.Office.Interop.Word.Document doc = null;

            ////object fileName = filName;
            ////object falseValue = false;
            ////object trueValue = true;
            ////object missing = Type.Missing;

            ////doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ////ref missing, ref missing, ref missing, ref missing, ref missing,
            ////ref missing, ref missing, ref missing, ref missing, ref missing,
            ////ref missing, ref missing, ref missing);

            //////Дата начальная.
            ////object wdrepl2 = WdReplace.wdReplaceAll;
            //////object searchtxt = "GreetingLine";
            ////object searchtxt2 = "datestart";
            ////object newtxt2 = (object)this.ДиапазонДат.DataStart.ToShortDateString();
            //////object frwd = true;
            ////object frwd2 = false;
            ////doc.Content.Find.Execute(ref searchtxt2, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd2, ref missing, ref missing, ref newtxt2, ref wdrepl2, ref missing, ref missing,
            ////ref missing, ref missing);

            //////Дата конечная.
            ////object wdrepl3 = WdReplace.wdReplaceAll;
            //////object searchtxt = "GreetingLine";
            ////object searchtxt3 = "dateend";
            ////object newtxt3 = (object)this.ДиапазонДат.DataEnd.ToShortDateString();
            //////object frwd = true;
            ////object frwd3 = false;
            ////doc.Content.Find.Execute(ref searchtxt3, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd3, ref missing, ref missing, ref newtxt3, ref wdrepl3, ref missing, ref missing,
            ////ref missing, ref missing);

            //////Вставить таблицу
            ////object bookNaziv = "таблица";
            ////Range wrdRng = doc.Bookmarks.get_Item(ref  bookNaziv).Range;

            ////object behavior = Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord8TableBehavior;
            ////object autobehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow;


            ////Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(wrdRng, 1, 8, ref behavior, ref autobehavior);
            ////table.Range.ParagraphFormat.SpaceAfter = 8;

            //////выставим ширину столбцов
            ////table.Columns[1].Width = 80;
            ////table.Columns[2].Width = 260;
            ////table.Columns[3].Width = 80;
            ////table.Columns[4].Width = 60;
            ////table.Columns[5].Width = 60;
            ////table.Columns[6].Width = 60;
            ////table.Columns[7].Width = 60;
            ////table.Columns[8].Width = 80;
            ////table.Borders.Enable = 1; // Рамка - сплошная линия
            ////table.Range.Font.Name = "Times New Roman";
            ////table.Range.Font.Size = 10;
            //////счётчик строк
            ////int i = 1;

                      
            //////// Добавим шапку в таблицу.
            //////table.Cell(i, 1).Range.Text = "№ п.п.";
            //////table.Cell(i, 2).Range.Text = "Наименование корреспондента";

            //////table.Cell(i, 3).Range.Text = "Количество входящих документов";
            //////table.Cell(i, 4).Range.Text = "Бумажный носитель шт.";

            //////table.Cell(i, 5).Range.Text = "Электронная почта шт.";
            //////table.Cell(i, 6).Range.Text = "VipNet шт.";

            //////table.Cell(i, 7).Range.Text = "Факс шт.";
            //////table.Cell(i, 8).Range.Text = "Исполнитель";



            ////// Добавим шапку в таблицу.
            ////table.Cell(1, 1).Range.Text = "№ п.п.";
            ////table.Cell(1, 2).Range.Text = "Наименование корреспондента";

            ////table.Cell(1, 3).Range.Text = "Количество входящих документов";
            ////table.Cell(1, 4).Range.Text = "Бумажный носитель шт.";

            ////table.Cell(1, 5).Range.Text = "Электронная почта шт.";
            ////table.Cell(1, 6).Range.Text = "VipNet шт.";

            ////table.Cell(1, 7).Range.Text = "Факс шт.";
            ////table.Cell(1, 8).Range.Text = "Исполнитель";

            //////doc.Words.Count.ToString();
            ////Object beforeRow1 = Type.Missing;
            ////table.Rows.Add(ref beforeRow1);

            //////i++;
            ////i = 2;

            //////запишем данные в таблицу
            ////foreach (StatisticDocInput item in list)
            ////{
            ////    table.Cell(i, 1).Range.Text = item.Num;
            ////    table.Cell(i, 2).Range.Text = item.НаименованиеКорреспондента.Trim();

            ////    if (item.КолвоВходКорреспонденции > 0)
            ////    {
            ////        table.Cell(i, 3).Range.Text = item.КолвоВходКорреспонденции.ToString().Trim();
            ////    }
            ////    else
            ////    {
            ////        table.Cell(i, 3).Range.Text = "";
            ////    }

            ////    if (item.БумажныйНоситель > 0)
            ////    {
            ////        table.Cell(i, 4).Range.Text = item.БумажныйНоситель.ToString().Trim();
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
            ////    table.Cell(i, 8).Range.Text = item.Исполнитель.Trim();

            ////    //doc.Words.Count.ToString();
            ////    Object beforeRow2 = Type.Missing;
            ////    table.Rows.Add(ref beforeRow2);

            ////    i++;
            ////}
            ////table.Rows[i].Delete();


            ////// Отобразим документ.
            ////app.Visible = true;
            
            //DataTable tPerson =  SelectPerson();

            //List<СтатистикаВходИсполнителей> list = new List<СтатистикаВходИсполнителей>();

            //foreach (DataRow row in tPerson.Rows)
            //{
            //    // Получим перечень строк для текущего пользователя.
            //    DataRow[] rows = rez.Select("ОписаниеПолучателя = '" + row["ОписаниеПолучателя"].ToString().Trim() + "' ");

            //    int iCount = 1;
            //    foreach (DataRow r in rows)
            //    {
            //        СтатистикаВходИсполнителей itm = new СтатистикаВходИсполнителей();
            //        itm.НомерПП = iCount.ToString();
            //        itm.НаименованиеКорреспондента = r["ОписаниеКорреспондента"].ToString().Trim();

            //        itm.КоличесвтоВходДокументов = r["КоличествоВходящихДокументов"].ToString().Trim();
            //        itm.БумажныйНоститель = r["Бумага"].ToString().Trim();
            //        itm.EMail = r["e-mail"].ToString().Trim();
            //        itm.VipNet = r["VipNet"].ToString().Trim();
            //        itm.Fax = r["факс"].ToString().Trim();
            //        itm.Исполнитель = r["ОписаниеПолучателя"].ToString().Trim();

            //        list.Add(itm);

            //        iCount++;
            //    }

            //    DataRow rCount = SelectCountPerson(row["ОписаниеПолучателя"].ToString().Trim()).Rows[0];

            //    // Зпапишем строку итого для исполнителя.
            //    СтатистикаВходИсполнителей itCount = new СтатистикаВходИсполнителей();
            //    itCount.НомерПП = "Итого по исполнителю " + rCount["ОписаниеПолучателя"].ToString().Trim();
            //    itCount.Исполнитель = "----------";
            //    itCount.КоличесвтоВходДокументов = rCount["КолВходДокументов"].ToString().Trim();
            //    itCount.БумажныйНоститель = rCount["бумага"].ToString().Trim();
            //    itCount.EMail = rCount["мыло"].ToString().Trim();
            //    itCount.VipNet = rCount["вип"].ToString().Trim();
            //    itCount.Fax = rCount["факс"].ToString().Trim();
            //    itCount.Исполнитель = rCount["ОписаниеПолучателя"].ToString().Trim();

            //    list.Add(itCount);
            //}

            //int iCount2 = 1;

            //int countВсеДокументы = 0;
            //int countВсеБумага = 0;
            //int countВсеМыло = 0;
            //int countВсеВип = 0;
            //int countВсеФакс = 0;

            //DataTable dtRowsCorrespondent = SelectCountCorrespondent();
            //// Запишем итого по по комитету по корреспондентам.
            //foreach (DataRow r in dtRowsCorrespondent.Rows)
            //{
            //    СтатистикаВходИсполнителей itm = new СтатистикаВходИсполнителей();
            //    itm.НомерПП = iCount2.ToString();
            //    itm.НаименованиеКорреспондента = r["ОписаниеКорреспондента"].ToString().Trim();

            //    itm.КоличесвтоВходДокументов = r["все"].ToString().Trim();

            //    if (DBNull.Value != r["все"])
            //    {
            //        countВсеДокументы += Convert.ToInt32(r["все"]);
            //    }

            //    itm.БумажныйНоститель = r["бумага"].ToString().Trim();

            //    if (DBNull.Value != r["бумага"])
            //    {
            //        countВсеБумага += Convert.ToInt32(r["бумага"]);
            //    }

            //    itm.EMail = r["мыло"].ToString().Trim();

            //    if (DBNull.Value != r["мыло"])
            //    {
            //        countВсеМыло += Convert.ToInt32(r["мыло"]);
            //    }

            //    itm.VipNet = r["вип"].ToString().Trim();

            //    if (DBNull.Value != r["вип"])
            //    {
            //        countВсеВип += Convert.ToInt32(r["вип"]);
            //    }

            //    itm.Fax = r["факс"].ToString().Trim();

            //    if (DBNull.Value != r["факс"])
            //    {
            //        countВсеФакс += Convert.ToInt32(r["факс"]);
            //    }

            //    itm.Исполнитель = "-----";

            //    list.Add(itm);

            //    iCount2++;
            //}

            //СтатистикаВходИсполнителей count = new СтатистикаВходИсполнителей();
            //count.НомерПП = "Итого в целом по комитету";
            //count.НаименованиеКорреспондента = "-----";
            //count.КоличесвтоВходДокументов = countВсеДокументы.ToString();
            //count.БумажныйНоститель = countВсеБумага.ToString();
            //count.EMail = countВсеМыло.ToString();
            //count.VipNet = countВсеВип.ToString();
            //count.Fax = countВсеФакс.ToString();

            //list.Add(count);
            

            //List<СтатистикаВходИсполнителей> listTest = list;



            //ExcelPrint excel = new ExcelPrint(" Статистика по входящей корреспонденции c " + this.ДиапазонДат.DataStart.ToShortDateString() + " по " + this.ДиапазонДат.DataEnd.ToShortDateString() + " по КСЗН г. Саратова");
            //excel.PrintСтатистикаВходящейКорреспонденции(list);

            // Передадим всё жто в Excel.
            this.Close();

        }
    }
}