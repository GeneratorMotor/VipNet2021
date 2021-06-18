using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
//using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;



using RegKor.Classess;


namespace RegKor
{
    public partial class FormSelectDatePerson : Form
    {
        //������ Excel
        private Microsoft.Office.Interop.Excel.Application ObjExcel;

        //������ ������ excel ����
        private Microsoft.Office.Interop.Excel.Workbooks ObjWorkBooks;

        //������ excel �����
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;

        //������ excel ����
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

        public FormSelectDatePerson()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string beginDate = this.dateTimePicker1.Value.ToShortDateString();
            string endDate = this.dateTimePicker2.Value.ToShortDateString();

            string fName = beginDate + "�������������������������������������" + endDate;


            // ���������� ������ ��� ���������� ������.
            //string query = "SELECT [id_��������] " +
            //               ",[����������������������] " +
            //               ",[����������] " +
            //               ",[�������������������������������] " +
            //               ",[������������������������] " +
            //               ",[���������������������] " +
            //               ",[�������] " +
            //               "FROM [View������] " +
            //               "where [����������] >= '" + beginDate + "' and [����������] <= '" + endDate + "' ";

            ������������ connect = new ������������();
            string sConnect = connect.�����������������();

            // ������ ������ ��� ������.
            //List<���������������> list = new List<���������������>();

            // ���������� ��� �������� ������ ��� �������.
            Dictionary<string, ���������������> dictionary = new Dictionary<string, ���������������>();


            using (SqlConnection con = new SqlConnection(sConnect))
            {
                con.Open();
                string query = "select * from dbo.View�����������������������������2 " +
                               "where [����] >= '" + beginDate + "' and [����] <= '" + endDate + "' ";

                DataSet ds = new DataSet();

                 // �������� DataSet.
                SqlDataAdapter da = new SqlDataAdapter(query, con);
                da.Fill(ds, "View�����������������������������");

                // ������� ������ �� ������� ������� ������������ 
                System.Data.DataTable tab1 = ds.Tables["View�����������������������������"];

                // ������� ����� ��� ������������ �������.
                //int iKey = 1;
                int iKey = 0;

                foreach (DataRow row in tab1.Rows)
                {
                    iKey = Convert.ToInt32(row["id_��������"]);

                    if (Convert.ToInt32(row["id_��������"]) == 44440)
                    {
                        string iTest = "test";
                    }

                    ��������������� ������ = new ���������������();
                    
#region ������ ��������� (���� �� ������� ����� ����������� � ��� ���������).
                    //// ���� ������ �� ������������ (������� id ���������� ���������).
                    //if (row["id_������������������"] != DBNull.Value)
                    //{

                    //    // ����.
                    //    string key = row["�������������"].ToString().Trim();

                    //    ������.���������������������� = row["����������������������"].ToString().Trim();

                    //    //������.����������������� = row["�����������������"].ToString().Trim();

                    //    ������.�������������� = row["��������������"].ToString().Trim();

                    //    ������.������������ = Convert.ToDateTime(row["����"]).ToShortDateString();

                    //    //������.������������� = row["�������������"].ToString().Trim();

                    //    // �������� ���� �� ��� ������ � ��������� ������� [����������������������������������] �������������� id ��������.
                    //    string numPoust = string.Empty;
                    //    string about = string.Empty;
                    //    numPoust = InputNumberPoust(con, ds, Convert.ToInt32(row["id_��������"]), out about);
                    //    if (numPoust != "")
                    //    {
                    //        ������.������������� = numPoust.Trim();
                    //        ������.����������������� = about.Trim();
                    //    }
                    //    else
                    //    {
                    //        ������.������������� = row["�������������"].ToString().Trim();
                    //        ������.����������������� = row["�����������������"].ToString().Trim();
                    //    }

                    //    try
                    //    {

                    //        // ������� �������� ��������.
                    //        ��������������� ������2 = ������;

                    //        // ������� � ���������� ����� �������.
                    //        dictionary.Add(key, ������);


                    //    }
                    //    catch
                    //    {

                    //        //����� ������ View 2 ������� � ��� � �������� ����� 

                    //        //��������������� ������Cath = new ���������������();

                    //        //����� id �������� � �������� �������

                    //        /*
                    //         * ���� ����� ���� ��� ���������� � ����������, ����� ������ � ������� ������ ��������� 
                    //         * � �������������� � ������ �������.
                    //        */

                    //        ��������������� ������Cath = new ���������������();

                    //        // ������� ������� ������ �� ����������.
                    //        string strTime = dictionary[key].����������������������.Trim();
                    //        ������Cath.����������������� = dictionary[key].�����������������.Trim();
                    //        ������Cath.�������������� = dictionary[key].��������������.Trim();
                    //        ������Cath.������������ = dictionary[key].������������.Trim();
                    //        ������Cath.������������� = dictionary[key].�������������.Trim();


                    //        string[] numN = key.Split('/');
                    //        string ���� = ����SQL.����(DateTime.Now.Year.ToString()) + "0101";


                    //        string[] ��������������� = row["��������������������������"].ToString().Split('/');

                    //        // ������ ������������ ����������� ���� ���������� ������.

                    //        string quer = "declare @id_���������������� int " +
                    //                     "set @id_���������������� = 0 " +
                    //                     "select @id_���������������� = id_���������������� from dbo.���������������������������������� " +
                    //                     "where id_����������������� in ( " +
                    //                     "select id_�������� from ����������������� " +
                    //                     "where ��������������� = '" + ���������������[1] + "') " +
                    //                     "if(@id_���������������� = 0) " +
                    //                     "begin  " +
                    //                     "select distinct ���������������������� from �������������� " +
                    //                     "where id_�������������� in ( " +
                    //                     "select id_�������� from ����������������� " +
                    //            //"where ��������������� = '6125') " +
                    //                     "where ��������������� = '" + ���������������[1] + "' and ���� >= '20140101') " +
                    //                     "end " +
                    //                     "else " +
                    //                     "begin  " +
                    //                     "select ���������������������� from dbo.��������������  " +
                    //                     "where id_�������������� in  " +
                    //                     "( select id_�������������� from ��������  " +
                    //                     "where id_�������� in  " +
                    //                     "( select id_���������������� from dbo.����������������������������������  " +
                    //                     "where id_����������������� in  " +
                    //                     "( select id_�������� from dbo.View�����������������������������2  " +
                    //                     "where ������������� = '" + row["�������������"].ToString().Trim() + "'))) " +
                    //                     "end ";

                    //        DataSet ds2 = new DataSet();

                    //        SqlDataAdapter da2 = new SqlDataAdapter(quer, con);
                    //        da2.Fill(ds2, "��������");

                    //        // ������� ������ �� ������� ������� ������������ 
                    //        System.Data.DataTable tab22 = ds2.Tables["��������"];

                    //        // ���� ���������� ������� � ������� ��������������� ������ 1
                    //        if (tab22.Rows.Count > 1)
                    //        {
                    //            StringBuilder buildCorr = new StringBuilder();

                    //            // ������� �������� �������������� ������� ��� ��� � 1 ��� ������� � ����������.
                    //            buildCorr.Append(strTime + ",");

                    //            foreach (DataRow r in tab22.Rows)
                    //            {
                    //                string ���� = r["����������������������"].ToString().Trim();
                    //                buildCorr.Append(���� + ",");
                    //            }

                    //            // ����� ��������� �������.
                    //            int leng = buildCorr.Length;

                    //            string ������������� = string.Empty;
                    //            ������������� = buildCorr.Remove(leng - 1, 1).ToString().Trim();

                    //            ������Cath.���������������������� = �������������;

                    //            // ������� ��������� ������� �� �����������, ����� � ��������� ������ �� ������ �� ������ ������.
                    //            tab22.Clear();

                    //        }
                    //        else
                    //        {
                    //            ������Cath.���������������������� = row["����������������������"].ToString().Trim() + ", \n" + strTime.Trim();

                    //            string querSelect = "select CONVERT(nvarchar, �������) + N'/' + RTRIM(LTRIM(CONVERT(nvarchar, " +
                    //                                "���������))) AS '�����'  from �������� " +
                    //                                "where id_�������� in ( " +
                    //                                "select id_���������������� from dbo.���������������������������������� " +
                    //                                "where id_����������������� = " + Convert.ToInt32(row["id_��������"]) + ")";

                    //            SqlDataAdapter daM = new SqlDataAdapter(querSelect, con);
                    //            daM.Fill(ds, "��������");

                    //            // ������� ������ �� ������� ������� ������������ 
                    //            System.Data.DataTable tabM = ds.Tables["��������"];

                    //            // ������ �������� ��������.
                    //            StringBuilder buld = new StringBuilder();

                    //            foreach (DataRow r in tabM.Rows)
                    //            {
                    //                buld.Append(r["�����"].ToString().Trim() + ", \n");
                    //            }

                    //            // ������ ��������� �������.
                    //            int countChar = buld.Length;
                    //            string numbers = string.Empty;

                    //            if (countChar > 0)
                    //            {
                    //                numbers = buld.Remove(countChar - 3, 3).ToString();
                    //                ������Cath.������������� = numbers.ToString().Trim();
                    //            }
                    //        }

                    //        // ������ ������.
                    //        //dictionary.Remove(key);

                    //        // ������� ��������� ������.
                    //        dictionary.Add(key + 1, ������Cath);
                    //    }
                    //}
#endregion

                    // �������� ������������ ������ ��� ���.
                    if (row["id_������������������"] != DBNull.Value)
                    {
                        // ������� ���� ��� ���������, � �������� ����� ���������� id �������� ���������.
                        string key = row["id_��������"].ToString().Trim();// row["�������������"].ToString().Trim();

                        ������.���������������������� = row["����������������������"].ToString().Trim();

                        ������.�������������� = row["��������������"].ToString().Trim();

                        ������.������������ = Convert.ToDateTime(row["����"]).ToShortDateString();

                        // �������� ���� �� ��� ������ � ��������� ������� [����������������������������������] �������������� id ��������.
                        string numPoust = string.Empty;
                        string about = string.Empty;
                        numPoust = InputNumberPoust(con, ds, Convert.ToInt32(row["id_��������"]), out about);
                        if (numPoust != "")
                        {
                            ������.������������� = numPoust.Trim();
                            ������.����������������� = about.Trim();
                        }
                        else
                        {
                            ������.������������� = row["�������������"].ToString().Trim();
                            ������.����������������� = row["�����������������"].ToString().Trim();
                        }

                        try
                        {
                            // ������� �������� ��������.
                            ��������������� ������2 = ������;

                            // ������� � ���������� ����� �������.
                            dictionary.Add(key, ������);
                        }
                        catch
                        {

                            ��������������� ������Cath = new ���������������();

                            // ������� ������� ������ �� ����������.
                            string strTime = dictionary[key].����������������������.Trim();
                            ������Cath.����������������� = dictionary[key].�����������������.Trim();
                            ������Cath.�������������� = dictionary[key].��������������.Trim();
                            ������Cath.������������ = dictionary[key].������������.Trim();
                            ������Cath.������������� = dictionary[key].�������������.Trim();

                            ������Cath.���������������������� = row["����������������������"].ToString().Trim() + ", \n" + strTime.Trim();

                            dictionary.Remove(key);

                            dictionary.Add(key, ������Cath);

                        }
                    }
                    else
                    {

                        ������.���������������������� = row["����������������������"].ToString().Trim();

                        // ���� ������ ������������.
                        ������.����������������� = row["����������������������"].ToString().Trim();
                        ������.�������������� = row["��������������������������"].ToString().Trim();
                        ������.������������ = Convert.ToDateTime(row["����"]).ToShortDateString();

                        // ������� ����� ����������.
                        string[] numArr = ������.��������������.Split('/');

                        if (Convert.ToInt32(row["id_��������"]) == 46095)
                        {
                            string test = "Test";
                        }

                        string queryOP = "select ����������������� from ����������������� " +
                                         "where id_����������������� in ( " +
                                         "select id_����������������� from ��������������������������������������� " +
                                         "where id_�������� = " + Convert.ToInt32(row["id_��������"]) + ") ";


                        SqlDataAdapter daOP = new SqlDataAdapter(queryOP, con);
                        daOP.Fill(ds, "�����������������");

                        // ������� ������ �� ������� ������� ������������ 
                        System.Data.DataTable tabOP = ds.Tables["�����������������"];

                        if (tabOP.Rows.Count != 0)
                        {
                            ������������ strOt = new ������������(tabOP);
                            string str���� = strOt.ConvertStringBuilder();

                            ������.������������� = str����;

                            // ������� �������.
                            ds.Tables["�����������������"].Clear();
                        }
                        else
                        {
                            ������.������������� = "������������";
                        }

                        // ����.
                        //string key = row["�������������"].ToString().Trim();

                        //iKey++; - ������ ����������.

                        //try
                        //{
                            // ������� � ���������� ����� �������.
                            dictionary.Add(iKey.ToString().Trim(), ������);
                        //}
                        //catch
                        //{
                        //    //string strTime = dictionary[iKey.ToString().Trim()].����������������������.Trim();
                        //    iKey++;
                        //    dictionary.Add(iKey.ToString().Trim(), ������);
                          
                        //}
                    }

                    //list.Add(������);
                }

            }

            // ���������� ��� ��������� ������ �������.
            int width1Column = 15;
            int width5Column = 20;
            int widthColumn = 50;
            int widthShortContColumn = 70;

            int ������������ = 90;
            int ������������2 = 50;

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

            //�����.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);

            //�������.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            // ��������� ��������� ���������� ������.
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;

            // ������� ������� � 55%.
            ObjWorkSheet.PageSetup.Zoom = 70;

            // ��������� ������� � ���� � � ����� = 0.
            ObjWorkSheet.PageSetup.LeftMargin = 0;
            ObjWorkSheet.PageSetup.RightMargin = 0;

            // ��������� ������ � ���� � � ������.
            ObjWorkSheet.PageSetup.TopMargin = 0;
            ObjWorkSheet.PageSetup.BottomMargin = 0;


            // �������� �� ������.
            ObjWorkSheet.PageSetup.CenterHorizontally = true;

            // ��������� �������.
            //ObjExcel.ActiveWindow.Zoom = 50;



            //������� �����
            //��������� ������
            ObjWorkSheet.get_Range("E1", "F1").Merge(Type.Missing);
            ObjWorkSheet.get_Range("E1", "F1").Font.Size = 12;
            ObjWorkSheet.get_Range("E1", "F1").Font.Bold = true;
            ObjWorkSheet.get_Range("E1", Type.Missing).Value2 = "������Ĩ�";

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("E1", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("E1", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // ������� ����� � ������ E2 F2. � ��������� ������ ������ 12, �� ������
            ObjWorkSheet.get_Range("E2", "F2").Merge(Type.Missing);
            ObjWorkSheet.get_Range("E2", "F2").Font.Size = 12;
            ObjWorkSheet.get_Range("E2", "F2").Font.Bold = false;
            ObjWorkSheet.get_Range("E2", Type.Missing).Value2 = "�������� �.�. ��������� \n��� �� \"����� ����������� ������� \" \n" +
                                                                "\n�� 09.01.2019 �. � 25";

            // ������� ������ �������� E � F.
            ObjWorkSheet.get_Range("E1", "E1").ColumnWidth = width5Column;
            ObjWorkSheet.get_Range("F1", "F1").ColumnWidth = width5Column;

            // ��������� ������ ������.
            ObjWorkSheet.get_Range("E2", "E2").RowHeight = ������������;
            ObjWorkSheet.get_Range("F2", "F2").RowHeight = ������������;

            // ������� ������ ��������.
            ObjWorkSheet.get_Range("E2", Type.Missing).HorizontalAlignment = Excel.Constants.xlLeft;
            ObjWorkSheet.get_Range("E2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ObjWorkSheet.get_Range("F1", Type.Missing).HorizontalAlignment = Excel.Constants.xlLeft;
            ObjWorkSheet.get_Range("F1", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // ������� �������� �������.
            ObjWorkSheet.get_Range("C5", "E5").Merge(Type.Missing);
            ObjWorkSheet.get_Range("C5", "E5").Font.Size = 12;
            ObjWorkSheet.get_Range("C5", "E5").Font.Bold = true;
            ObjWorkSheet.get_Range("C5", Type.Missing).Value2 = "������ ����� �������� ������������ ������ \n � "+ beginDate +" �� "+endDate+" ";

            // ������� ������ �������� E � F.
            ObjWorkSheet.get_Range("C1", "C1").ColumnWidth = width5Column;
            ObjWorkSheet.get_Range("D1", "D1").ColumnWidth = width5Column;
            ObjWorkSheet.get_Range("E1", "E1").ColumnWidth = width5Column;

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("C5", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("C5", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ObjWorkSheet.get_Range("C5", "C5").RowHeight = ������������2;

            // ������� ����� �������.
            ObjWorkSheet.get_Range("A7", "A7").Merge(Type.Missing);
            ObjWorkSheet.get_Range("A7", Type.Missing).Value2 = "� �/�";

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("A7", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("A7", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // �������� �������.
            Excel������ A7 = new Excel������();
            A7.�������������("A7", "A7", ObjWorkSheet);

            // ������� ����� �������.
            ObjWorkSheet.get_Range("B7", "B7").Merge(Type.Missing);
            ObjWorkSheet.get_Range("B7", Type.Missing).Value2 = "�������� � ������������� ����";

            ObjWorkSheet.get_Range("B7", "B7").ColumnWidth = widthColumn;

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("B7", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("B7", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // �������� �������.
            Excel������ B7 = new Excel������();
            B7.�������������("B7", "B7", ObjWorkSheet);
            ObjWorkSheet.get_Range("B7", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;

            string cell = "C7";
            ObjWorkSheet.get_Range(cell, cell).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cell, Type.Missing).Value2 = "������� ���������� ������� ��� \n������������ �������� ��";

            ObjWorkSheet.get_Range(cell, cell).ColumnWidth = widthShortContColumn;
            ObjWorkSheet.get_Range(cell, cell).RowHeight = ������������2;

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range(cell, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cell, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // �������� �������.
            Excel������ C7 = new Excel������();
            C7.�������������(cell, cell, ObjWorkSheet);
            ObjWorkSheet.get_Range(cell, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;

            // ������� D.
            string cellD = "D7";
            ObjWorkSheet.get_Range(cellD, cellD).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cellD, Type.Missing).Value2 = "������� � �������� ��� ������ � \n�������� ��";

            ObjWorkSheet.get_Range(cellD, cellD).ColumnWidth = width5Column;
            ObjWorkSheet.get_Range(cellD, cellD).RowHeight = ������������2;

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range(cellD, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cellD, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // �������� �������.
            Excel������ D7 = new Excel������();
            D7.�������������(cellD, cellD, ObjWorkSheet);
            ObjWorkSheet.get_Range(cellD, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;

            // ������� E.
            string cellE = "E7";
            ObjWorkSheet.get_Range(cellE, cellE).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cellE, Type.Missing).Value2 = "���� �������� (������ � \n��������)��";

            ObjWorkSheet.get_Range(cellE, cellE).ColumnWidth = width5Column;
            ObjWorkSheet.get_Range(cellE, cellE).RowHeight = ������������2;

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range(cellE, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cellE, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // �������� �������.
            Excel������ E7 = new Excel������();
            E7.�������������(cellE, cellE, ObjWorkSheet);
            ObjWorkSheet.get_Range(cellE, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;


            // ������� F.
            string cellF = "F7";
            ObjWorkSheet.get_Range(cellF, cellF).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cellF, Type.Missing).Value2 = "��������� �������� �� \n(����� �������)";

            ObjWorkSheet.get_Range(cellF, cellF).ColumnWidth = width5Column;
            ObjWorkSheet.get_Range(cellF, cellF).RowHeight = ������������2;

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range(cellF, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cellF, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // �������� �������.
            Excel������ F7 = new Excel������();
            F7.�������������(cellF, cellF, ObjWorkSheet);
            ObjWorkSheet.get_Range(cellF, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;

            // ����� ��������� ����� � 8 ������, ��� ��� ������ 7 ����� ������ ��� ����� �������.
            int iCount = 8;

            // ������� ��������� �����.
            int num = 1;

            // ������ ���������� ������� � ������ list.
            int countRowsReport = dictionary.Values.Count;// list.Count;

            // ��������� ������� �������.
            foreach (��������������� item in dictionary.Values)
            {
                // ������ �� ������� ��������.
                for (int i = 1; i <= 6; i++)
                {
                    // ������� ����� ������������� �������.
                    string exclB = Excel������.������������(i);

                    switch(i)
                    {
                        case 1:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = num.ToString().Trim();
                            

                            Excel������ excCel = new Excel������();
                            excCel.�������������(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


                            break;
                        case 2:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.����������������������.Trim();
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).WrapText = true;

                            Excel������ excCelB = new Excel������();
                            excCelB.�������������(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        case 3:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.�����������������.Trim();
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).WrapText = true;

                            Excel������ excCelC = new Excel������();
                            excCelC.�������������(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        case 4:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.��������������.Trim();

                            Excel������ excCelD = new Excel������();
                            excCelD.�������������(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        case 5:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.������������.Trim();

                            Excel������ excCelE = new Excel������();
                            excCelE.�������������(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        // ������ F
                        case 6:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.�������������.Trim();

                            Excel������ excCelF = new Excel������();
                            excCelF.�������������(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                    }
                }

                num++;
                iCount++;
            }

            // ������� ������������ ������� ����������� �����.

            // ��������� ����� ������ ��� ����� �������� ��� ������������ ������� ����������� �����, ��� countRowsReport - ���������� ����� � ������ � ������� , � 11 - 8 ����� � ����� + 3 ������ ��������� �� �������.
            int numRow = countRowsReport + 11;

            string cellUsrE = "D" + numRow.ToString();
            string cellUsrF = "F" + numRow.ToString();

            ObjWorkSheet.get_Range(cellUsrE, cellUsrF).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cellUsrE, cellUsrF).Font.Size = 10;
            ObjWorkSheet.get_Range(cellUsrE, cellUsrF).Font.Bold = false; ;
            ObjWorkSheet.get_Range(cellUsrE, Type.Missing).Value2 = "����������� " + MyAplicationIdentity.GetUses();

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range(cellUsrE, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cellUsrE, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // ������� �������� �� �����.
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;


            #region ������ ���������� WORD

            //������ ����� Word.Application
    //        Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();


    //            //��������� ��������
    //            Microsoft.Office.Interop.Word.Document doc = null;

    //            object fileName = filName;
    //            object falseValue = false;
    //            object trueValue = true;
    //            object missing = Type.Missing;
    //            object writePasswordDocument = "12A86Asd";

    //            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
    //ref missing, ref missing, ref missing, ref missing, ref writePasswordDocument,
    //ref missing, ref missing, ref missing, ref missing, ref trueValue,
    //ref missing, ref missing, ref missing);

    //            ////���� ������ ������.
    //            object wdrepl = WdReplace.wdReplaceAll;
    //            //object searchtxt = "GreetingLine";
    //            object searchtxt = "DATESTART";
    //            object newtxt = (object)beginDate;
    //            //object frwd = true;
    //            object frwd = false;
    //            doc.Content.Find.Execute(ref searchtxt, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd, ref missing, ref missing, ref newtxt, ref wdrepl, ref missing, ref missing,
    //            ref missing, ref missing);

    //            // ���� ��������� ������.
    //            object wdrepl2 = WdReplace.wdReplaceAll;
    //            //object searchtxt = "GreetingLine";
    //            object searchtxt2 = "DATEEND";
    //            object newtxt2 = (object)endDate;
    //            //object frwd = true;
    //            object frwd2 = false;
    //            doc.Content.Find.Execute(ref searchtxt2, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd2, ref missing, ref missing, ref newtxt2, ref wdrepl2, ref missing, ref missing,
    //            ref missing, ref missing);

    //            //�������� �������
    //            object bookNaziv = "�������";
    //            Range wrdRng = doc.Bookmarks.get_Item(ref  bookNaziv).Range;

    //            object behavior = Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord8TableBehavior;
    //            object autobehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow;


    //            Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(wrdRng, 1, 6, ref behavior, ref autobehavior);
    //            table.Range.ParagraphFormat.SpaceAfter = 11;

    //            table.Columns[1].Width = 40;
    //            table.Columns[2].Width = 150;
    //            table.Columns[3].Width = 150;
    //            table.Columns[4].Width = 150;
    //            table.Columns[5].Width = 120;
    //            table.Columns[6].Width = 120;

    //            table.Borders.Enable = 1; // ����� - �������� �����
    //            table.Range.Font.Name = "Times New Roman";
    //            table.Range.Font.Size = 9;

    //            // ������� ����� �������.
    //            table.Cell(1, 1).Range.Text = "� �/�";
    //            table.Cell(1, 2).Range.Text = "�������� � ������������� ����";
    //            table.Cell(1, 3).Range.Text = "������� ���������� ������� ��� ������������ �������� ��";
    //            table.Cell(1, 4).Range.Text = "������� � �������� ��� ������ � �������� ��";
    //            table.Cell(1, 5).Range.Text = "���� ��������(������ � ��������)��";
    //            table.Cell(1, 6).Range.Text = "��������� �������� �� (����� �������)";

    //            Object beforeRow1 = Type.Missing;
    //            table.Rows.Add(ref beforeRow1);


    //            int count = 1;

    //            // �������� ������� �������.
    //            foreach (��������������� item in list)
    //            {
    //                table.Cell(count+1, 1).Range.Text = count.ToString().Trim();
    //                table.Cell(count+1, 2).Range.Text = item.����������������������.Trim();
    //                table.Cell(count + 1, 3).Range.Text = item.�����������������.Trim();
    //                table.Cell(count + 1, 4).Range.Text = item.��������������.Trim();
    //                table.Cell(count + 1, 5).Range.Text = item.������������.Trim();
    //                table.Cell(count + 1, 6).Range.Text = item.�������������.Trim();

    //                Object beforeRow2 = Type.Missing;
    //                table.Rows.Add(ref beforeRow2);

    //                count++;
    //            }

    //            //������ ��������� ������
    //            table.Rows[count+1].Delete();

    //            // ������� ��� ����������� �����.
    //            string user = MyAplicationIdentity.GetUses();

    //            string[] arryFIO = user.Split(' ');
    //            string ����������� = arryFIO[1].Substring(0, 1);
    //            string ���������������� = arryFIO[2].Substring(0, 1);
    //            string fio = arryFIO[0] + " " + ����������� + "." + " " + ���������������� + ".";

    //            // ���� ��������� ������.
    //            object wdrepl3 = WdReplace.wdReplaceAll;
    //            //object searchtxt = "GreetingLine";
    //            object searchtxt3 = "USER";
    //            object newtxt3 = (object)fio;
    //            //object frwd = true;
    //            object frwd3 = false;
    //            doc.Content.Find.Execute(ref searchtxt3, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd3, ref missing, ref missing, ref newtxt3, ref wdrepl3, ref missing, ref missing,
    //            ref missing, ref missing);



    //            // ���������� �������� � ������� ����.
            //            app.Visible = true;
            #endregion

            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// ���������� ������ �������� �����.
        /// </summary>
        /// <param name="con"></param>
        /// <param name="ds"></param>
        /// <param name="id��������"></param>
        /// <returns></returns>
        private string InputNumberPoust(SqlConnection con,DataSet ds, int id��������, out string shortAbout)
        {
            string querSelect = "select CONVERT(nvarchar, �������) + N'/' + RTRIM(LTRIM(CONVERT(nvarchar, " +
                                                    "���������))) AS '�����',�����������������  from �������� " +
                                                    "where id_�������� in ( " +
                                                    "select id_���������������� from dbo.���������������������������������� " +
                                                    "where id_����������������� = " + Convert.ToInt32(id��������) + ")";

            SqlDataAdapter daM = new SqlDataAdapter(querSelect, con);
            daM.Fill(ds, "��������");

            // ������� ������ �� ������� ������� ������������ 
            System.Data.DataTable tabM = ds.Tables["��������"];

            // ������ �������� ��������.
            StringBuilder buld = new StringBuilder();

            // ������ �������� ������.
            StringBuilder buldAbout = new StringBuilder();

            foreach (DataRow r in tabM.Rows)
            {
                buld.Append(r["�����"].ToString().Trim() + ", \n");
                buldAbout.Append(r["�����������������"].ToString().Trim() + ", \n");
            }

            // ������ ��������� �������.
            int countChar = buld.Length;
            string numbers = string.Empty;

            if (countChar > 0)
            {
                numbers = buld.Remove(countChar - 3, 3).ToString();
            }

            // ������ ��������� ������� �� ��������.
            int countAbout = buldAbout.Length;
            string aboutS = string.Empty;

            // �������� ������ ������.
            shortAbout = "";

            if (countAbout > 0)
            {
                aboutS = buldAbout.Remove(countAbout - 3, 3).ToString();
                shortAbout = aboutS;
            }


            ds.Tables["��������"].Clear();

            return numbers.ToString().Trim();
        }
    }
}