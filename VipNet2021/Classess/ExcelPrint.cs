using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;



namespace RegKor.Classess
{
    public class ExcelPrint
    {
        private string _��������� = string.Empty;

        public ExcelPrint(string ���������)
        {
            _��������� = ���������;
        }

        public void Print���������������������������������(List<��������������������������> list)
        {
            ////������ Excel
            //Microsoft.Office.Interop.Excel.Application ObjExcel;

            ////������ ������ excel ����
            //Microsoft.Office.Interop.Excel.Workbooks ObjWorkBooks;

            ////������ excel �����
            //Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;

            ////������ excel ����
            //Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

            //�����.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);

            //�������.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            ObjWorkSheet.PageSetup.Zoom = false; ;
            ObjWorkSheet.PageSetup.FitToPagesWide = 1;
            ObjWorkSheet.PageSetup.FitToPagesTall = 800;
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;

            //������ ������ ������ ���������.
            int ��������������� = 80;
            int ������������ = 50;

            //������� �����
            //��������� ������
            ObjWorkSheet.get_Range("A1", "H1").Merge(Type.Missing);
            ObjWorkSheet.get_Range("A1", "H1").Font.Size = 12;
            ObjWorkSheet.get_Range("A1", "H1").Font.Bold = true;
            //ObjWorkSheet.get_Range("A1", Type.Missing).Value2 = " ���������� �� �������� ��������������� �� " + DateTime.Now.ToShortDateString() + " �� ���� �. ��������";
            ObjWorkSheet.get_Range("A1", Type.Missing).Value2 = _���������.Trim();
            ObjWorkSheet.get_Range("A1", "H1").HorizontalAlignment = Excel.Constants.xlCenter;

            //������� �����
            //��������� ������
            ObjWorkSheet.get_Range("E1", "H1").Merge(Type.Missing);
            ObjWorkSheet.get_Range("E1", "H1").Font.Size = 12;
            ObjWorkSheet.get_Range("E1", "H1").Font.Bold = true;
            ObjWorkSheet.get_Range("E1", Type.Missing).Value2 = "���������� �� ����������� ������������������ �� " + DateTime.Now.ToShortDateString() + " �� ���� �. ��������";

            // ��������� �������.

            //��������� ������
            ObjWorkSheet.get_Range("A2", "A3").Merge(Type.Missing);
            ObjWorkSheet.get_Range("A2", Type.Missing).Value2 = "� �.�.";
            ObjWorkSheet.get_Range("A2", "A3").ColumnWidth = 70;
            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("A2", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("A2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            //ObjWorkSheet.get_Range("A2", "A3").RowHeight = ������������;

            // �������� �������.
            Excel������ exc�� = new Excel������();
            exc��.�������������("A2", "A3", ObjWorkSheet);

            //��������� ������
            ObjWorkSheet.get_Range("B2", "B3").Merge(Type.Missing);

            // ������� ������ �������.
            ObjWorkSheet.get_Range("B2", "B3").ColumnWidth = 70;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ������������;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ���������������;
            ObjWorkSheet.get_Range("B2", Type.Missing).Value2 = "������������ ��������������";

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("B2", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("B2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // �������� �������.
            Excel������ exc����� = new Excel������();
            exc�����.�������������("B2", "B3", ObjWorkSheet);

            // ������ �2-�3.
            //��������� ������
            ObjWorkSheet.get_Range("C2", "C3").Merge(Type.Missing);

            // ������� ������ �������.
            ObjWorkSheet.get_Range("C2", "C3").ColumnWidth = 25;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ������������;

            ObjWorkSheet.get_Range("C2", "C3").WrapText = true;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ���������������;
            ObjWorkSheet.get_Range("C2", Type.Missing).Value2 = "���������� ��������� ����������";

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("C2", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("C2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            Excel������ exc1 = new Excel������();
            exc1.�������������("C2", "C3", ObjWorkSheet);

            // ��������� ������ 
            ObjWorkSheet.get_Range("D2", "G2").Merge(Type.Missing);

            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ���������������;
            ObjWorkSheet.get_Range("D2", "G2").Value2 = "� ��� ����� - �� ������� ��������� ���������";

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("D2", "G2").HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("D2", "G2").VerticalAlignment = Excel.Constants.xlCenter;

            Excel������ exc2 = new Excel������();
            exc2.�������������("D2", "G2", ObjWorkSheet);



            // ������ D3.
            // ������� ������ �������.
            ObjWorkSheet.get_Range("D3", "D3").ColumnWidth = 15;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ������������;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ���������������;
            ObjWorkSheet.get_Range("D3", Type.Missing).Value2 = "�������� ��������, ��.";

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("D3", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("D3", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ObjWorkSheet.get_Range("D3", Type.Missing).WrapText = true;

            Excel������ exc21 = new Excel������();
            exc21.�������������("D3", "D3", ObjWorkSheet);

            // ������ �3.
            ObjWorkSheet.get_Range("E3", "E3").ColumnWidth = 15;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ������������;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ���������������;
            ObjWorkSheet.get_Range("E3", Type.Missing).Value2 = "����������� �����, ��.";

            ObjWorkSheet.get_Range("E3", Type.Missing).WrapText = true;

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("E3", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("E3", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            Excel������ exc3 = new Excel������();
            exc3.�������������("E3", "E3", ObjWorkSheet);

            // ������ F3.
            ObjWorkSheet.get_Range("F3", "F3").ColumnWidth = 15;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ������������;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ���������������;
            ObjWorkSheet.get_Range("F3", Type.Missing).Value2 = "VipNet ��.";

            ObjWorkSheet.get_Range("F3", Type.Missing).WrapText = true;

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("F3", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("F3", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            Excel������ exc4 = new Excel������();
            exc4.�������������("F3", "F3", ObjWorkSheet);

            // ������ G3
            ObjWorkSheet.get_Range("G3", "G3").ColumnWidth = 15;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ������������;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ���������������;
            ObjWorkSheet.get_Range("G3", Type.Missing).Value2 = "���� ��.";

            ObjWorkSheet.get_Range("G3", Type.Missing).WrapText = true;

            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("G3", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("G3", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            Excel������ exc5 = new Excel������();
            exc5.�������������("G3", "G3", ObjWorkSheet);

            //������ H3.
            ObjWorkSheet.get_Range("H2", "H3").Merge(Type.Missing);
            ObjWorkSheet.get_Range("H2", Type.Missing).Value2 = "�����������";
            // �������� ����� �� �����������.
            ObjWorkSheet.get_Range("H2", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("H2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ObjWorkSheet.get_Range("H2", "H3").ColumnWidth = 25;

            ObjWorkSheet.get_Range("H2", "H3").WrapText = true;

            //ObjWorkSheet.get_Range("A2", "A3").RowHeight = ������������;

            // �������� �������.
            Excel������ exc6 = new Excel������();
            exc6.�������������("H2", "H3", ObjWorkSheet);

            // �������� ������� �������.
            // ������� ������. ����� � 5 ������ ��� ���������� ������� ���������� � 5 ������.
            int iCount = 4;

            foreach (�������������������������� item in list)
            {

                System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"\D");
                MatchCollection matches =reg.Matches(item.�������);
                if (matches.Count > 0)
                {
                    // ������� ������ � �������� ���������.
                    CellFontBold(ObjWorkSheet, "A", iCount, item.�������);
                    CellFontBold(ObjWorkSheet, "B", iCount, item.��������������������������);
                    CellFontBold(ObjWorkSheet, "C", iCount, item.������������������������);
                    CellFontBold(ObjWorkSheet, "D", iCount, item.�����������������);
                    CellFontBold(ObjWorkSheet, "E", iCount, item.EMail);
                    CellFontBold(ObjWorkSheet, "F", iCount, item.VipNet);
                    CellFontBold(ObjWorkSheet, "G", iCount, item.Fax);
                    CellFontBold(ObjWorkSheet, "H", iCount, item.�����������);
                }
                else
                {
                    Cell(ObjWorkSheet, "A", iCount, item.�������);
                    Cell(ObjWorkSheet, "B", iCount, item.��������������������������);
                    Cell(ObjWorkSheet, "C", iCount, item.������������������������);
                    Cell(ObjWorkSheet, "D", iCount, item.�����������������);
                    Cell(ObjWorkSheet, "E", iCount, item.EMail);
                    Cell(ObjWorkSheet, "F", iCount, item.VipNet);
                    Cell(ObjWorkSheet, "G", iCount, item.Fax);
                    Cell(ObjWorkSheet, "H", iCount, item.�����������);
                }

                iCount++;
            }

            ObjExcel.Save(@"D:\111\Test\Book1.xml");

            System.Windows.Forms.MessageBox.Show("���� ����������");
            
            //// ��������� ��������.
            //ObjExcel.Visible = true;
            //ObjExcel.UserControl = true;

        }

        private void Cell(Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet, string excl����, int iCount, string value)
        {
            ObjWorkSheet.get_Range(excl���� + iCount.ToString(), Type.Missing).Value2 = value;
            // �������� �������������.
            ObjWorkSheet.get_Range(excl���� + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            // �������� �����������.
            ObjWorkSheet.get_Range(excl���� + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // ������� ������.
            ObjWorkSheet.get_Range(excl���� + iCount.ToString(), Type.Missing).WrapText = true;

            // �������� �������.
            Excel������ excNum = new Excel������();
            excNum.�������������(excl���� + iCount.ToString(), excl���� + iCount.ToString(), ObjWorkSheet);

        }

        private void CellFontBold(Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet, string excl����, int iCount, string value)
        {
            ObjWorkSheet.get_Range(excl���� + iCount.ToString(), Type.Missing).Value2 = value;
            ObjWorkSheet.get_Range(excl���� + iCount.ToString(), Type.Missing).Font.Bold = 1;
            // �������� �������������.
            ObjWorkSheet.get_Range(excl���� + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            // �������� �����������.
            ObjWorkSheet.get_Range(excl���� + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // ������� ������.
            ObjWorkSheet.get_Range(excl���� + iCount.ToString(), Type.Missing).WrapText = true;



            // �������� �������.
            Excel������ excNum = new Excel������();
            excNum.�������������(excl���� + iCount.ToString(), excl���� + iCount.ToString(), ObjWorkSheet);

        }

        public void SaveFileCSV(List<��������������������������> list)
        {

            // ������ ��� �������� ��������� ������.
            �������������������������� itHead = new ��������������������������();

            itHead.������� = this._���������;

            list.Insert(0, itHead);

            //��������� ������ � ����� ���������
            //FileInfo fn = new FileInfo(System.Windows.Forms.Application.StartupPath + @"\������\�������.doc");
            //fn.CopyTo(System.Windows.Forms.Application.StartupPath + @"\���������\" + fName + ".doc", true);
            string path = System.Windows.Forms.Application.StartupPath + @"\���������\text.csv";
            //using (FileStream fs = File.Create(@"D:\111\Test\text.csv"))
            //using (TextWriter writer = new StreamWriter(fs))
            //{
            //    foreach(�������������������������� it in list)
            //    {
            //        writer.WriteLine(it.������� + ";" + it.�������������������������� + ";" + it.������������������������ + ";" + it.����������������� + ";" + it.VipNet + ";" + it.Fax + ";" + it.EMail + ";" + it.�����������);//, Encoding.Unicode);//.GetEncoding(1251));
            //    }
            //}

            //===========

            string[] strArry = new string[list.Count];
            int iCount =0;
            foreach (�������������������������� it in list)
            {
                strArry[iCount] = it.������� + ";" + it.�������������������������� + ";" + it.������������������������ + ";" + it.����������������� + ";" + it.VipNet + ";" + it.Fax + ";" + it.EMail + ";" + it.�����������;

                iCount++;
            }

            File.WriteAllLines(path, strArry, Encoding.GetEncoding(1251));
 
            //�������� ��������� Excel.
            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook book;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            excelapp.Visible = true;

            excelapp.Workbooks._OpenText(
             path,
             Excel.XlPlatform.xlWindows,
             1,            //� ������ ������
             Excel.XlTextParsingType.xlDelimited, //����� � �������������
             Excel.XlTextQualifier.xlTextQualifierDoubleQuote, //������� ��������� ������� ������
             true,          //����������� ���������
             false,          //����������� :Tab
             false,         //Semicolon
             false,         //Comma
             false,         //Space
             false,         //Other
             false,  //OtherChar
             new object[] {new object[]{1,Excel.XlColumnDataType.xlSkipColumn},
                                new object[]{2,Excel.XlColumnDataType.xlGeneralFormat},
                                new object[]{2,Excel.XlColumnDataType.xlMDYFormat},
                                new object[]{3,Excel.XlColumnDataType.xlMYDFormat},
                                new object[]{4,Excel.XlColumnDataType.xlTextFormat},
                                new object[]{5,Excel.XlColumnDataType.xlTextFormat}},
             Type.Missing,  //���������� ������
             ".",           //����������� ���������� ��������
            ",");           //����������� �����

           excelapp.get_Range("A1", "A1").ColumnWidth = 30;
           excelapp.get_Range("B1", "B1").ColumnWidth = 100;
           excelapp.get_Range("C1", "C1").ColumnWidth = 10;
           excelapp.get_Range("D1", "D1").ColumnWidth = 20;
           excelapp.get_Range("E1", "E1").ColumnWidth = 10;
           excelapp.get_Range("F1", "F1").ColumnWidth = 10;
           excelapp.get_Range("G1", "G1").ColumnWidth = 10;
           excelapp.get_Range("H1", "H1").ColumnWidth = 20;


           //�����.
           book = excelapp.Workbooks[1];


           ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

           int i = 1;
           foreach (�������������������������� it in list)
           {
               if (excelapp.get_Range("A" + i.ToString().Trim(), "A" + i.ToString().Trim()).Text.ToString().ToLower().Trim().IndexOf("�����".ToLower().Trim()) != -1)
               {
                   excelapp.get_Range("A" + i.ToString().Trim(), "H" + i.ToString().Trim()).Font.Bold = 1;
               }
               i++;
           }

            //Process.Start(@"D:\111\Test\text1.csv");
        }
       
    }
}
