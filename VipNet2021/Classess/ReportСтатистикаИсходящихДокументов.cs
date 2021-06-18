using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace RegKor.Classess
{
    /// <summary>
    /// ����� ����� (������� �������)
    /// </summary>
    public class Report����������������������������� : IReport
    {
        private string _���������;

        public Report�����������������������������(string ���������)
        {
            _��������� = ���������;
        }
        public void PrintReport(List<��������������������������> list)
        {
            // ������ ��� �������� ��������� ������.
            �������������������������� itHead = new ��������������������������();

            itHead.������� = this._���������;

            list.Insert(0, itHead);

            string path = System.Windows.Forms.Application.StartupPath + @"\���������\text.csv";

            string[] strArry = new string[list.Count];
            int iCount = 0;
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

            //�������.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

            // ��������� ������ � ������������ �����.
            ObjWorkSheet.PageSetup.Zoom = false;
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
            ObjWorkSheet.PageSetup.FitToPagesWide = 1;
            ObjWorkSheet.PageSetup.FitToPagesTall = 800;


            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

            // �������� �� ������ � ������� ������ ��������� ������.
            excelapp.get_Range("A1", "H1").Font.Bold = 1;
            excelapp.get_Range("A1", "H1").Merge(Type.Missing);
            excelapp.get_Range("A1", "H1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            

            int i = 1;
            foreach (�������������������������� it in list)
            {
                if (i > 1)
                {
                    �������������("A" + i.ToString().Trim(), "A" + i.ToString().Trim(), ObjWorkSheet);
                    �������������("B" + i.ToString().Trim(), "B" + i.ToString().Trim(), ObjWorkSheet);
                    �������������("C" + i.ToString().Trim(), "C" + i.ToString().Trim(), ObjWorkSheet);
                    �������������("D" + i.ToString().Trim(), "D" + i.ToString().Trim(), ObjWorkSheet);
                    �������������("E" + i.ToString().Trim(), "E" + i.ToString().Trim(), ObjWorkSheet);
                    �������������("F" + i.ToString().Trim(), "F" + i.ToString().Trim(), ObjWorkSheet);
                    �������������("G" + i.ToString().Trim(), "G" + i.ToString().Trim(), ObjWorkSheet);
                    �������������("H" + i.ToString().Trim(), "H" + i.ToString().Trim(), ObjWorkSheet);
                }

                if (excelapp.get_Range("A" + i.ToString().Trim(), "A" + i.ToString().Trim()).Text.ToString().ToLower().Trim().IndexOf("�����".ToLower().Trim()) != -1)
                {
                    excelapp.get_Range("A" + i.ToString().Trim(), "H" + i.ToString().Trim()).Font.Bold = 1;
                }
                i++;
            }
        }

        public void �������������(string cell1, string cell2, Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet)
        {
            //var cells = WorkSheet.get_Range("B2", "F5")
            //var cells = ObjWorkSheet.get_Range(cell1, cell2);

            // ������� �������.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;

            // ������ �������.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            // ����� �������.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;

            // ������ �������.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

        }
    }

    
}
