using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Core;
using CarlosAg.ExcelXmlWriter;

namespace RegKor.Classess
{
    public class ExcelStatistic
    {
        Workbook book;
        WorksheetStyle style;
        Worksheet sheet;

        private int year;

        /// <summary>
        /// ��������� ��� ��� ������.
        /// </summary>
        public int Year
        {
            get
            {
                return year;
            }
            set
            {
                year = value;
            }
        }

        Dictionary<int, Dictionary<string, List<DocExcelCell>>> list; 

        public ExcelStatistic(Dictionary<int, Dictionary<string, List<DocExcelCell>>> listExcel)
        {
            list = listExcel;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="countColumns">���������� �������� ����������� ����������</param>
        public void CreateFile(int countColumns)
        {
            string filename = @"D:\test.xls";
            book = new Workbook();

            //book.ExcelWorkbook.ActiveSheetIndex = 1;

            style = book.Styles.Add("HeaderStyle");
            style.Font.FontName = "Tahoma";
            style.Font.Size = 12;
            //style.Font.Bold = true;
            style.Alignment.Horizontal = StyleHorizontalAlignment.Center;
            
            //style.Font.Color = "White";
            style.Interior.Color = "black";

            style.Borders.Add(StylePosition.Bottom, LineStyleOption.Continuous, 1);
            style.Borders.Add(StylePosition.Left, LineStyleOption.Continuous, 1);
            style.Borders.Add(StylePosition.Right, LineStyleOption.Continuous, 1);
            style.Borders.Add(StylePosition.Top, LineStyleOption.Continuous, 1);

             // �������� Default Style � ������������� ��� ����
            style = book.Styles.Add("Default");
            style.Font.FontName = "Tahoma";
            style.Font.Size = 10;

            // �������� ���� � ��������� ������
            sheet = book.Worksheets.Add("Some Data");

            int countColumn = list[40].Count;

             // ������� ����� ��������� ������.
            int iCountValid = 1;

            // ������� ��������.
            int count = 1;

            // ��������� ������ ��������.
            for (int iWdth = 0; iWdth <= countColumn - 1; iWdth++)
            {
                if(count <= 5)
                {
                    if(count <= 2)
                    {
                        sheet.Table.Columns.Add(new WorksheetColumn(150));
                    }

                    if(count > 2 && count <= 5)
                    {
                        sheet.Table.Columns.Add(new WorksheetColumn(70));
                    }
                }
                else if (count > 5)
                {
                    if (iCountValid == 1)
                    {
                        sheet.Table.Columns.Add(new WorksheetColumn(150));
                    }
                    else
                    {
                        sheet.Table.Columns.Add(new WorksheetColumn(70));
                    }

                    iCountValid++;

                    if (iCountValid == 5)
                    {
                        iCountValid = 1;
                    }
                }

                count++;
            }

            // ������ ������� Excel.
            WorksheetRow row = sheet.Table.Rows.Add();
            
            // ���������� �������.
            Dictionary<int, Dictionary<string, List<DocExcelCell>>> dict = list;

            // ���������� ����� � ������.
            int countColumntPrint = list[40].Values.Count;

            /*
               * ������� �������� � ������� �����.
               * ��� ��� � ��� �������� ����� ���������� �������� ������� ���������� � ������ (���������� countColumntPrint).
               * ��� �� � ��� �������� ������� ������ �������� �������� ����������� ����������, ������� ������� ����� ��� �������� ��������.
           */

            // ������ ��������� ��������.
            int indxPaper = countColumntPrint - 4;

            // ������ ��������� ������������ �� e-mail.
            int indxEmil = countColumntPrint - 2;

            // ������ ��������� ����������� �� VipNet.
            int indxVipNet = countColumntPrint - 1;

            // ������ ��������� ����������� �� Fax.
            int indxFax = countColumntPrint;


            int i = 1;

            foreach (int month in list.Keys)
            {
                // ������� ������.
                row = sheet.Table.Rows.Add();

                int iStartIndex = 2;

                Dictionary<string, List<DocExcelCell>> listValues = list[month];
               
                if (month < 20)
                {
                    string ����� = Montch.GetMonth(month);
                    int num = CellValue(row, �����, "HeaderStyle", 1, 1);
                }

                // ������� ������ �������� ��� �����.
                int index = 0;

                foreach (List<DocExcelCell> cells in listValues.Values)
                {
                    
                    foreach (DocExcelCell cell in cells)
                    {
                        int iTest2 = iStartIndex;

                        // �������� ���� ��������������.
                        if (cell.FlagEdit == false && month < 20)
                        {
                            cell.ValueCell = string.Empty;
                        }

                        if (cell.FlagEdit == true && month < 20)
                        {
                            string s =  cell.ValueCell;
                        }

                        // ������� �������� � ������.
                        int a = CellValue(row, cell.ValueCell, "HeaderStyle", iStartIndex, cell.CountColumn);

                        if (month < 20)
                        {
                            if (index == indxPaper && month < 13)
                            {
                                List<ItemStatisticDoc> itemNameDoc = �����������������������������.GetStatisticMontch(this.Year, month, "�������� ��������");

                                //CellValue(row, itemNameDoc[0].Count.ToString().Trim(), "HeaderStyle", 72, cell.CountColumn);

                                //int a1 = CellValue(row, itemNameDoc[0].Count.ToString().Trim(), "HeaderStyle", indxPaper, cell.CountColumn);
                            }
                           

                        }

                        iStartIndex = a;

                        string sTest = "";
                    }

                        index++;
                }
            }

            book.Save(filename);

            System.Diagnostics.Process.Start(filename);

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="row">������� ������</param>
        /// <param name="value">�������� ������� ������ ������������ � ������</param>
        /// <param name="nameStyle">�������� ���� ������</param>
        /// <param name="index">������ ������� (���������� � 1)</param>
        /// <param name="countColumn">���������� ������� � �������</param>
        /// <returns>����� ���������� ������� �������</returns>
        private int CellValue(WorksheetRow row, string value, string nameStyle, int index, int countColumn)
        {
            int count = 1;

            int columnCount = countColumn - 1;

            // ������� �������� � ������.
            WorksheetCell cell = row.Cells.Add(value);
            if (countColumn >= 2)
            {
                // ���������� ��������
                cell.MergeAcross = columnCount; // 1;
            }
            cell.StyleID = nameStyle;
            cell.Index = index;

            return index + countColumn;
        }


    }
}
