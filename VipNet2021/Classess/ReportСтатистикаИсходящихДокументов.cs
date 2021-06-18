using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace RegKor.Classess
{
    /// <summary>
    ///  ласс отчет (приЄмник команды)
    /// </summary>
    public class Report—татистика»сход€щихƒокументов : IReport
    {
        private string _заголовок;

        public Report—татистика»сход€щихƒокументов(string заголовок)
        {
            _заголовок = заголовок;
        }
        public void PrintReport(List<—татистика¬ход»сполнителей> list)
        {
            // —трока дл€ хранени€ заголовка отчета.
            —татистика¬ход»сполнителей itHead = new —татистика¬ход»сполнителей();

            itHead.Ќомерѕѕ = this._заголовок;

            list.Insert(0, itHead);

            string path = System.Windows.Forms.Application.StartupPath + @"\ƒокументы\text.csv";

            string[] strArry = new string[list.Count];
            int iCount = 0;
            foreach (—татистика¬ход»сполнителей it in list)
            {
                strArry[iCount] = it.Ќомерѕѕ + ";" + it.Ќаименование орреспондента + ";" + it. оличесвто¬ходƒокументов + ";" + it.ЅумажныйЌоститель + ";" + it.VipNet + ";" + it.Fax + ";" + it.EMail + ";" + it.»сполнитель;

                iCount++;
            }

            File.WriteAllLines(path, strArry, Encoding.GetEncoding(1251));

            //—оздадим экземпл€р Excel.
            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook book;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            excelapp.Visible = true;

            excelapp.Workbooks._OpenText(
             path,
             Excel.XlPlatform.xlWindows,
             1,            //— первой строки
             Excel.XlTextParsingType.xlDelimited, //“екст с разделител€ми
             Excel.XlTextQualifier.xlTextQualifierDoubleQuote, //ѕризнак окончани€ разбора строки
             true,          //–азделители одинарные
             false,          //–азделители :Tab
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
             Type.Missing,  //–азмещение текста
             ".",           //–азделитель дес€тичных разр€дов
            ",");           //–азделитель тыс€ч

            excelapp.get_Range("A1", "A1").ColumnWidth = 30;
            excelapp.get_Range("B1", "B1").ColumnWidth = 100;
            excelapp.get_Range("C1", "C1").ColumnWidth = 10;
            excelapp.get_Range("D1", "D1").ColumnWidth = 20;
            excelapp.get_Range("E1", "E1").ColumnWidth = 10;
            excelapp.get_Range("F1", "F1").ColumnWidth = 10;
            excelapp.get_Range("G1", "G1").ColumnWidth = 10;
            excelapp.get_Range("H1", "H1").ColumnWidth = 20;


            // нига.
            book = excelapp.Workbooks[1];

            //“аблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

            // ”становим размер и расположение листа.
            ObjWorkSheet.PageSetup.Zoom = false;
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
            ObjWorkSheet.PageSetup.FitToPagesWide = 1;
            ObjWorkSheet.PageSetup.FitToPagesTall = 800;


            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

            // ¬ыровним по центру и выделим жирным заглавную строку.
            excelapp.get_Range("A1", "H1").Font.Bold = 1;
            excelapp.get_Range("A1", "H1").Merge(Type.Missing);
            excelapp.get_Range("A1", "H1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            

            int i = 1;
            foreach (—татистика¬ход»сполнителей it in list)
            {
                if (i > 1)
                {
                    √раницаячейки("A" + i.ToString().Trim(), "A" + i.ToString().Trim(), ObjWorkSheet);
                    √раницаячейки("B" + i.ToString().Trim(), "B" + i.ToString().Trim(), ObjWorkSheet);
                    √раницаячейки("C" + i.ToString().Trim(), "C" + i.ToString().Trim(), ObjWorkSheet);
                    √раницаячейки("D" + i.ToString().Trim(), "D" + i.ToString().Trim(), ObjWorkSheet);
                    √раницаячейки("E" + i.ToString().Trim(), "E" + i.ToString().Trim(), ObjWorkSheet);
                    √раницаячейки("F" + i.ToString().Trim(), "F" + i.ToString().Trim(), ObjWorkSheet);
                    √раницаячейки("G" + i.ToString().Trim(), "G" + i.ToString().Trim(), ObjWorkSheet);
                    √раницаячейки("H" + i.ToString().Trim(), "H" + i.ToString().Trim(), ObjWorkSheet);
                }

                if (excelapp.get_Range("A" + i.ToString().Trim(), "A" + i.ToString().Trim()).Text.ToString().ToLower().Trim().IndexOf("»того".ToLower().Trim()) != -1)
                {
                    excelapp.get_Range("A" + i.ToString().Trim(), "H" + i.ToString().Trim()).Font.Bold = 1;
                }
                i++;
            }
        }

        public void √раницаячейки(string cell1, string cell2, Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet)
        {
            //var cells = WorkSheet.get_Range("B2", "F5")
            //var cells = ObjWorkSheet.get_Range(cell1, cell2);

            // верхн€€ внешн€€.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;

            // права€ внешн€€.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            // Ћева€ внешна€.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;

            // Ќижна€ верхна€.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

        }
    }

    
}
