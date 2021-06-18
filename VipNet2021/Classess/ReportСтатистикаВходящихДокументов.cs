using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;
using System.Configuration;

namespace RegKor.Classess
{
    class ReportСтатистикаВходящихДокументов:IPrintReport
    {
        private string _заголовок;
        private List<StatisticDocInput> list;

        public ReportСтатистикаВходящихДокументов(string заголовок)
        {
            _заголовок = заголовок;
        }

        /// <summary>
        /// Сеттер данных для отчета.
        /// </summary>
        public List<StatisticDocInput> SetDate
        {
            get
            {
                return list;
            }
            set
            {
                list = value;
            }
        }

        public void Execute()
        {
            // Строка для хранения заголовка отчета.
            StatisticDocInput itHead = new StatisticDocInput();

            itHead.Num = this._заголовок;

            list.Insert(0, itHead);

            //========
            StatisticDocInput itHead2 = new StatisticDocInput();

            itHead2.Num = "";

            list.Insert(1, itHead2);

            string path = System.Windows.Forms.Application.StartupPath + @"\Документы\text.csv";

            string[] strArry = new string[list.Count];
            int iCount = 0;

            // Из за того что коллекцию нужно перевести в массив строк не используем стандарный метод ToArray;
            foreach (StatisticDocInput it in list)
            {
                if (iCount == 0)
                {
                    strArry[iCount] = it.Num + ";;;;;;;";
                }
                else if(iCount == 1)
                {
                    strArry[iCount] = "№ п.п.;Наименование корреспондента;Количество исходящих документов;Бумажный носитель;VipNet;факс;e-mail;Исполнитель";
                }
                else
                {
                    strArry[iCount] = it.Num + ";" + it.НаименованиеКорреспондента + ";" + it.КолвоВходКорреспонденции + ";" + it.БумажныйНоситель + ";" + it.VipNet + ";" + it.Fax + ";" + it.Email + ";" + it.Исполнитель;
                }

                iCount++;
            }

            File.WriteAllLines(path, strArry, Encoding.GetEncoding(1251));

            //Создадим экземпляр Excel.
            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook book;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            excelapp.Visible = true;

            excelapp.Workbooks._OpenText(
             path,
             Excel.XlPlatform.xlWindows,
             1,            //С первой строки
             Excel.XlTextParsingType.xlDelimited, //Текст с разделителями
             Excel.XlTextQualifier.xlTextQualifierDoubleQuote, //Признак окончания разбора строки
             true,          //Разделители одинарные
             false,          //Разделители :Tab
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
             Type.Missing,  //Размещение текста
             ".",           //Разделитель десятичных разрядов
            ",");           //Разделитель тысяч

            excelapp.get_Range("A1", "A1").ColumnWidth = 30;
            excelapp.get_Range("B1", "B1").ColumnWidth = 100;
            excelapp.get_Range("C1", "C1").ColumnWidth = 10;
            excelapp.get_Range("D1", "D1").ColumnWidth = 20;
            excelapp.get_Range("E1", "E1").ColumnWidth = 10;
            excelapp.get_Range("F1", "F1").ColumnWidth = 10;
            excelapp.get_Range("G1", "G1").ColumnWidth = 10;
            excelapp.get_Range("H1", "H1").ColumnWidth = 20;


            //Книга.
            book = excelapp.Workbooks[1];

            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

            ObjWorkSheet.PageSetup.Zoom = false;

            // Зададим горизонтальное расположение листа.
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
            ObjWorkSheet.PageSetup.FitToPagesWide = 1;
            ObjWorkSheet.PageSetup.FitToPagesTall = 800;


            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

            // Выровним по центру и выделим жирным заглавную строку.
            excelapp.get_Range("A1", "H1").Font.Bold = 1;
            excelapp.get_Range("A1", "H1").Merge(Type.Missing);
            excelapp.get_Range("A1", "H1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            //excelapp.get_Range("A2", "H2").Font.Bold = 1;
            //excelapp.get_Range("A2", "H2").Merge(Type.Missing);
            //excelapp.get_Range("A2", "H2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;



            int i = 1;
            foreach (StatisticDocInput it in list)
            {
                if (ConfigurationSettings.AppSettings["ExcelTablePrint"].Trim() == "1".Trim())
                {
                    if (i > 1)
                    {
                        ГраницаЯчейки("A" + i.ToString().Trim(), "A" + i.ToString().Trim(), ObjWorkSheet);
                        ГраницаЯчейки("B" + i.ToString().Trim(), "B" + i.ToString().Trim(), ObjWorkSheet);
                        ГраницаЯчейки("C" + i.ToString().Trim(), "C" + i.ToString().Trim(), ObjWorkSheet);
                        ГраницаЯчейки("D" + i.ToString().Trim(), "D" + i.ToString().Trim(), ObjWorkSheet);
                        ГраницаЯчейки("E" + i.ToString().Trim(), "E" + i.ToString().Trim(), ObjWorkSheet);
                        ГраницаЯчейки("F" + i.ToString().Trim(), "F" + i.ToString().Trim(), ObjWorkSheet);
                        ГраницаЯчейки("G" + i.ToString().Trim(), "G" + i.ToString().Trim(), ObjWorkSheet);
                        ГраницаЯчейки("H" + i.ToString().Trim(), "H" + i.ToString().Trim(), ObjWorkSheet);
                    }
                }

                if (excelapp.get_Range("A" + i.ToString().Trim(), "A" + i.ToString().Trim()).Text.ToString().ToLower().Trim().IndexOf("Итого".ToLower().Trim()) != -1)
                {
                    excelapp.get_Range("A" + i.ToString().Trim(), "H" + i.ToString().Trim()).Font.Bold = 1;
                }
                i++;
            }
        }

        public void ГраницаЯчейки(string cell1, string cell2, Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet)
        {
            //var cells = WorkSheet.get_Range("B2", "F5")
            //var cells = ObjWorkSheet.get_Range(cell1, cell2);

            // верхняя внешняя.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;

            // правая внешняя.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;

            // Левая внешная.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;

            // Нижная верхная.
            ObjWorkSheet.get_Range(cell1, cell2).Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

        }

    }
}
