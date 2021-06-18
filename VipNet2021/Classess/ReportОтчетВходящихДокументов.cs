using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.Sql;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Windows.Forms;
using System.Configuration;

namespace RegKor.Classess
{
    /// <summary>
    /// Класс формирования отчета входящих документов.
    /// </summary>
    public class ReportОтчетВходящихДокументов : IPrintReport
    {
        private string _заголовок;
        private DataGridViewRowCollection collection;

        public ReportОтчетВходящихДокументов(string заголовок)
        {
            _заголовок = заголовок;
        }

        /// <summary>
        /// Сеттер данных для отчета.
        /// </summary>
        public DataGridViewRowCollection SetDate
        {
            get
            {
                return collection;
            }
            set
            {
                collection = value;
            }
        }


        /// <summary>
        /// Формирует файл Excel и выводит его на экран.
        /// </summary>
        public void Execute()
        {

            string path = System.Windows.Forms.Application.StartupPath + @"\Документы\text.csv";

            // Здесь добавим еще две строки для шапки таблицы отчета и заголовка отчета.
            string[] strArry = new string[collection.Count + 3];

            // Первая строка содержит заголовок.
            strArry[0] = _заголовок + ";;;;;;;;";
            // Вторая строка содержит шапку таблицы.
            strArry[1] = "№ п.п.;Корреспондент;Дата исходящая;Номер исходящий;Краткое содержание;Дата входящая;Номер входящий;Срок исполнения;Результат исполнения;Исполнитель";

            // Счетчик строк.
            int iCount = 2;

            // Заполним данными массив строк.
            foreach (DataGridViewRow it in collection)
            {
                if (iCount <= collection.Count)
                    strArry[iCount] = (iCount - 1).ToString().Trim() + ";" 
                        + it.Cells["Корреспондент"].Value.ToString().Trim()
                        + ";" + Convert.ToDateTime(it.Cells["ДатаИсхода"].Value).ToShortDateString().Trim() + ";" 
                        + it.Cells["НомерИсход"].Value.ToString().Trim() + ";" 
                        + it.Cells["КраткоеСодержание"].Value.ToString().Trim()
                    + ";" + it.Cells["ДатаПоступ"].Value.ToString().Trim() + ";" 
                    + it.Cells["Номер входящий"].Value.ToString().Trim() +
                    ";" + it.Cells["СрокВыполнения"].Value.ToString().Trim() +
                    ";" + it.Cells["РезультатВыполнения"].Value.ToString().Trim() +
                    ";" + it.Cells["Исполнитель"].Value.ToString().Trim();

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

            excelapp.get_Range("A1", "A1").ColumnWidth = 10;
            excelapp.get_Range("B1", "B1").ColumnWidth = 90;
            excelapp.get_Range("C1", "C1").ColumnWidth = 16;
            excelapp.get_Range("D1", "D1").ColumnWidth = 18;
            excelapp.get_Range("E1", "E1").ColumnWidth = 60;
            excelapp.get_Range("F1", "F1").ColumnWidth = 16;
            excelapp.get_Range("G1", "G1").ColumnWidth = 20;
            excelapp.get_Range("H1", "H1").ColumnWidth = 19;
            excelapp.get_Range("I1", "I1").ColumnWidth = 50;
            excelapp.get_Range("J1", "J1").ColumnWidth = 16;

            //Книга.
            book = excelapp.Workbooks[1];

            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

            // Установим горизонтальное ориентирование страницы.
            ObjWorkSheet.PageSetup.Zoom = false;
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
            ObjWorkSheet.PageSetup.FitToPagesWide = 1;
            ObjWorkSheet.PageSetup.FitToPagesTall = 800;


            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

            // Выровним по центру и выделим жирным заглавную строку.
            excelapp.get_Range("A1", "F1").Font.Bold = 1;
            // Объеденим верхную ячейку.
            excelapp.get_Range("A1", "F1").Merge(Type.Missing);
            // Выравним по центру.
            excelapp.get_Range("A1", "F1").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            excelapp.get_Range("A2", "J2").Font.Bold = 1;
            excelapp.get_Range("A2", "J2").HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Нарисуем границы ячеек.
            int i = 1;

           // ExcelTablePrint

            if (ConfigurationSettings.AppSettings["ExcelTablePrint"].Trim() == "1".Trim())
            {
                foreach (string s in strArry)
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
                        ГраницаЯчейки("I" + i.ToString().Trim(), "I" + i.ToString().Trim(), ObjWorkSheet);
                        ГраницаЯчейки("J" + i.ToString().Trim(), "J" + i.ToString().Trim(), ObjWorkSheet);
                    }
                    i++;
                }
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

            // Выравнивание.
            //ObjWorkSheet.get_Range(cell1, cell2).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
        }



    }


}
