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
        private string _заголовок = string.Empty;

        public ExcelPrint(string заголовок)
        {
            _заголовок = заголовок;
        }

        public void PrintСтатистикаВходящейКорреспонденции(List<СтатистикаВходИсполнителей> list)
        {
            ////Объект Excel
            //Microsoft.Office.Interop.Excel.Application ObjExcel;

            ////объект массив excel книг
            //Microsoft.Office.Interop.Excel.Workbooks ObjWorkBooks;

            ////Объект excel книга
            //Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;

            ////объект excel лист
            //Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

            //Книга.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);

            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            ObjWorkSheet.PageSetup.Zoom = false; ;
            ObjWorkSheet.PageSetup.FitToPagesWide = 1;
            ObjWorkSheet.PageSetup.FitToPagesTall = 800;
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;

            //ширина первой строки документа.
            int ширПервойСтроки = 80;
            int ширинаСтроки = 50;

            //Запишем шапку
            //Объеденим ячейки
            ObjWorkSheet.get_Range("A1", "H1").Merge(Type.Missing);
            ObjWorkSheet.get_Range("A1", "H1").Font.Size = 12;
            ObjWorkSheet.get_Range("A1", "H1").Font.Bold = true;
            //ObjWorkSheet.get_Range("A1", Type.Missing).Value2 = " Статистика по входящей корреспонденции на " + DateTime.Now.ToShortDateString() + " по КСЗН г. Саратова";
            ObjWorkSheet.get_Range("A1", Type.Missing).Value2 = _заголовок.Trim();
            ObjWorkSheet.get_Range("A1", "H1").HorizontalAlignment = Excel.Constants.xlCenter;

            //Запишем шапку
            //Объеденим ячейки
            ObjWorkSheet.get_Range("E1", "H1").Merge(Type.Missing);
            ObjWorkSheet.get_Range("E1", "H1").Font.Size = 12;
            ObjWorkSheet.get_Range("E1", "H1").Font.Bold = true;
            ObjWorkSheet.get_Range("E1", Type.Missing).Value2 = "Информация по бесплатному зубопротезированию на " + DateTime.Now.ToShortDateString() + " по КСЗН г. Саратова";

            // Формируем таблицу.

            //Объеденим ячейки
            ObjWorkSheet.get_Range("A2", "A3").Merge(Type.Missing);
            ObjWorkSheet.get_Range("A2", Type.Missing).Value2 = "№ п.п.";
            ObjWorkSheet.get_Range("A2", "A3").ColumnWidth = 70;
            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("A2", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("A2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            //ObjWorkSheet.get_Range("A2", "A3").RowHeight = ширинаСтроки;

            // Нарисуем границу.
            ExcelЯчейка excПП = new ExcelЯчейка();
            excПП.ГраницаЯчейки("A2", "A3", ObjWorkSheet);

            //Объеденим ячейки
            ObjWorkSheet.get_Range("B2", "B3").Merge(Type.Missing);

            // Зададим ширину колонки.
            ObjWorkSheet.get_Range("B2", "B3").ColumnWidth = 70;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширинаСтроки;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширПервойСтроки;
            ObjWorkSheet.get_Range("B2", Type.Missing).Value2 = "Наименование корреспондента";

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("B2", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("B2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // Нарисуем границу.
            ExcelЯчейка excРайон = new ExcelЯчейка();
            excРайон.ГраницаЯчейки("B2", "B3", ObjWorkSheet);

            // Ячейка С2-С3.
            //Объеденим ячейки
            ObjWorkSheet.get_Range("C2", "C3").Merge(Type.Missing);

            // Зададим ширину колонки.
            ObjWorkSheet.get_Range("C2", "C3").ColumnWidth = 25;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширинаСтроки;

            ObjWorkSheet.get_Range("C2", "C3").WrapText = true;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширПервойСтроки;
            ObjWorkSheet.get_Range("C2", Type.Missing).Value2 = "Количество исходящих документов";

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("C2", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("C2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ExcelЯчейка exc1 = new ExcelЯчейка();
            exc1.ГраницаЯчейки("C2", "C3", ObjWorkSheet);

            // Объединим ячейки 
            ObjWorkSheet.get_Range("D2", "G2").Merge(Type.Missing);

            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширПервойСтроки;
            ObjWorkSheet.get_Range("D2", "G2").Value2 = "В том числе - по способу получения документа";

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("D2", "G2").HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("D2", "G2").VerticalAlignment = Excel.Constants.xlCenter;

            ExcelЯчейка exc2 = new ExcelЯчейка();
            exc2.ГраницаЯчейки("D2", "G2", ObjWorkSheet);



            // Ячейка D3.
            // Зададим ширину колонки.
            ObjWorkSheet.get_Range("D3", "D3").ColumnWidth = 15;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширинаСтроки;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширПервойСтроки;
            ObjWorkSheet.get_Range("D3", Type.Missing).Value2 = "бумажный носитель, шт.";

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("D3", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("D3", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ObjWorkSheet.get_Range("D3", Type.Missing).WrapText = true;

            ExcelЯчейка exc21 = new ExcelЯчейка();
            exc21.ГраницаЯчейки("D3", "D3", ObjWorkSheet);

            // Ячейка Е3.
            ObjWorkSheet.get_Range("E3", "E3").ColumnWidth = 15;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширинаСтроки;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширПервойСтроки;
            ObjWorkSheet.get_Range("E3", Type.Missing).Value2 = "электронная почта, шт.";

            ObjWorkSheet.get_Range("E3", Type.Missing).WrapText = true;

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("E3", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("E3", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ExcelЯчейка exc3 = new ExcelЯчейка();
            exc3.ГраницаЯчейки("E3", "E3", ObjWorkSheet);

            // Ячейка F3.
            ObjWorkSheet.get_Range("F3", "F3").ColumnWidth = 15;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширинаСтроки;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширПервойСтроки;
            ObjWorkSheet.get_Range("F3", Type.Missing).Value2 = "VipNet шт.";

            ObjWorkSheet.get_Range("F3", Type.Missing).WrapText = true;

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("F3", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("F3", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ExcelЯчейка exc4 = new ExcelЯчейка();
            exc4.ГраницаЯчейки("F3", "F3", ObjWorkSheet);

            // Ячейка G3
            ObjWorkSheet.get_Range("G3", "G3").ColumnWidth = 15;
            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширинаСтроки;


            //ObjWorkSheet.get_Range("B2", "B3").RowHeight = ширПервойСтроки;
            ObjWorkSheet.get_Range("G3", Type.Missing).Value2 = "факс шт.";

            ObjWorkSheet.get_Range("G3", Type.Missing).WrapText = true;

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("G3", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("G3", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ExcelЯчейка exc5 = new ExcelЯчейка();
            exc5.ГраницаЯчейки("G3", "G3", ObjWorkSheet);

            //Ячейка H3.
            ObjWorkSheet.get_Range("H2", "H3").Merge(Type.Missing);
            ObjWorkSheet.get_Range("H2", Type.Missing).Value2 = "Исполнитель";
            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("H2", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("H2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ObjWorkSheet.get_Range("H2", "H3").ColumnWidth = 25;

            ObjWorkSheet.get_Range("H2", "H3").WrapText = true;

            //ObjWorkSheet.get_Range("A2", "A3").RowHeight = ширинаСтроки;

            // Нарисуем границу.
            ExcelЯчейка exc6 = new ExcelЯчейка();
            exc6.ГраницаЯчейки("H2", "H3", ObjWorkSheet);

            // Заполним данными таблицу.
            // Счётчик циклов. Начнём с 5 потому что содержание таблица начинается с 5 строки.
            int iCount = 4;

            foreach (СтатистикаВходИсполнителей item in list)
            {

                System.Text.RegularExpressions.Regex reg = new System.Text.RegularExpressions.Regex(@"\D");
                MatchCollection matches =reg.Matches(item.НомерПП);
                if (matches.Count > 0)
                {
                    // Пометим строки с итоговым значением.
                    CellFontBold(ObjWorkSheet, "A", iCount, item.НомерПП);
                    CellFontBold(ObjWorkSheet, "B", iCount, item.НаименованиеКорреспондента);
                    CellFontBold(ObjWorkSheet, "C", iCount, item.КоличесвтоВходДокументов);
                    CellFontBold(ObjWorkSheet, "D", iCount, item.БумажныйНоститель);
                    CellFontBold(ObjWorkSheet, "E", iCount, item.EMail);
                    CellFontBold(ObjWorkSheet, "F", iCount, item.VipNet);
                    CellFontBold(ObjWorkSheet, "G", iCount, item.Fax);
                    CellFontBold(ObjWorkSheet, "H", iCount, item.Исполнитель);
                }
                else
                {
                    Cell(ObjWorkSheet, "A", iCount, item.НомерПП);
                    Cell(ObjWorkSheet, "B", iCount, item.НаименованиеКорреспондента);
                    Cell(ObjWorkSheet, "C", iCount, item.КоличесвтоВходДокументов);
                    Cell(ObjWorkSheet, "D", iCount, item.БумажныйНоститель);
                    Cell(ObjWorkSheet, "E", iCount, item.EMail);
                    Cell(ObjWorkSheet, "F", iCount, item.VipNet);
                    Cell(ObjWorkSheet, "G", iCount, item.Fax);
                    Cell(ObjWorkSheet, "H", iCount, item.Исполнитель);
                }

                iCount++;
            }

            ObjExcel.Save(@"D:\111\Test\Book1.xml");

            System.Windows.Forms.MessageBox.Show("Файл сохранился");
            
            //// Отобразим документ.
            //ObjExcel.Visible = true;
            //ObjExcel.UserControl = true;

        }

        private void Cell(Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet, string exclБукв, int iCount, string value)
        {
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).Value2 = value;
            // выровним горизонтально.
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            // выровним вертикально.
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // Перенос текста.
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).WrapText = true;

            // Нарисуем границу.
            ExcelЯчейка excNum = new ExcelЯчейка();
            excNum.ГраницаЯчейки(exclБукв + iCount.ToString(), exclБукв + iCount.ToString(), ObjWorkSheet);

        }

        private void CellFontBold(Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet, string exclБукв, int iCount, string value)
        {
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).Value2 = value;
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).Font.Bold = 1;
            // выровним горизонтально.
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            // выровним вертикально.
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // Перенос текста.
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).WrapText = true;



            // Нарисуем границу.
            ExcelЯчейка excNum = new ExcelЯчейка();
            excNum.ГраницаЯчейки(exclБукв + iCount.ToString(), exclБукв + iCount.ToString(), ObjWorkSheet);

        }

        public void SaveFileCSV(List<СтатистикаВходИсполнителей> list)
        {

            // Строка для хранения заголовка отчета.
            СтатистикаВходИсполнителей itHead = new СтатистикаВходИсполнителей();

            itHead.НомерПП = this._заголовок;

            list.Insert(0, itHead);

            //Скопируем шаблон в папку Документы
            //FileInfo fn = new FileInfo(System.Windows.Forms.Application.StartupPath + @"\Шаблон\Договор.doc");
            //fn.CopyTo(System.Windows.Forms.Application.StartupPath + @"\Документы\" + fName + ".doc", true);
            string path = System.Windows.Forms.Application.StartupPath + @"\Документы\text.csv";
            //using (FileStream fs = File.Create(@"D:\111\Test\text.csv"))
            //using (TextWriter writer = new StreamWriter(fs))
            //{
            //    foreach(СтатистикаВходИсполнителей it in list)
            //    {
            //        writer.WriteLine(it.НомерПП + ";" + it.НаименованиеКорреспондента + ";" + it.КоличесвтоВходДокументов + ";" + it.БумажныйНоститель + ";" + it.VipNet + ";" + it.Fax + ";" + it.EMail + ";" + it.Исполнитель);//, Encoding.Unicode);//.GetEncoding(1251));
            //    }
            //}

            //===========

            string[] strArry = new string[list.Count];
            int iCount =0;
            foreach (СтатистикаВходИсполнителей it in list)
            {
                strArry[iCount] = it.НомерПП + ";" + it.НаименованиеКорреспондента + ";" + it.КоличесвтоВходДокументов + ";" + it.БумажныйНоститель + ";" + it.VipNet + ";" + it.Fax + ";" + it.EMail + ";" + it.Исполнитель;

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


           ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.Sheets[1];

           int i = 1;
           foreach (СтатистикаВходИсполнителей it in list)
           {
               if (excelapp.get_Range("A" + i.ToString().Trim(), "A" + i.ToString().Trim()).Text.ToString().ToLower().Trim().IndexOf("Итого".ToLower().Trim()) != -1)
               {
                   excelapp.get_Range("A" + i.ToString().Trim(), "H" + i.ToString().Trim()).Font.Bold = 1;
               }
               i++;
           }

            //Process.Start(@"D:\111\Test\text1.csv");
        }
       
    }
}
