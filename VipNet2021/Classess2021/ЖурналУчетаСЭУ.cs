using System;
using System.Collections.Generic;
using System.Text;
using RegKor.Classess;
using Excel = Microsoft.Office.Interop.Excel;

namespace RegKor.Classess2021
{
    public class ЖурналУчетаСЭУ : IPrintReport
    {
        //Объект Excel
        private Microsoft.Office.Interop.Excel.Application ObjExcel;

        //объект массив excel книг
        private Microsoft.Office.Interop.Excel.Workbooks ObjWorkBooks;

        //Объект excel книга
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;

        //объект excel лист
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

        private List<IntemЖурналУчетаСЭУ> list;

        private string beginDate = string.Empty;
        private string endDate = string.Empty;

        // 
        public ЖурналУчетаСЭУ(List<IntemЖурналУчетаСЭУ> list, string dataStart, string dataEnd)
        {
            if (list == null)
                throw new ArgumentNullException("Нет данных для отчета");
            this.list = list;

            beginDate =  ДатаSQL.SqlToДата(dataStart);

            endDate = ДатаSQL.SqlToДата(dataEnd);


        }

        public void Execute()
        {
            // Переменные для установки ширины колонки.
            int width1Column = 15;
            int width5Column = 30;// 40;// 20;
            int widthColumn = 50;
            int widthShortContColumn = 70;

            int ширинаСтроки = 90;
            int ширинаСтроки2 = 100;

            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

            //Книга.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);

            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            // Установим альбомную ориентацию бумаги.
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;

            // Зададим масштаб в 55%.
            ObjWorkSheet.PageSetup.Zoom = 65;

            // Установим отступы с лева и с права = 0.
            ObjWorkSheet.PageSetup.LeftMargin = 0;
            ObjWorkSheet.PageSetup.RightMargin = 0;

            // Установим отступ с низу и с вверху.
            ObjWorkSheet.PageSetup.TopMargin = 0;
            ObjWorkSheet.PageSetup.BottomMargin = 0;


            // Выровним по центру.
            ObjWorkSheet.PageSetup.CenterHorizontally = true;

            //Запишем шапку
            //Объеденим ячейки
            ObjWorkSheet.get_Range("E1", "F1").Merge(Type.Missing);
            ObjWorkSheet.get_Range("E1", "F1").Font.Size = 12;
            ObjWorkSheet.get_Range("E1", "F1").Font.Bold = true;
            ObjWorkSheet.get_Range("E1", Type.Missing).Value2 = "УТВЕРЖДЁН";

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("E1", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("E1", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // Запишем текст в ячейки E2 F2. и установим размер шрифта 12, не жирный
            ObjWorkSheet.get_Range("E2", "F2").Merge(Type.Missing);
            ObjWorkSheet.get_Range("E2", "F2").Font.Size = 12;
            ObjWorkSheet.get_Range("E2", "F2").Font.Bold = false;
            ObjWorkSheet.get_Range("E2", Type.Missing).Value2 = "Приказом и.о. директора\nГКУ СО \"ЦКСЗН Саратовской области\" " +
                                                                "\nот 09.01.2019 г. № 25";

            // Зададим ширину столбцов E и F.
            ObjWorkSheet.get_Range("E1", "E1").ColumnWidth = width5Column;
            ObjWorkSheet.get_Range("F1", "F1").ColumnWidth = width5Column;

            // Установим ширину строки.
            ObjWorkSheet.get_Range("E2", "E2").RowHeight = ширинаСтроки;
            ObjWorkSheet.get_Range("F2", "F2").RowHeight = ширинаСтроки;

            // Зададим ширину столбцов.
            ObjWorkSheet.get_Range("E2", Type.Missing).HorizontalAlignment = Excel.Constants.xlLeft;
            ObjWorkSheet.get_Range("E2", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ObjWorkSheet.get_Range("F1", Type.Missing).HorizontalAlignment = Excel.Constants.xlLeft;
            ObjWorkSheet.get_Range("F1", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // Запишем название журнала.
            ObjWorkSheet.get_Range("C5", "E5").Merge(Type.Missing);
            ObjWorkSheet.get_Range("C5", "E5").Font.Size = 12;
            ObjWorkSheet.get_Range("C5", "E5").Font.Bold = true;
            ObjWorkSheet.get_Range("C5", Type.Missing).Value2 = "Журнал учёта передачи персональных данных \n с " + beginDate + " по " + endDate + " ";

            // Зададим ширину столбцов E и F.
            ObjWorkSheet.get_Range("C1", "C1").ColumnWidth = width5Column;
            ObjWorkSheet.get_Range("D1", "D1").ColumnWidth = width5Column;
            ObjWorkSheet.get_Range("E1", "E1").ColumnWidth = width5Column;

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("C5", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("C5", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            ObjWorkSheet.get_Range("C5", "C5").RowHeight = ширинаСтроки2;

            // Запишем шапку таблицы.
            ObjWorkSheet.get_Range("A7", "A7").Merge(Type.Missing);
            ObjWorkSheet.get_Range("A7", Type.Missing).Value2 = "№ п/п";

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("A7", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("A7", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // Нарисуем границу.
            ExcelЯчейка A7 = new ExcelЯчейка();
            A7.ГраницаЯчейки("A7", "A7", ObjWorkSheet);

            // Запишем шапку таблицы.
            ObjWorkSheet.get_Range("B7", "B7").Merge(Type.Missing);
            ObjWorkSheet.get_Range("B7", Type.Missing).Value2 = "Сведения о запрашивающем лице \nили адресат";

            ObjWorkSheet.get_Range("B7", "B7").ColumnWidth = widthColumn;

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range("B7", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range("B7", Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // Нарисуем границу.
            ExcelЯчейка B7 = new ExcelЯчейка();
            B7.ГраницаЯчейки("B7", "B7", ObjWorkSheet);
            ObjWorkSheet.get_Range("B7", Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;

            string cell = "C7";
            ObjWorkSheet.get_Range(cell, cell).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cell, Type.Missing).Value2 = "Краткое содержание запроса или \nинициативной передачи ПД";

            ObjWorkSheet.get_Range(cell, cell).ColumnWidth = widthShortContColumn;
            ObjWorkSheet.get_Range(cell, cell).RowHeight = ширинаСтроки2;

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range(cell, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cell, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // Нарисуем границу.
            ExcelЯчейка C7 = new ExcelЯчейка();
            C7.ГраницаЯчейки(cell, cell, ObjWorkSheet);
            ObjWorkSheet.get_Range(cell, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;

            // Колонка D.
            string cellD = "D7";
            ObjWorkSheet.get_Range(cellD, cellD).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cellD, Type.Missing).Value2 = "Отметка о передаче или отказе в \nпередаче ПД";

            ObjWorkSheet.get_Range(cellD, cellD).ColumnWidth = width5Column;
            ObjWorkSheet.get_Range(cellD, cellD).RowHeight = ширинаСтроки2;

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range(cellD, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cellD, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // Нарисуем границу.
            ExcelЯчейка D7 = new ExcelЯчейка();
            D7.ГраницаЯчейки(cellD, cellD, ObjWorkSheet);
            ObjWorkSheet.get_Range(cellD, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;

            // Колонка E.
            string cellE = "E7";
            ObjWorkSheet.get_Range(cellE, cellE).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cellE, Type.Missing).Value2 = "Дата передачи (отказа в \nпередаче)ПД";

            ObjWorkSheet.get_Range(cellE, cellE).ColumnWidth = width5Column;
            ObjWorkSheet.get_Range(cellE, cellE).RowHeight = ширинаСтроки2;

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range(cellE, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cellE, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // Нарисуем границу.
            ExcelЯчейка E7 = new ExcelЯчейка();
            E7.ГраницаЯчейки(cellE, cellE, ObjWorkSheet);
            ObjWorkSheet.get_Range(cellE, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;


            // Колонка F.
            string cellF = "F7";
            ObjWorkSheet.get_Range(cellF, cellF).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cellF, Type.Missing).Value2 = "Основание передачи ПД \n(номер ответа на запрос)";

            ObjWorkSheet.get_Range(cellF, cellF).ColumnWidth = width5Column;
            ObjWorkSheet.get_Range(cellF, cellF).RowHeight = ширинаСтроки2;

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range(cellF, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cellF, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


            // Нарисуем границу.
            ExcelЯчейка F7 = new ExcelЯчейка();
            F7.ГраницаЯчейки(cellF, cellF, ObjWorkSheet);
            ObjWorkSheet.get_Range(cellF, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;

            // Начнём нумерацию строк с 8 строки, так как первые 7 строк заняты под шапку журнала.
            int iCount = 8;

            // Счётчик нумерации строк.
            int num = 1;

            // Узнаем количество пунктов в списке list.
            int countRowsReport = list.Count;

            // Заполняем таблицу данными.
            foreach (IntemЖурналУчетаСЭУ item in list)
            {
                // Пройдём по номерам столбцов.
                for (int i = 1; i <= 6; i++)
                {
                    // Получим букву обозначаюущую столбец.
                    string exclB = ExcelЯчейка.БукваКолонка(i);

                    switch (i)
                    {
                        case 1:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = num.ToString().Trim();


                            ExcelЯчейка excCel = new ExcelЯчейка();
                            excCel.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


                            break;
                        case 2:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.ОписаниеКорреспондента.Trim();
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).WrapText = true;

                            ExcelЯчейка excCelB = new ExcelЯчейка();
                            excCelB.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        case 3:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.КраткоеСодержание.Trim();
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).WrapText = true;

                            ExcelЯчейка excCelC = new ExcelЯчейка();
                            excCelC.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        case 4:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.НомерИсход.Trim();

                            ExcelЯчейка excCelD = new ExcelЯчейка();
                            excCelD.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        case 5:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.ДатаПоступления.Trim();

                            ExcelЯчейка excCelE = new ExcelЯчейка();
                            excCelE.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        // Ячейка F
                        case 6:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.ОснованиеПередачи.Trim();

                            ExcelЯчейка excCelF = new ExcelЯчейка();
                            excCelF.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                    }
                }

                num++;
                iCount++;
            }

            // Установим номер строки где будет размещён ФИО пользователя который сформировал отчёт, где countRowsReport - количество строк в списке с данными , а 11 - 8 строк в шапке + 3 строки отступить от таблицы.
            int numRow = countRowsReport + 11;

            string cellUsrE = "D" + numRow.ToString();
            string cellUsrF = "F" + numRow.ToString();

            ObjWorkSheet.get_Range(cellUsrE, cellUsrF).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cellUsrE, cellUsrF).Font.Size = 10;
            ObjWorkSheet.get_Range(cellUsrE, cellUsrF).Font.Bold = false; ;
            ObjWorkSheet.get_Range(cellUsrE, Type.Missing).Value2 = "Сформировал " +MyAplicationIdentity.GetUses();

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range(cellUsrE, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cellUsrE, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // Выведим документ на экран.
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;

        }
    }
}
