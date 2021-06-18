using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace RegKor.Classess
{
    public class ExcelОтчет
    {
        /// <summary>
        /// Вывод отчета на печать, где в качестве источник с данными содержится класс с DataTable.
        /// </summary>
        /// <param name="tableData">Класс содержащий DataTable с данными для отчета.</param>
        /// <param name="captionText">Текст для заголовка отчета.</param>
        /// <param name="captionCellStart">Адрес первой ячейки отчета</param>
        /// <param name="captionCellEnd">Адрес крайней ячейки отчёта</param>
        public void PrintОтчетOfDataTable(IОтчет tableData, string captionText, string captionCellStart, string captionCellEnd)
        {


            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

           

            //Книга.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);

            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            // Зададим ориентацию.
            ObjWorkSheet.PageSetup.Zoom = false;
            ObjWorkSheet.PageSetup.FitToPagesWide = 1;
            ObjWorkSheet.PageSetup.FitToPagesTall = 800;
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;

            //Объеденим ячейки
            ObjWorkSheet.get_Range(captionCellStart, captionCellEnd).Merge(Type.Missing);
            ObjWorkSheet.get_Range(captionCellStart, captionCellEnd).Font.Size = 12;
            ObjWorkSheet.get_Range(captionCellStart, captionCellEnd).Font.Bold = true;
            ObjWorkSheet.get_Range(captionCellStart, Type.Missing).Value2 = captionText;
            ObjWorkSheet.get_Range(captionCellStart, captionCellEnd).HorizontalAlignment = Excel.Constants.xlCenter;

            ОтчетОВходДокументах report = (ОтчетОВходДокументах)tableData;

            int iCount = 3;

           

            // Сворганим шапку.
                this.Cell(ObjWorkSheet, "A", iCount, Convert.ToString("№ п.п"), 5);
                this.Cell(ObjWorkSheet, "B", iCount, "Корреспондент", 20);
                this.Cell(ObjWorkSheet, "C", iCount, "Дата исходящая", 10);
                this.Cell(ObjWorkSheet, "D", iCount, "Номер исходящий", 15);
                this.Cell(ObjWorkSheet, "E", iCount, "Краткое содержание", 30);
                this.Cell(ObjWorkSheet, "F", iCount, "Дата входящая", 10);
                this.Cell(ObjWorkSheet, "G", iCount, "Номер входящий", 10);
                this.Cell(ObjWorkSheet, "H", iCount, "Срок исполнения", 15);
                this.Cell(ObjWorkSheet, "I", iCount, "Результат исполнения", 20);
                this.Cell(ObjWorkSheet, "J", iCount, "Исполнитель", 15);

                iCount++;

            foreach(DataGridViewRow r in report.DataGridView1.Rows)
            {
                int iCountCell = 1;

                foreach(DataGridViewCell cell in r.Cells)
                {
                    if (cell.Value != null)
                    {
                        if (iCountCell == 1)
                        {
                            this.Cell(ObjWorkSheet, "A", iCount, Convert.ToString(iCount-3), 5);
                            this.Cell(ObjWorkSheet, "B", iCount, cell.Value.ToString(), 20);
                       }

                        if (iCountCell == 2)
                        {
                            this.Cell(ObjWorkSheet, "C", iCount, Convert.ToDateTime(cell.Value.ToString()).ToShortDateString(), 10);
                        }

                        if (iCountCell == 3)
                        {
                            this.Cell(ObjWorkSheet, "D", iCount, cell.Value.ToString(), 15);
                        }

                        if (iCountCell == 4)
                        {
                            this.Cell(ObjWorkSheet, "E", iCount, cell.Value.ToString(), 30);
                        }

                        if (iCountCell == 5)
                        {
                            this.Cell(ObjWorkSheet, "F", iCount, Convert.ToDateTime(cell.Value.ToString()).ToShortDateString(), 10);
                        }

                        if (iCountCell == 6)
                        {
                            this.Cell(ObjWorkSheet, "G", iCount, cell.Value.ToString(), 10);
                        }

                        if (iCountCell == 7)
                        {
                            if (cell.Value.ToString().Trim().Length  != 0)
                            {
                                this.Cell(ObjWorkSheet, "H", iCount, Convert.ToDateTime(cell.Value.ToString()).ToShortDateString(), 15);
                            }
                            else
                            {
                                this.Cell(ObjWorkSheet, "H", iCount, "", 15);
                            }
                        }

                        if (iCountCell == 8)
                        {
                            this.Cell(ObjWorkSheet, "I", iCount, cell.Value.ToString(), 20);
                        }

                        if (iCountCell == 9)
                        {
                            this.Cell(ObjWorkSheet, "J", iCount, cell.Value.ToString(), 15);
                        }

                        iCountCell++;
                    }
                    else
                    {
                        break;
                    }
                }

                iCount++;
                
            }

            // Отобразим документ.
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;

        }

        /// <summary>
        /// Ячейка Excel.
        /// </summary>
        /// <param name="ObjWorkSheet"></param>
        /// <param name="exclБукв"></param>
        /// <param name="iCount"></param>
        /// <param name="value"></param>
        private void Cell(Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet, string exclБукв, int iCount, string value, int columhWidth)
        {
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).Value2 = value;
            // выровним горизонтально.
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            // выровним вертикально.
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // Перенос текста.
            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), Type.Missing).WrapText = true;

            ObjWorkSheet.get_Range(exclБукв + iCount.ToString(), exclБукв + iCount.ToString()).ColumnWidth = columhWidth;

            // Нарисуем границу.
            ExcelЯчейка excNum = new ExcelЯчейка();
            excNum.ГраницаЯчейки(exclБукв + iCount.ToString(), exclБукв + iCount.ToString(), ObjWorkSheet);

        }

        public void PrintОтчетИсходящиеДокументы(IОтчет tableData, string captionText, string captionCellStart, string captionCellEnd)
        {
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

            //Книга.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);

            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            // Альбомное расположение листа.
            ObjWorkSheet.PageSetup.Zoom = false; ;
            ObjWorkSheet.PageSetup.FitToPagesWide = 1;
            ObjWorkSheet.PageSetup.FitToPagesTall = 800;
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;


            //Объеденим ячейки
            ObjWorkSheet.get_Range(captionCellStart, captionCellEnd).Merge(Type.Missing);
            ObjWorkSheet.get_Range(captionCellStart, captionCellEnd).Font.Size = 12;
            ObjWorkSheet.get_Range(captionCellStart, captionCellEnd).Font.Bold = true;
            ObjWorkSheet.get_Range(captionCellStart, Type.Missing).Value2 = captionText;
            ObjWorkSheet.get_Range(captionCellStart, captionCellEnd).HorizontalAlignment = Excel.Constants.xlCenter;

            ОтчетОВходДокументах report = (ОтчетОВходДокументах)tableData;

            int iCount = 3;



            // Сворганим шапку.
            this.Cell(ObjWorkSheet, "A", iCount, Convert.ToString("№ п.п"), 5);
            this.Cell(ObjWorkSheet, "B", iCount, "Адресат", 40);
            this.Cell(ObjWorkSheet, "C", iCount, "Дата исходящая", 10);
            this.Cell(ObjWorkSheet, "D", iCount, "Номер исходящий", 25);
            this.Cell(ObjWorkSheet, "E", iCount, "Краткое содержание", 30);
            this.Cell(ObjWorkSheet, "F", iCount, "Исполнитель", 20);
            this.Cell(ObjWorkSheet, "G", iCount, "Документ на который дан ответ", 60);

            iCount++;

            foreach (DataGridViewRow r in report.DataGridView1.Rows)
            {
                int iCountCell = 1;

                foreach (DataGridViewCell cell in r.Cells)
                {
                    if (cell.Value != null)
                    {
                        if (iCountCell == 3)
                        {
                            this.Cell(ObjWorkSheet, "A", iCount, Convert.ToString(iCount - 3), 5);
                            this.Cell(ObjWorkSheet, "B", iCount, cell.Value.ToString(), 40);
                        }

                        if (iCountCell == 4)
                        {
                            string asd = cell.Value.ToString();
                            this.Cell(ObjWorkSheet, "C", iCount, Convert.ToDateTime(cell.Value.ToString()).ToShortDateString(), 10);
                        }

                        if (iCountCell ==5)
                        {
                            this.Cell(ObjWorkSheet, "D", iCount, cell.Value.ToString(), 25);
                        }

                        if (iCountCell == 6)
                        {
                            this.Cell(ObjWorkSheet, "E", iCount, cell.Value.ToString(), 30);
                        }

                        if (iCountCell == 7)
                        {
                            this.Cell(ObjWorkSheet, "F", iCount, cell.Value.ToString(), 20);
                        }
                        if (iCountCell == 8)
                        {
                            this.Cell(ObjWorkSheet, "G", iCount, cell.Value.ToString(), 20);
                        }
                    }

                    iCountCell++;
                }

                iCount++;
            }

            // Отобразим документ.
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;

            iCount++;
        }
    }
}
