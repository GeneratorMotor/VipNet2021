using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;

using Excel = Microsoft.Office.Interop.Excel;

namespace RegKor
{
    public partial class FormДокументооборот : Form
    {
        public FormДокументооборот()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormДокументооборот_Load(object sender, EventArgs e)
        {
            // Установим по умолчанию распечатаь все документы.
            this.radButAll.Checked = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Переменная для хранения строки запроса к БД.
            string query = string.Empty;

            // Переменные для хранения диапазонов отчёта.
            string beginDate = ДатаSQL.Дата(this.dt1.Value.ToShortDateString());
            string endDate = ДатаSQL.Дата(this.dt2.Value.ToShortDateString());

            // Сформируем строку запроса.
            if (radButAll.Checked == true)
            {
                // Выведим все документы.
                query = "SELECT [Регистрационный номер документа] " +
                      ",[ДатаПоступ] " +
                      ",[КраткоеСодержание] " +
                      ",[NameFileDocument] " +
                      ",[GuidName] " +
                      ",[ОписаниеКорреспондента] " +
                      ",[СрокВыполнения] " +
                      ",[ОтметкаПрочтение] " +
                      ",[РезультатВыполнения] " +
                      ",[ОписаниеПолучателя] " +
                      "FROM [ViewДокументооборот] " +
                      "where [ДатаПоступ] >= '"+ beginDate +"' and [ДатаПоступ] <= '"+ endDate +"' ";
            }
            else if (radButtonRead.Checked == true)
            {
                // выведим только прочитанные документы.
                query = "SELECT [Регистрационный номер документа] " +
                     ",[ДатаПоступ] " +
                     ",[КраткоеСодержание] " +
                     ",[NameFileDocument] " +
                     ",[GuidName] " +
                     ",[ОписаниеКорреспондента] " +
                     ",[СрокВыполнения] " +
                     ",[ОтметкаПрочтение] " +
                     ",[РезультатВыполнения] " +
                     ",[ОписаниеПолучателя] " +
                     "FROM [ViewДокументооборот] " +
                      "where [ДатаПоступ] >= '" + beginDate + "' and [ДатаПоступ] <= '" + endDate + "' and " +
                     "ОтметкаПрочтение is not null ";

            }
            else if (radButtonNoRead.Checked == true)
            { 
                // Выведим только не прочитанные документы.
                query = "SELECT [Регистрационный номер документа] " +
                     ",[ДатаПоступ] " +
                     ",[КраткоеСодержание] " +
                     ",[NameFileDocument] " +
                     ",[GuidName] " +
                     ",[ОписаниеКорреспондента] " +
                     ",[СрокВыполнения] " +
                     ",[ОтметкаПрочтение] " +
                     ",[РезультатВыполнения] " +
                     ",[ОписаниеПолучателя] " +
                     "FROM [ViewДокументооборот] " +
                      "where [ДатаПоступ] >= '" + beginDate + "' and [ДатаПоступ] <= '" + endDate + "' and " +
                     "ОтметкаПрочтение is null ";
            }

            ПодключитьБД connect = new ПодключитьБД();
            string sConnect = connect.СтрокаПодключения();

            // Список для хранения данных для отчёта.
            List<PrintДокументооборот> list = new List<PrintДокументооборот>();

            // Переменная типа DataTable для хранения данных для отчёта.
            System.Data.DataTable tab1;

            using (SqlConnection con = new SqlConnection(sConnect))
            {
                con.Open();
                
                DataSet ds = new DataSet();

                // Заполним DataSet.
                SqlDataAdapter da = new SqlDataAdapter(query, con);
                da.Fill(ds, "ViewДокументооборот");

                // Получим данные из таблицы которые соответсвуют 
                tab1 = ds.Tables["ViewДокументооборот"];

            }

            // Заполним шапку.
            PrintДокументооборот cap = new PrintДокументооборот();
            cap.РегистрационныйHомерДокумента = "Регистрационный номер документа";
            cap.ДатаПоступ = "Дата регистрации";
            cap.КраткоеСодержание = "Краткое содержание";
            cap.NameFileDocument = "Префикс файла";
            cap.GuidName = "Guid файла";
            cap.ОписаниеКорреспондента = "Описание корреспондента";
            cap.СрокВыполнения = "Срок выполнения";
            cap.ОтметкаПрочтение = "Отметка о прочтении";
            cap.РезультатВыполнения = "Результат выполнения";
            cap.ОписаниеПолучателя = "Описание получателя";

            list.Add(cap);


            // Заполним список list.
            foreach (DataRow row in tab1.Rows)
            {

                PrintДокументооборот item = new PrintДокументооборот();
                item.РегистрационныйHомерДокумента = row["Регистрационный номер документа"].ToString().Trim();
                item.ДатаПоступ = Convert.ToDateTime(row["ДатаПоступ"]).ToShortDateString().Trim();
                item.КраткоеСодержание = row["КраткоеСодержание"].ToString().Trim();
                item.NameFileDocument = row["NameFileDocument"].ToString().Trim();
                item.GuidName = row["GuidName"].ToString().Trim();
                item.ОписаниеКорреспондента = row["ОписаниеКорреспондента"].ToString().Trim();
                item.СрокВыполнения = Convert.ToDateTime(row["СрокВыполнения"]).ToShortDateString().Trim();
                item.ОтметкаПрочтение = row["ОтметкаПрочтение"].ToString().Trim();
                item.РезультатВыполнения = row["РезультатВыполнения"].ToString().Trim();
                item.ОписаниеПолучателя = row["ОписаниеПолучателя"].ToString().Trim();

                list.Add(item);

            }

            List<PrintДокументооборот> listTest = list;

            // Выведим информацию в Excel.
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

            //Книга.
            ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);

            //Таблица.
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

            // Установим альбомную ориентацию бумаги.
            ObjWorkSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;

            // Зададим масштаб в 70%.
            ObjWorkSheet.PageSetup.Zoom = 70;

            // Установим отступы с лева и с права = 0.
            ObjWorkSheet.PageSetup.LeftMargin = 0;
            ObjWorkSheet.PageSetup.RightMargin = 0;

            // Установим отступ с низу и с вверху.
            ObjWorkSheet.PageSetup.TopMargin = 0;
            ObjWorkSheet.PageSetup.BottomMargin = 0;


            // Выровним по центру.
            ObjWorkSheet.PageSetup.CenterHorizontally = true;

            // Начнём нумерацию строк с 1 строки, так как первые 7 строк заняты под шапку журнала.
            int iCount = 1;
            int num = 0;

            foreach (PrintДокументооборот it in list)
            {
                // Пройдём по номерам столбцов.
                for (int i = 1; i <= 9; i++)
                {

                    // Получим букву обозначаюущую столбец.
                    string exclB = ExcelЯчейка.БукваКолонка(i);

                    switch (i)
                    {
                        case 1:

                            if (iCount == 1)
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = "№ п.п";
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Font.Bold = true;
                            }
                            else
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = num.ToString().Trim();
                            }


                            ExcelЯчейка excCel = new ExcelЯчейка();
                            excCel.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;


                            break;

                        case 2:
                            if (iCount == 1)
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.РегистрационныйHомерДокумента.Trim();
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Font.Bold = true;
                            }
                            else
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.РегистрационныйHомерДокумента.Trim();
                            }


                            ExcelЯчейка excCel2 = new ExcelЯчейка();
                            excCel2.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).ColumnWidth = 40;


                            break;

                        case 3:
                            if (iCount == 1)
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.ДатаПоступ.Trim();
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Font.Bold = true;
                            }
                            else
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.ДатаПоступ.Trim();
                            }


                            ExcelЯчейка excCel3 = new ExcelЯчейка();
                            excCel3.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).ColumnWidth = 30;


                            break;

                        case 4:
                            if (iCount == 1)
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.КраткоеСодержание.Trim();
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Font.Bold = true;
                            }
                            else
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.КраткоеСодержание.Trim();
                            }


                            ExcelЯчейка excCel4 = new ExcelЯчейка();
                            excCel4.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).ColumnWidth = 50;


                            break;

                        case 5:
                            if (iCount == 1)
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.ОписаниеКорреспондента.Trim();
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Font.Bold = true;
                            }
                            else
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.ОписаниеКорреспондента.Trim();
                            }

                            ExcelЯчейка excCel5 = new ExcelЯчейка();
                            excCel5.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).ColumnWidth = 40;


                            break;

                        case 6:
                            if (iCount == 1)
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.СрокВыполнения.Trim();
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Font.Bold = true;
                            }
                            else
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.СрокВыполнения.Trim();
                            }


                            ExcelЯчейка excCel6 = new ExcelЯчейка();
                            excCel6.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).ColumnWidth = 17;


                            break;


                        case 7:
                            if (iCount == 1)
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.ОтметкаПрочтение.Trim();
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Font.Bold = true;
                            }
                            else
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.ОтметкаПрочтение.Trim();
                            }


                            ExcelЯчейка excCel7 = new ExcelЯчейка();
                            excCel7.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).ColumnWidth = 22;


                            break;

                        case 8:
                            if (iCount == 1)
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.РезультатВыполнения.Trim();
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Font.Bold = true;
                            }
                            else
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.РезультатВыполнения.Trim();
                            }


                            ExcelЯчейка excCel8 = new ExcelЯчейка();
                            excCel8.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).ColumnWidth = 40;


                            break;


                        case 9:
                            if (iCount == 1)
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.ОписаниеПолучателя.Trim();
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Font.Bold = true;
                            }
                            else
                            {
                                ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = it.ОписаниеПолучателя.Trim();
                            }


                            ExcelЯчейка excCel9 = new ExcelЯчейка();
                            excCel9.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).ColumnWidth = 40;


                            break;


                    }
                }

                iCount++;
                num++;
            }







            // Выведим документ на экран.
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;




            this.Close();


        }
    }
}