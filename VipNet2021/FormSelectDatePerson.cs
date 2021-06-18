using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
//using Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;



using RegKor.Classess;


namespace RegKor
{
    public partial class FormSelectDatePerson : Form
    {
        //Объект Excel
        private Microsoft.Office.Interop.Excel.Application ObjExcel;

        //объект массив excel книг
        private Microsoft.Office.Interop.Excel.Workbooks ObjWorkBooks;

        //Объект excel книга
        private Microsoft.Office.Interop.Excel.Workbook ObjWorkBook;

        //объект excel лист
        private Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;

        public FormSelectDatePerson()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string beginDate = this.dateTimePicker1.Value.ToShortDateString();
            string endDate = this.dateTimePicker2.Value.ToShortDateString();

            string fName = beginDate + "ЖурналУчетаПередачиПерсональныхДанных" + endDate;


            // Сформируем данные для выполнения отчёта.
            //string query = "SELECT [id_карточки] " +
            //               ",[ОписаниеКорреспондента] " +
            //               ",[ДатаИсхода] " +
            //               ",[ЦельПолученияПерсональныхДанных] " +
            //               ",[СоставПерсональныхДанных] " +
            //               ",[ОтметкаПолученияОтказ] " +
            //               ",[Причина] " +
            //               "FROM [ViewЖурнал] " +
            //               "where [ДатаИсхода] >= '" + beginDate + "' and [ДатаИсхода] <= '" + endDate + "' ";

            ПодключитьБД connect = new ПодключитьБД();
            string sConnect = connect.СтрокаПодключения();

            // Хранит данные для отчета.
            //List<ЖурналПерДанных> list = new List<ЖурналПерДанных>();

            // Библиотека для хранения данных для журнала.
            Dictionary<string, ЖурналПерДанных> dictionary = new Dictionary<string, ЖурналПерДанных>();


            using (SqlConnection con = new SqlConnection(sConnect))
            {
                con.Open();
                string query = "select * from dbo.ViewЖурналУчётаПерсональныхДанных2 " +
                               "where [Дата] >= '" + beginDate + "' and [Дата] <= '" + endDate + "' ";

                DataSet ds = new DataSet();

                 // Заполним DataSet.
                SqlDataAdapter da = new SqlDataAdapter(query, con);
                da.Fill(ds, "ViewЖурналУчётаПерсональныхДанных");

                // Получим данные из таблицы которые соответсвуют 
                System.Data.DataTable tab1 = ds.Tables["ViewЖурналУчётаПерсональныхДанных"];

                // Счётчик ключа для инициативных записей.
                //int iKey = 1;
                int iKey = 0;

                foreach (DataRow row in tab1.Rows)
                {
                    iKey = Convert.ToInt32(row["id_карточки"]);

                    if (Convert.ToInt32(row["id_карточки"]) == 44440)
                    {
                        string iTest = "test";
                    }

                    ЖурналПерДанных журнал = new ЖурналПерДанных();
                    
#region Старая релизация (Пока не стирать вдруг понадобится к ней вернуться).
                    //// Если письмо не инициативное (имеется id входящегно документа).
                    //if (row["id_ВходящегоДокумента"] != DBNull.Value)
                    //{

                    //    // ключ.
                    //    string key = row["НомерВходящий"].ToString().Trim();

                    //    журнал.ОписаниеКорреспондента = row["ОписаниеКорреспондента"].ToString().Trim();

                    //    //журнал.КраткоеСодержание = row["КраткоеСодержание"].ToString().Trim();

                    //    журнал.НомерИсходящий = row["НомерИсходящий"].ToString().Trim();

                    //    журнал.ДатаОтправки = Convert.ToDateTime(row["Дата"]).ToShortDateString();

                    //    //журнал.НомерВходящий = row["НомерВходящий"].ToString().Trim();

                    //    // Проверим есть ли ещё записи в связующей таблице [СвязующаяКарточкаВходящаяИсходящая] соответсвующие id карточки.
                    //    string numPoust = string.Empty;
                    //    string about = string.Empty;
                    //    numPoust = InputNumberPoust(con, ds, Convert.ToInt32(row["id_карточки"]), out about);
                    //    if (numPoust != "")
                    //    {
                    //        журнал.НомерВходящий = numPoust.Trim();
                    //        журнал.КраткоеСодержание = about.Trim();
                    //    }
                    //    else
                    //    {
                    //        журнал.НомерВходящий = row["НомерВходящий"].ToString().Trim();
                    //        журнал.КраткоеСодержание = row["КраткоеСодержание"].ToString().Trim();
                    //    }

                    //    try
                    //    {

                    //        // Получим значение тестовое.
                    //        ЖурналПерДанных журнал2 = журнал;

                    //        // добавим в библиотеку пункт журнала.
                    //        dictionary.Add(key, журнал);


                    //    }
                    //    catch
                    //    {

                    //        //Возмём запрос View 2 который у нас и заполним класс 

                    //        //ЖурналПерДанных журналCath = new ЖурналПерДанных();

                    //        //возмём id карточки в качестве фильтра

                    //        /*
                    //         * Если такой ключ уже существует в библиотеке, тогда запись с текущим ключом удаляется 
                    //         * и перезаписываем с новыми данными.
                    //        */

                    //        ЖурналПерДанных журналCath = new ЖурналПерДанных();

                    //        // Получим текущие данные из библиотеки.
                    //        string strTime = dictionary[key].ОписаниеКорреспондента.Trim();
                    //        журналCath.КраткоеСодержание = dictionary[key].КраткоеСодержание.Trim();
                    //        журналCath.НомерИсходящий = dictionary[key].НомерИсходящий.Trim();
                    //        журналCath.ДатаОтправки = dictionary[key].ДатаОтправки.Trim();
                    //        журналCath.НомерВходящий = dictionary[key].НомерВходящий.Trim();


                    //        string[] numN = key.Split('/');
                    //        string дата = ДатаSQL.Дата(DateTime.Now.Year.ToString()) + "0101";


                    //        string[] номерПорядковый = row["НомерИсходящийИнициативный"].ToString().Split('/');

                    //        // Узнаем наименование организации куда направлено письмо.

                    //        string quer = "declare @id_карточкаВходящая int " +
                    //                     "set @id_карточкаВходящая = 0 " +
                    //                     "select @id_карточкаВходящая = id_карточкаВходящая from dbo.СвязующаяКарточкаВходящаяИсходящая " +
                    //                     "where id_карточкаИсходящая in ( " +
                    //                     "select id_карточки from КарточкаИсходящая " +
                    //                     "where НомерПорядковый = '" + номерПорядковый[1] + "') " +
                    //                     "if(@id_карточкаВходящая = 0) " +
                    //                     "begin  " +
                    //                     "select distinct ОписаниеКорреспондента from Корреспонденты " +
                    //                     "where id_корреспондента in ( " +
                    //                     "select id_Адресата from КарточкаИсходящая " +
                    //            //"where НомерПорядковый = '6125') " +
                    //                     "where НомерПорядковый = '" + номерПорядковый[1] + "' and Дата >= '20140101') " +
                    //                     "end " +
                    //                     "else " +
                    //                     "begin  " +
                    //                     "select ОписаниеКорреспондента from dbo.Корреспонденты  " +
                    //                     "where id_корреспондента in  " +
                    //                     "( select id_корреспондента from Карточка  " +
                    //                     "where id_карточки in  " +
                    //                     "( select id_карточкаВходящая from dbo.СвязующаяКарточкаВходящаяИсходящая  " +
                    //                     "where id_карточкаИсходящая in  " +
                    //                     "( select id_карточки from dbo.ViewЖурналУчётаперсональныхДанных2  " +
                    //                     "where НомерВходящий = '" + row["НомерВходящий"].ToString().Trim() + "'))) " +
                    //                     "end ";

                    //        DataSet ds2 = new DataSet();

                    //        SqlDataAdapter da2 = new SqlDataAdapter(quer, con);
                    //        da2.Fill(ds2, "Проверка");

                    //        // Получим данные из таблицы которые соответсвуют 
                    //        System.Data.DataTable tab22 = ds2.Tables["Проверка"];

                    //        // Если количество записей в таблице корреспондентов больше 1
                    //        if (tab22.Rows.Count > 1)
                    //        {
                    //            StringBuilder buildCorr = new StringBuilder();

                    //            // Запишем описание корреспондента который был уже в 1 раз записан в библиотеку.
                    //            buildCorr.Append(strTime + ",");

                    //            foreach (DataRow r in tab22.Rows)
                    //            {
                    //                string корр = r["ОписаниеКорреспондента"].ToString().Trim();
                    //                buildCorr.Append(корр + ",");
                    //            }

                    //            // Уберём последную запятую.
                    //            int leng = buildCorr.Length;

                    //            string корреспондент = string.Empty;
                    //            корреспондент = buildCorr.Remove(leng - 1, 1).ToString().Trim();

                    //            журналCath.ОписаниеКорреспондента = корреспондент;

                    //            // Очистим временную таблицу от содержимого, чтобы в следующую строку не попали не нужные данные.
                    //            tab22.Clear();

                    //        }
                    //        else
                    //        {
                    //            журналCath.ОписаниеКорреспондента = row["ОписаниеКорреспондента"].ToString().Trim() + ", \n" + strTime.Trim();

                    //            string querSelect = "select CONVERT(nvarchar, НомерПП) + N'/' + RTRIM(LTRIM(CONVERT(nvarchar, " +
                    //                                "НомерВход))) AS 'Номер'  from Карточка " +
                    //                                "where id_карточки in ( " +
                    //                                "select id_карточкаВходящая from dbo.СвязующаяКарточкаВходящаяИсходящая " +
                    //                                "where id_карточкаИсходящая = " + Convert.ToInt32(row["id_карточки"]) + ")";

                    //            SqlDataAdapter daM = new SqlDataAdapter(querSelect, con);
                    //            daM.Fill(ds, "Входящие");

                    //            // Получим данные из таблицы которые соответсвуют 
                    //            System.Data.DataTable tabM = ds.Tables["Входящие"];

                    //            // Хранит итоговое значение.
                    //            StringBuilder buld = new StringBuilder();

                    //            foreach (DataRow r in tabM.Rows)
                    //            {
                    //                buld.Append(r["Номер"].ToString().Trim() + ", \n");
                    //            }

                    //            // Удалим последную запятую.
                    //            int countChar = buld.Length;
                    //            string numbers = string.Empty;

                    //            if (countChar > 0)
                    //            {
                    //                numbers = buld.Remove(countChar - 3, 3).ToString();
                    //                журналCath.НомерВходящий = numbers.ToString().Trim();
                    //            }
                    //        }

                    //        // Удалим запись.
                    //        //dictionary.Remove(key);

                    //        // Добавим изменённые данные.
                    //        dictionary.Add(key + 1, журналCath);
                    //    }
                    //}
#endregion

                    // Проверим инициативное письмо или нет.
                    if (row["id_ВходящегоДокумента"] != DBNull.Value)
                    {
                        // Получим ключ для билиотеки, в качестве ключа используем id карточки исходящей.
                        string key = row["id_карточки"].ToString().Trim();// row["НомерВходящий"].ToString().Trim();

                        журнал.ОписаниеКорреспондента = row["ОписаниеКорреспондента"].ToString().Trim();

                        журнал.НомерИсходящий = row["НомерИсходящий"].ToString().Trim();

                        журнал.ДатаОтправки = Convert.ToDateTime(row["Дата"]).ToShortDateString();

                        // Проверим есть ли ещё записи в связующей таблице [СвязующаяКарточкаВходящаяИсходящая] соответсвующие id карточки.
                        string numPoust = string.Empty;
                        string about = string.Empty;
                        numPoust = InputNumberPoust(con, ds, Convert.ToInt32(row["id_карточки"]), out about);
                        if (numPoust != "")
                        {
                            журнал.НомерВходящий = numPoust.Trim();
                            журнал.КраткоеСодержание = about.Trim();
                        }
                        else
                        {
                            журнал.НомерВходящий = row["НомерВходящий"].ToString().Trim();
                            журнал.КраткоеСодержание = row["КраткоеСодержание"].ToString().Trim();
                        }

                        try
                        {
                            // Получим значение тестовое.
                            ЖурналПерДанных журнал2 = журнал;

                            // добавим в библиотеку пункт журнала.
                            dictionary.Add(key, журнал);
                        }
                        catch
                        {

                            ЖурналПерДанных журналCath = new ЖурналПерДанных();

                            // Получим текущие данные из библиотеки.
                            string strTime = dictionary[key].ОписаниеКорреспондента.Trim();
                            журналCath.КраткоеСодержание = dictionary[key].КраткоеСодержание.Trim();
                            журналCath.НомерИсходящий = dictionary[key].НомерИсходящий.Trim();
                            журналCath.ДатаОтправки = dictionary[key].ДатаОтправки.Trim();
                            журналCath.НомерВходящий = dictionary[key].НомерВходящий.Trim();

                            журналCath.ОписаниеКорреспондента = row["ОписаниеКорреспондента"].ToString().Trim() + ", \n" + strTime.Trim();

                            dictionary.Remove(key);

                            dictionary.Add(key, журналCath);

                        }
                    }
                    else
                    {

                        журнал.ОписаниеКорреспондента = row["ОписаниеКорреспондента"].ToString().Trim();

                        // Если письмо инициативное.
                        журнал.КраткоеСодержание = row["СодержаниеИнициативное"].ToString().Trim();
                        журнал.НомерИсходящий = row["НомерИсходящийИнициативный"].ToString().Trim();
                        журнал.ДатаОтправки = Convert.ToDateTime(row["Дата"]).ToShortDateString();

                        // Получим номер порядковый.
                        string[] numArr = журнал.НомерИсходящий.Split('/');

                        if (Convert.ToInt32(row["id_карточки"]) == 46095)
                        {
                            string test = "Test";
                        }

                        string queryOP = "select ОснованиеПередачи from Основаниепередачи " +
                                         "where id_основаниеПередачи in ( " +
                                         "select id_ОснованиеПередачи from СвязующаяЦельПолучениперсональныхДанных " +
                                         "where id_карточки = " + Convert.ToInt32(row["id_карточки"]) + ") ";


                        SqlDataAdapter daOP = new SqlDataAdapter(queryOP, con);
                        daOP.Fill(ds, "ОснованиеПередачи");

                        // Получим данные из таблицы которые соответсвуют 
                        System.Data.DataTable tabOP = ds.Tables["ОснованиеПередачи"];

                        if (tabOP.Rows.Count != 0)
                        {
                            СтрокаОтчёта strOt = new СтрокаОтчёта(tabOP);
                            string strИниц = strOt.ConvertStringBuilder();

                            журнал.НомерВходящий = strИниц;

                            // Очистим таблицу.
                            ds.Tables["ОснованиеПередачи"].Clear();
                        }
                        else
                        {
                            журнал.НомерВходящий = "Инициативное";
                        }

                        // ключ.
                        //string key = row["НомерВходящий"].ToString().Trim();

                        //iKey++; - старый функционал.

                        //try
                        //{
                            // добавим в библиотеку пункт журнала.
                            dictionary.Add(iKey.ToString().Trim(), журнал);
                        //}
                        //catch
                        //{
                        //    //string strTime = dictionary[iKey.ToString().Trim()].ОписаниеКорреспондента.Trim();
                        //    iKey++;
                        //    dictionary.Add(iKey.ToString().Trim(), журнал);
                          
                        //}
                    }

                    //list.Add(журнал);
                }

            }

            // Переменные для установки ширины колонки.
            int width1Column = 15;
            int width5Column = 20;
            int widthColumn = 50;
            int widthShortContColumn = 70;

            int ширинаСтроки = 90;
            int ширинаСтроки2 = 50;

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
            ObjWorkSheet.PageSetup.Zoom = 70;

            // Установим отступы с лева и с права = 0.
            ObjWorkSheet.PageSetup.LeftMargin = 0;
            ObjWorkSheet.PageSetup.RightMargin = 0;

            // Установим отступ с низу и с вверху.
            ObjWorkSheet.PageSetup.TopMargin = 0;
            ObjWorkSheet.PageSetup.BottomMargin = 0;


            // Выровним по центру.
            ObjWorkSheet.PageSetup.CenterHorizontally = true;

            // Установим масштаб.
            //ObjExcel.ActiveWindow.Zoom = 50;



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
            ObjWorkSheet.get_Range("E2", Type.Missing).Value2 = "приказом и.о. директора \nГКУ СО \"ЦКСЗН Саратовской области \" \n" +
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
            ObjWorkSheet.get_Range("C5", Type.Missing).Value2 = "Журнал учёта передачи персональных данных \n с "+ beginDate +" по "+endDate+" ";

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
            ObjWorkSheet.get_Range("B7", Type.Missing).Value2 = "Сведения о запрашивающем лице";

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
            ObjWorkSheet.get_Range(cellF, Type.Missing).Value2 = "Основание передачи ПД \n(номер запроса)";

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
            int countRowsReport = dictionary.Values.Count;// list.Count;

            // Заполняем таблицу данными.
            foreach (ЖурналПерДанных item in dictionary.Values)
            {
                // Пройдём по номерам столбцов.
                for (int i = 1; i <= 6; i++)
                {
                    // Получим букву обозначаюущую столбец.
                    string exclB = ExcelЯчейка.БукваКолонка(i);

                    switch(i)
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
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.НомерИсходящий.Trim();

                            ExcelЯчейка excCelD = new ExcelЯчейка();
                            excCelD.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        case 5:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.ДатаОтправки.Trim();

                            ExcelЯчейка excCelE = new ExcelЯчейка();
                            excCelE.ГраницаЯчейки(exclB + iCount.ToString(), exclB + iCount.ToString(), ObjWorkSheet);
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

                            break;
                        // Ячейка F
                        case 6:
                            ObjWorkSheet.get_Range(exclB + iCount.ToString(), Type.Missing).Value2 = item.НомерВходящий.Trim();

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

            // Добавим пользователя который сформировал отчёт.

            // Установим номер строки где будет размещён ФИО пользователя который сформировал отчёт, где countRowsReport - количество строк в списке с данными , а 11 - 8 строк в шапке + 3 строки отступить от таблицы.
            int numRow = countRowsReport + 11;

            string cellUsrE = "D" + numRow.ToString();
            string cellUsrF = "F" + numRow.ToString();

            ObjWorkSheet.get_Range(cellUsrE, cellUsrF).Merge(Type.Missing);
            ObjWorkSheet.get_Range(cellUsrE, cellUsrF).Font.Size = 10;
            ObjWorkSheet.get_Range(cellUsrE, cellUsrF).Font.Bold = false; ;
            ObjWorkSheet.get_Range(cellUsrE, Type.Missing).Value2 = "Сформировал " + MyAplicationIdentity.GetUses();

            // Выровним текст по горизонтали.
            ObjWorkSheet.get_Range(cellUsrE, Type.Missing).HorizontalAlignment = Excel.Constants.xlCenter;
            ObjWorkSheet.get_Range(cellUsrE, Type.Missing).VerticalAlignment = Excel.Constants.xlCenter;

            // Выведим документ на экран.
            ObjExcel.Visible = true;
            ObjExcel.UserControl = true;


            #region Старая реализация WORD

            //Создаём новый Word.Application
    //        Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();


    //            //Загружаем документ
    //            Microsoft.Office.Interop.Word.Document doc = null;

    //            object fileName = filName;
    //            object falseValue = false;
    //            object trueValue = true;
    //            object missing = Type.Missing;
    //            object writePasswordDocument = "12A86Asd";

    //            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
    //ref missing, ref missing, ref missing, ref missing, ref writePasswordDocument,
    //ref missing, ref missing, ref missing, ref missing, ref trueValue,
    //ref missing, ref missing, ref missing);

    //            ////Дата начало отчёта.
    //            object wdrepl = WdReplace.wdReplaceAll;
    //            //object searchtxt = "GreetingLine";
    //            object searchtxt = "DATESTART";
    //            object newtxt = (object)beginDate;
    //            //object frwd = true;
    //            object frwd = false;
    //            doc.Content.Find.Execute(ref searchtxt, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd, ref missing, ref missing, ref newtxt, ref wdrepl, ref missing, ref missing,
    //            ref missing, ref missing);

    //            // Дата окончания отчтёта.
    //            object wdrepl2 = WdReplace.wdReplaceAll;
    //            //object searchtxt = "GreetingLine";
    //            object searchtxt2 = "DATEEND";
    //            object newtxt2 = (object)endDate;
    //            //object frwd = true;
    //            object frwd2 = false;
    //            doc.Content.Find.Execute(ref searchtxt2, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd2, ref missing, ref missing, ref newtxt2, ref wdrepl2, ref missing, ref missing,
    //            ref missing, ref missing);

    //            //Вставить таблицу
    //            object bookNaziv = "таблица";
    //            Range wrdRng = doc.Bookmarks.get_Item(ref  bookNaziv).Range;

    //            object behavior = Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord8TableBehavior;
    //            object autobehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow;


    //            Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(wrdRng, 1, 6, ref behavior, ref autobehavior);
    //            table.Range.ParagraphFormat.SpaceAfter = 11;

    //            table.Columns[1].Width = 40;
    //            table.Columns[2].Width = 150;
    //            table.Columns[3].Width = 150;
    //            table.Columns[4].Width = 150;
    //            table.Columns[5].Width = 120;
    //            table.Columns[6].Width = 120;

    //            table.Borders.Enable = 1; // Рамка - сплошная линия
    //            table.Range.Font.Name = "Times New Roman";
    //            table.Range.Font.Size = 9;

    //            // Запишем шапку таблицы.
    //            table.Cell(1, 1).Range.Text = "№ п/п";
    //            table.Cell(1, 2).Range.Text = "Сведения о запрашивающем лице";
    //            table.Cell(1, 3).Range.Text = "Краткое содержание запроса или инициативной передачи ПД";
    //            table.Cell(1, 4).Range.Text = "Отметка о передаче или отказе в передаче ПД";
    //            table.Cell(1, 5).Range.Text = "Дата передачи(отказа в передаче)ПД";
    //            table.Cell(1, 6).Range.Text = "Основание передачи ПД (номер запроса)";

    //            Object beforeRow1 = Type.Missing;
    //            table.Rows.Add(ref beforeRow1);


    //            int count = 1;

    //            // Заполним таблицу данными.
    //            foreach (ЖурналПерДанных item in list)
    //            {
    //                table.Cell(count+1, 1).Range.Text = count.ToString().Trim();
    //                table.Cell(count+1, 2).Range.Text = item.ОписаниеКорреспондента.Trim();
    //                table.Cell(count + 1, 3).Range.Text = item.КраткоеСодержание.Trim();
    //                table.Cell(count + 1, 4).Range.Text = item.НомерИсходящий.Trim();
    //                table.Cell(count + 1, 5).Range.Text = item.ДатаОтправки.Trim();
    //                table.Cell(count + 1, 6).Range.Text = item.НомерВходящий.Trim();

    //                Object beforeRow2 = Type.Missing;
    //                table.Rows.Add(ref beforeRow2);

    //                count++;
    //            }

    //            //удалим последную строку
    //            table.Rows[count+1].Delete();

    //            // Добавим кто сформировал отчёт.
    //            string user = MyAplicationIdentity.GetUses();

    //            string[] arryFIO = user.Split(' ');
    //            string инициалыИмя = arryFIO[1].Substring(0, 1);
    //            string инициалыОтчество = arryFIO[2].Substring(0, 1);
    //            string fio = arryFIO[0] + " " + инициалыИмя + "." + " " + инициалыОтчество + ".";

    //            // Дата окончания отчтёта.
    //            object wdrepl3 = WdReplace.wdReplaceAll;
    //            //object searchtxt = "GreetingLine";
    //            object searchtxt3 = "USER";
    //            object newtxt3 = (object)fio;
    //            //object frwd = true;
    //            object frwd3 = false;
    //            doc.Content.Find.Execute(ref searchtxt3, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd3, ref missing, ref missing, ref newtxt3, ref wdrepl3, ref missing, ref missing,
    //            ref missing, ref missing);



    //            // Отобрпазим документ и закроем окно.
            //            app.Visible = true;
            #endregion

            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Возвращает номера входящих писем.
        /// </summary>
        /// <param name="con"></param>
        /// <param name="ds"></param>
        /// <param name="idКарточки"></param>
        /// <returns></returns>
        private string InputNumberPoust(SqlConnection con,DataSet ds, int idКарточки, out string shortAbout)
        {
            string querSelect = "select CONVERT(nvarchar, НомерПП) + N'/' + RTRIM(LTRIM(CONVERT(nvarchar, " +
                                                    "НомерВход))) AS 'Номер',КраткоеСодержание  from Карточка " +
                                                    "where id_карточки in ( " +
                                                    "select id_карточкаВходящая from dbo.СвязующаяКарточкаВходящаяИсходящая " +
                                                    "where id_карточкаИсходящая = " + Convert.ToInt32(idКарточки) + ")";

            SqlDataAdapter daM = new SqlDataAdapter(querSelect, con);
            daM.Fill(ds, "Входящие");

            // Получим данные из таблицы которые соответсвуют 
            System.Data.DataTable tabM = ds.Tables["Входящие"];

            // Хранит итоговое значение.
            StringBuilder buld = new StringBuilder();

            // Хранит описание письма.
            StringBuilder buldAbout = new StringBuilder();

            foreach (DataRow r in tabM.Rows)
            {
                buld.Append(r["Номер"].ToString().Trim() + ", \n");
                buldAbout.Append(r["КраткоеСодержание"].ToString().Trim() + ", \n");
            }

            // Удалим последную запятую.
            int countChar = buld.Length;
            string numbers = string.Empty;

            if (countChar > 0)
            {
                numbers = buld.Remove(countChar - 3, 3).ToString();
            }

            // Удалим последную запятую из описания.
            int countAbout = buldAbout.Length;
            string aboutS = string.Empty;

            // Присвоим пустую строку.
            shortAbout = "";

            if (countAbout > 0)
            {
                aboutS = buldAbout.Remove(countAbout - 3, 3).ToString();
                shortAbout = aboutS;
            }


            ds.Tables["Входящие"].Clear();

            return numbers.ToString().Trim();
        }
    }
}