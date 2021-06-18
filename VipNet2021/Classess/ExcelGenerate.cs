using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace RegKor.Classess
{
    public class ExcelGenerate
    {
        // Список начальников отделов и управлений.
        IСписокНачальников _listDirectors;

        IProcessGetDoc _listProcessDoc;

        // Переменная для хранения экземпляра книги.
        private CarlosAg.ExcelXmlWriter.Workbook book2;

        // Переменная для хранения стиля ячейки.
        CarlosAg.ExcelXmlWriter.WorksheetStyle style;

        private Excel.Sheets excelsheets;
        private Excel.Worksheet excelworksheet;
        private Excel.Range excelcellsA;
        private Excel.Range excelcellsB;
        private Excel.Range excelcellsC;
        private Excel.Range excelcellsD;
        private Excel.Range excelcellsE;
        private Excel.Range excelcellsF;
        private Excel.Range excelcellsG;
        private Excel.Range excelcellsH;
        private Excel.Range excelcellsI;
        private Excel.Range excelcellsJ;
        private Excel.Range excelcellsK;
        private Excel.Range excelcellsL;
        private Excel.Range excelcellsM;
        private Excel.Range excelcellsN;
        private Excel.Range excelcellsO;


        private Excel.Workbooks excelappworkbooks;
        private Excel.Workbook excelappworkbook;

        private int startMonth = 0;
        private int endMonth = 0;

        private int year;

        /// <summary>
        /// Выбранный год для отчета.
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

        /// <summary>
        /// Превый месяц отчета.
        /// </summary>
        public int StartMonth
        {
            get
            {
                return startMonth;
            }
            set
            {
                startMonth = value;
            }
        }

        /// <summary>
        /// Крайний месяц отчета.
        /// </summary>
        public int EndMonth
        {
            get
            {
                return endMonth;
            }
            set
            {
                endMonth = value;
            }
        }

        public ExcelGenerate(IСписокНачальников listDirectors, IProcessGetDoc listProcessDoc)
        {
            _listDirectors = listDirectors;
            _listProcessDoc = listProcessDoc;
        }

        public void CreateExcel(List<YearMonth> listYearMonth, int countColumn)
        {
            _listDirectors = new СписокНачальников();
            _listProcessDoc = new ProcessGetDoc();

            List<string> listDir = _listDirectors.GetDirectors();

            List<string> listProcess = _listProcessDoc.GetDoc();

            // Библиотека строк для отчета.
            Dictionary<int, Dictionary<string,List<DocExcelCell>>> listExcelCell = new Dictionary<int, Dictionary<string,List<DocExcelCell>>>();

            // Список способов поступления документов для текущего начальника 
            List<PersonCell> listCell = new List<PersonCell>();

            foreach (string strName in listDir)
            {
                PersonCell pCell = new PersonCell();

                pCell.Name = strName.Trim();

                foreach (string strProc in listProcess)
                {
                    pCell.СпособПолучДокумент.Add(strProc);
                }

                listCell.Add(pCell);
            }

            // Добавим итоговые значения.
            PersonCell pCellCount = new PersonCell();

            pCellCount.Name = "Итого";
            pCellCount.СпособПолучДокумент.Add("Бумажный носитель");
            pCellCount.СпособПолучДокумент.Add("E-mail");
            pCellCount.СпособПолучДокумент.Add("VipNet");
            pCellCount.СпособПолучДокумент.Add("ФАКС");

            listCell.Add(pCellCount);

            // Библиотека для хранения данных полученных из БД по статисике поступления документов.
            Dictionary<int, List<ItemStatisticDoc>> dictionary = new Dictionary<int, List<ItemStatisticDoc>>();

            // Получим информацию из БД о статистики запросов по месяцам.
            foreach (YearMonth itYearMonth in listYearMonth)
            {
                List<ItemStatisticDoc> listStatic = СтатистикаПолученияДокументов.GetStatisticDoc(itYearMonth.Year, itYearMonth.NumMonth);
                dictionary.Add(itYearMonth.NumMonth, listStatic);
            }

            // Итоговое значение за отчет.
            List<ItemStatisticDoc> listStaticCount = СтатистикаПолученияДокументов.GetStatisticDocCount(this.Year, this.StartMonth, this.EndMonth);

            // Установим для строки ИТОГО номер месяца 13 (так как 13 месяца небывает).
            dictionary.Add(13, listStaticCount);

            Dictionary<int, List<ItemStatisticDoc>> dictionaryTest = dictionary;

            //Сформируем данные для отчета.
            // Для этого пройдем по списку способов получения документов.

            // Счетчик циклов.
            int count = 1;

            // Список содержит ячейки с ФИО начальников (ключ ФИО).
            Dictionary<string, List<DocExcelCell>> strExcelFio = new Dictionary<string, List<DocExcelCell>>();

            // Список содержит способы поступления документов (ключ Спопосб поступления документов). 
            Dictionary<string,List<DocExcelCell>> strExcelProcessDocGet = new Dictionary<string,List<DocExcelCell>>();

            #region Первая строка отчета с ФИО начальников

            // Получим первую строку отчета которая сдержит ФИО начальников управлений и отделов.
            foreach (PersonCell pCellFio in listCell)
            {
                // Запишем фамилии начальников отделов и управлений.
                DocExcelCell dCF = new DocExcelCell();

                // Установим количество способов получения документов.
                dCF.CountColumn = countColumn;
                dCF.ValueCell = pCellFio.Name.Trim();

                List<DocExcelCell> listdCF = new List<DocExcelCell>();
                listdCF.Add(dCF);

                strExcelFio.Add(pCellFio.Name.Trim(), listdCF);

                foreach(string str in pCellFio.СпособПолучДокумент)
                {
                    // Список содержит названия процессов.
                    List<DocExcelCell> listPN = new List<DocExcelCell>();

                    // Запишем способы поступления документов.
                    DocExcelCell cellProcessName = new DocExcelCell();
                    cellProcessName.CountColumn = 1;
                    cellProcessName.ValueCell = str;

                    listPN.Add(cellProcessName);

                    strExcelProcessDocGet.Add(pCellFio.Name.Trim()+ "_" + str.Trim(), listPN);
                }
            }

            // Добавим строку с фамилиями в отчет.
            listExcelCell.Add(count * 20 , strExcelFio);

            #endregion

            // Добавим вторую строку которая содержит название способов поступления документов.
            count++;

            listExcelCell.Add(count * 20, strExcelProcessDocGet);

            Dictionary<int, Dictionary<string, List<DocExcelCell>>> iTest = listExcelCell;

            #region Запишем значения ячеек из базы данных.

            //// Строка содержит данные для строки итого.
            Dictionary<string, List<DocExcelCell>> strExcelProcessDocGetPrintCount = new Dictionary<string, List<DocExcelCell>>();

            DocExcelCell dcCellCount = new DocExcelCell();
            
            // Запишем значения ячеек из базы данных.
            foreach (int month in dictionary.Keys)
            {
                List<ItemStatisticDoc> list = dictionary[month];

                // === тест
                Dictionary<string, List<DocExcelCell>> strExcelProcessDocGetPrint = new Dictionary<string, List<DocExcelCell>>();

                // Короче здесь чере foreach записываем в strExcelProcessDocGetPrint данные из strExcelProcessDocGet

                foreach (string itKey in strExcelProcessDocGet.Keys)
                {
                    List<DocExcelCell> listToCell = new List<DocExcelCell>();
                    foreach (DocExcelCell dCell in strExcelProcessDocGet[itKey])
                    {
                        DocExcelCell dcCell = new DocExcelCell();
                        dcCell.CountColumn = dCell.CountColumn;
                        dcCell.FlagEdit = false;
                        dcCell.ValueCell = dCell.ValueCell;
                        listToCell.Add(dcCell);

                    }

                    strExcelProcessDocGetPrint.Add(itKey, listToCell);

                  
                }

                // ==== конец тест

                // Создадим копию строки содержащий значения ячеек способов поступления документов по ФИО начальников отделов и служб.
                Dictionary<string, List<DocExcelCell>> listValue = strExcelProcessDocGet;

               
                // Пройдемся по выгрузке из БД.
                foreach (ItemStatisticDoc item in list)
                {
                    string strKey = item.ФИО.Trim() + "_" + item.ВидПоступления.Trim();

                    

                    foreach (DocExcelCell it in strExcelProcessDocGetPrint[strKey])
                    {
                        it.ValueCell = item.Count.ToString().Trim();
                        it.FlagEdit = true;
                    }
                }
  
                listExcelCell.Add(month,strExcelProcessDocGetPrint);
            }


            //// Строка которая содержит значения ИТОГО для отчета.
            //List<DocExcelCell> itemCount = new List<DocExcelCell>();

            Dictionary<int, Dictionary<string, List<DocExcelCell>>> lTest = listExcelCell;

            // Создадим файл.
            ExcelStatistic excel = new ExcelStatistic(listExcelCell);
            excel.Year = Year;
            excel.CreateFile(4);

            string sTest = "";

            #endregion

        }

       
    }
}
