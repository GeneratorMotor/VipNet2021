using System;
using System.Collections.Generic;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace RegKor.Classess
{
    public class ExcelGenerate
    {
        // ������ ����������� ������� � ����������.
        I����������������� _listDirectors;

        IProcessGetDoc _listProcessDoc;

        // ���������� ��� �������� ���������� �����.
        private CarlosAg.ExcelXmlWriter.Workbook book2;

        // ���������� ��� �������� ����� ������.
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

        /// <summary>
        /// ������ ����� ������.
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
        /// ������� ����� ������.
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

        public ExcelGenerate(I����������������� listDirectors, IProcessGetDoc listProcessDoc)
        {
            _listDirectors = listDirectors;
            _listProcessDoc = listProcessDoc;
        }

        public void CreateExcel(List<YearMonth> listYearMonth, int countColumn)
        {
            _listDirectors = new �����������������();
            _listProcessDoc = new ProcessGetDoc();

            List<string> listDir = _listDirectors.GetDirectors();

            List<string> listProcess = _listProcessDoc.GetDoc();

            // ���������� ����� ��� ������.
            Dictionary<int, Dictionary<string,List<DocExcelCell>>> listExcelCell = new Dictionary<int, Dictionary<string,List<DocExcelCell>>>();

            // ������ �������� ����������� ���������� ��� �������� ���������� 
            List<PersonCell> listCell = new List<PersonCell>();

            foreach (string strName in listDir)
            {
                PersonCell pCell = new PersonCell();

                pCell.Name = strName.Trim();

                foreach (string strProc in listProcess)
                {
                    pCell.�������������������.Add(strProc);
                }

                listCell.Add(pCell);
            }

            // ������� �������� ��������.
            PersonCell pCellCount = new PersonCell();

            pCellCount.Name = "�����";
            pCellCount.�������������������.Add("�������� ��������");
            pCellCount.�������������������.Add("E-mail");
            pCellCount.�������������������.Add("VipNet");
            pCellCount.�������������������.Add("����");

            listCell.Add(pCellCount);

            // ���������� ��� �������� ������ ���������� �� �� �� ��������� ����������� ����������.
            Dictionary<int, List<ItemStatisticDoc>> dictionary = new Dictionary<int, List<ItemStatisticDoc>>();

            // ������� ���������� �� �� � ���������� �������� �� �������.
            foreach (YearMonth itYearMonth in listYearMonth)
            {
                List<ItemStatisticDoc> listStatic = �����������������������������.GetStatisticDoc(itYearMonth.Year, itYearMonth.NumMonth);
                dictionary.Add(itYearMonth.NumMonth, listStatic);
            }

            // �������� �������� �� �����.
            List<ItemStatisticDoc> listStaticCount = �����������������������������.GetStatisticDocCount(this.Year, this.StartMonth, this.EndMonth);

            // ��������� ��� ������ ����� ����� ������ 13 (��� ��� 13 ������ ��������).
            dictionary.Add(13, listStaticCount);

            Dictionary<int, List<ItemStatisticDoc>> dictionaryTest = dictionary;

            //���������� ������ ��� ������.
            // ��� ����� ������� �� ������ �������� ��������� ����������.

            // ������� ������.
            int count = 1;

            // ������ �������� ������ � ��� ����������� (���� ���).
            Dictionary<string, List<DocExcelCell>> strExcelFio = new Dictionary<string, List<DocExcelCell>>();

            // ������ �������� ������� ����������� ���������� (���� ������� ����������� ����������). 
            Dictionary<string,List<DocExcelCell>> strExcelProcessDocGet = new Dictionary<string,List<DocExcelCell>>();

            #region ������ ������ ������ � ��� �����������

            // ������� ������ ������ ������ ������� ������� ��� ����������� ���������� � �������.
            foreach (PersonCell pCellFio in listCell)
            {
                // ������� ������� ����������� ������� � ����������.
                DocExcelCell dCF = new DocExcelCell();

                // ��������� ���������� �������� ��������� ����������.
                dCF.CountColumn = countColumn;
                dCF.ValueCell = pCellFio.Name.Trim();

                List<DocExcelCell> listdCF = new List<DocExcelCell>();
                listdCF.Add(dCF);

                strExcelFio.Add(pCellFio.Name.Trim(), listdCF);

                foreach(string str in pCellFio.�������������������)
                {
                    // ������ �������� �������� ���������.
                    List<DocExcelCell> listPN = new List<DocExcelCell>();

                    // ������� ������� ����������� ����������.
                    DocExcelCell cellProcessName = new DocExcelCell();
                    cellProcessName.CountColumn = 1;
                    cellProcessName.ValueCell = str;

                    listPN.Add(cellProcessName);

                    strExcelProcessDocGet.Add(pCellFio.Name.Trim()+ "_" + str.Trim(), listPN);
                }
            }

            // ������� ������ � ��������� � �����.
            listExcelCell.Add(count * 20 , strExcelFio);

            #endregion

            // ������� ������ ������ ������� �������� �������� �������� ����������� ����������.
            count++;

            listExcelCell.Add(count * 20, strExcelProcessDocGet);

            Dictionary<int, Dictionary<string, List<DocExcelCell>>> iTest = listExcelCell;

            #region ������� �������� ����� �� ���� ������.

            //// ������ �������� ������ ��� ������ �����.
            Dictionary<string, List<DocExcelCell>> strExcelProcessDocGetPrintCount = new Dictionary<string, List<DocExcelCell>>();

            DocExcelCell dcCellCount = new DocExcelCell();
            
            // ������� �������� ����� �� ���� ������.
            foreach (int month in dictionary.Keys)
            {
                List<ItemStatisticDoc> list = dictionary[month];

                // === ����
                Dictionary<string, List<DocExcelCell>> strExcelProcessDocGetPrint = new Dictionary<string, List<DocExcelCell>>();

                // ������ ����� ���� foreach ���������� � strExcelProcessDocGetPrint ������ �� strExcelProcessDocGet

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

                // ==== ����� ����

                // �������� ����� ������ ���������� �������� ����� �������� ����������� ���������� �� ��� ����������� ������� � �����.
                Dictionary<string, List<DocExcelCell>> listValue = strExcelProcessDocGet;

               
                // ��������� �� �������� �� ��.
                foreach (ItemStatisticDoc item in list)
                {
                    string strKey = item.���.Trim() + "_" + item.��������������.Trim();

                    

                    foreach (DocExcelCell it in strExcelProcessDocGetPrint[strKey])
                    {
                        it.ValueCell = item.Count.ToString().Trim();
                        it.FlagEdit = true;
                    }
                }
  
                listExcelCell.Add(month,strExcelProcessDocGetPrint);
            }


            //// ������ ������� �������� �������� ����� ��� ������.
            //List<DocExcelCell> itemCount = new List<DocExcelCell>();

            Dictionary<int, Dictionary<string, List<DocExcelCell>>> lTest = listExcelCell;

            // �������� ����.
            ExcelStatistic excel = new ExcelStatistic(listExcelCell);
            excel.Year = Year;
            excel.CreateFile(4);

            string sTest = "";

            #endregion

        }

       
    }
}
