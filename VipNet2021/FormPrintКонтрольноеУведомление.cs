using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using RegKor.Classess;

namespace RegKor
{
    public partial class FormPrintКонтрольноеУведомление : Form
    {
        private DataSet ds;

        private StatisticControlNotific statistic;

        /// <summary>
        /// Свойство для хранения данных для отчета.
        /// </summary>
        public StatisticControlNotific DataStatistic
        {
            get
            {
                return statistic;
            }
            set
            {
                statistic = value;
            }
        }

        /// <summary>
        /// Получает DataSet.
        /// </summary>
        public DataSet DataSetForm
        {
            get
            {
                return ds;
            }
            set
            {
                ds = value;
            }
        }

        /// <summary>
        /// Количество документов на контроле.
        /// </summary>
        private int countDocControl = 0;

        /// <summary>
        /// Количество документов на контроле.
        /// </summary>
        public int CountDocControl
        {
            get
            {
                return countDocControl;
            }
            set
            {
                countDocControl = value;
            }
        }

        // Переменная документа Word.
        private Word.Application wordapp;

        // Параграф.
        private Word.Paragraphs wordparagraphs;
        private Word.Paragraph wordparagraph;

        // Word документы.
        private Word.Documents worddocuments;
        private Word.Document worddocument;


        public FormPrintКонтрольноеУведомление()
        {
            InitializeComponent();

            StatisticControlNotific statistic = new StatisticControlNotific();
        }

        private bool flagLoad = false;

        private void FormPrintКонтрольноеУведомление_Load(object sender, EventArgs e)
        {
            //DataTable tab = this.ds.Tables["Получатели"];
            //this.comboBox1.DataSource = this.ds.Tables["Получатели"];
            //this.comboBox1.DisplayMember = "ОписаниеПолучателя";
            //this.comboBox1.ValueMember = "id_Получателя";

            //flagLoad = true;

            //int id = (int)this.comboBox1.SelectedValue;

            //// Отобразим 
            //this.dataGridView1.DataSource = this.ds.Tables["Документы"].Select("id_Получателя =" + id );
            //DisplayDataGrid();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (flagLoad == true)
            //{
            //    int id = (int)this.comboBox1.SelectedValue;

            //    // Отобразим 
            //    this.dataGridView1.DataSource = this.ds.Tables["Документы"].Select("id_Получателя =" + id);
            //    DisplayDataGrid();
            //}
        }

        private void DisplayDataGrid()
        {
        //    this.dataGridView1.Columns["id_Документа"].Visible = false;
        //    this.dataGridView1.Columns["id_Получателя"].Visible = false;
        //    this.dataGridView1.Columns["ТипДокумента"].Visible = false;
        //    this.dataGridView1.Columns["RowError"].Visible = false;
        //    this.dataGridView1.Columns["RowState"].Visible = false;
        //    this.dataGridView1.Columns["Table"].Visible = false;
        //    this.dataGridView1.Columns["HasErrors"].Visible = false;
        //    this.dataGridView1.Columns["НомерПП"].DisplayIndex = 1;
        //    this.dataGridView1.Columns["ДатаПоступления"].DisplayIndex = 2;
        //    this.dataGridView1.Columns["ДатаПоступления"].HeaderText = "Дата поступления";
        //    this.dataGridView1.Columns["НомерВходящий"].DisplayIndex = 3;
        //    this.dataGridView1.Columns["НомерВходящий"].HeaderText = "Номер входящий";
        //    this.dataGridView1.Columns["ДатаКонтроля"].DisplayIndex = 4;
        //    this.dataGridView1.Columns["ДатаКонтроля"].HeaderText = "Срок контроля";

        //    this.label1.Text = "Всего документов на контроле : " + this.dataGridView1.Rows.Count.ToString().Trim();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            //try
            //{
                //Создаем объект Word - равносильно запуску Word
                wordapp = new Word.Application();
                //Делаем его видимым
                wordapp.Visible = true;

                Object template = Type.Missing;
                Object newTemplate = false;
                Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;
                ////Создаем документ 1
                //wordapp.Documents.Add(
                //ref template, ref newTemplate, ref documentType, ref visible);

                //Создаем документ 2 worddocument в данном случае создаваемый объект 
                worddocument =
                wordapp.Documents.Add(
                 ref template, ref newTemplate, ref documentType, ref visible);

                //Получаем ссылки на параграфы документа
                wordparagraphs = worddocument.Paragraphs;
                //Будем работать с первым параграфом
                wordparagraph = (Word.Paragraph)wordparagraphs[1];
                //Выводим текст в первый параграф
                wordparagraph.Range.Text = "КОНТРОЛЬНОЕ УВЕДОМЛЕНИЕ от " + DateTime.Today.ToShortDateString();
                //Меняем характеристики текста и параграфа
                //wordparagraph.Range.Font.Color = Word.WdColor.wdColorBlue;
                wordparagraph.Range.Font.Size = 14;
                wordparagraph.Range.Font.Name = "Times New Roman";
                //wordparagraph.Range.Font.Italic = 1;
                wordparagraph.Range.Font.Bold = 1;
                //wordparagraph.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                //wordparagraph.Range.Font.UnderlineColor = Word.WdColor.wdColorDarkRed;
                //wordparagraph.Range.Font.StrikeThrough=1; можно перечеркнуть
                //Выравнивание
                wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                // Добавим пустой параграф.
                object oMissingStart3 = System.Reflection.Missing.Value;
                worddocument.Paragraphs.Add(ref oMissingStart3);

                // Добавим пустой параграф.
                object oMissing = System.Reflection.Missing.Value;
                // Добавим пустой параграф.
                worddocument.Paragraphs.Add(ref oMissing);
                wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                wordparagraph.Range.Font.Size = 10;
                wordparagraph.Range.Text = "Всего документов на контроле : " + this.DataStatistic.ВсегоДокументыНаКонтроле.ToString().Trim(); //this.CountDocControl.ToString().Trim();
                wordparagraph.Range.Font.Bold = 1;
 
                // Добавим пустой параграф.
                object oMissingStart21 = System.Reflection.Missing.Value;
                worddocument.Paragraphs.Add(ref oMissingStart21);

                object oMissingStart2 = System.Reflection.Missing.Value;
                // Добавим пустой параграф.
                worddocument.Paragraphs.Add(ref oMissingStart2);
                wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                wordparagraph.Range.Font.Size = 10;
                wordparagraph.Range.Text = "В том числе с истекшим сроком : " + this.DataStatistic.КоличествоПросроченныхДокументов.ToString().Trim();//this.CountDocControl.ToString().Trim();
                wordparagraph.Range.Font.Bold = 1;

                // Добавим пустой параграф.
                object oMissingStart22 = System.Reflection.Missing.Value;
                worddocument.Paragraphs.Add(ref oMissingStart22);

                int iTable = 0;

                foreach (PersonDocument person in this.DataStatistic.СписокИсполнителей)
                {
                    // Добавим пустой параграф.
                    object oMissingStartA1 = System.Reflection.Missing.Value;
                    worddocument.Paragraphs.Add(ref oMissingStartA1);

                    // Добавим пустой параграф.
                    object oMissingStartA = System.Reflection.Missing.Value;
                    worddocument.Paragraphs.Add(ref oMissingStartA);

                    wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordparagraph.Range.Font.Size = 10;
                    wordparagraph.Range.Text = "Получатель : " + person.FioPerson;
                    wordparagraph.Range.Font.Bold = 1;

                    // Добавми параграф.
                    object oMissingStartB = System.Reflection.Missing.Value;
                    worddocument.Paragraphs.Add(ref oMissingStartB);

                    wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordparagraph.Range.Font.Size = 10;
                    wordparagraph.Range.Text = "Количество документов на контроле : " + person.ВсегоДокументыНаКонтроле;
                    wordparagraph.Range.Font.Bold = 0;

                    //// Печатаем документы которые на контроле.
                    //if (person.ВсегоДокументыНаКонтроле > 0)
                    //{
                    //    // Добавим параграф.
                    //    object oMissingStartD1 = System.Reflection.Missing.Value;
                    //    worddocument.Paragraphs.Add(ref oMissingStartD1);

                    //    // Добавим таблицу.
                    //    int iPar = worddocument.Paragraphs.Count;

                    //    wordparagraph = worddocument.Paragraphs[iPar];
                    //    Word.Range wordrange = wordparagraph.Range;

                    //    //Добавляем таблицу содержащую документы с истёкшим сроком в начало параграфа
                    //    Object defaultTableBehavior =
                    //    Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    //    Object autoFitBehavior =
                    //    Word.WdAutoFitBehavior.wdAutoFitWindow;

                    //    Word.Table wordtable1 = worddocument.Tables.Add(wordrange, person.ВсегоДокументыНаКонтроле * 2, 4,
                    //    ref defaultTableBehavior, ref autoFitBehavior);

                    //    Object style = "Классическая таблица 1";
                    //    wordtable1.set_Style(ref style);
                    //    //Далее можно добавлять выделение первых
                    //    //и последних строк и столбцов
                    //    wordtable1.ApplyStyleFirstColumn = false;
                    //    wordtable1.ApplyStyleHeadingRows = false;
                    //    wordtable1.ApplyStyleLastRow = false;
                    //    wordtable1.ApplyStyleLastColumn = false;

                    //    // счетчик таблиц.
                    //    iTable++;

                    //    int iCount = 1;
                    //    int iCountNum = 1;

                    //     //==== Здесь описываем документы чере foreach
                    //    foreach (DataRow r in person.ДокументыНаКонтроле)
                    //    {
                    //        // Выведим информацию.
                    //        Word.Range wordcellrange1 = worddocument.Tables[iTable].Cell(iCount, 1).Range;
                    //        wordcellrange1.Text = iCountNum.ToString().Trim();
                    //        wordcellrange1.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                    //        Word.Range wordcellrange2 = worddocument.Tables[iTable].Cell(iCount, 2).Range;
                    //        wordcellrange2.Text = Convert.ToDateTime(r["ДатаПоступ"]).ToShortDateString();
                    //        wordcellrange2.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                    //        Word.Range wordcellrange3 = worddocument.Tables[iTable].Cell(iCount, 3).Range;
                    //        wordcellrange3.Text = r["НомерВход"].ToString().Trim();
                    //        wordcellrange3.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                    //        Word.Range wordcellrange4 = worddocument.Tables[iTable].Cell(iCount, 4).Range;
                    //        wordcellrange4.Text = Convert.ToDateTime(r["СрокВыполнения"]).ToShortDateString();
                    //        wordcellrange4.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                    //        iCount++;

                    //        // Добавим подписи внизу.
                    //        Word.Range wordcellrange12 = worddocument.Tables[iTable].Cell(iCount, 1).Range;
                    //        wordcellrange12.Text = "";
                    //        wordcellrange12.Font.Size = 8;


                    //        Word.Range wordcellrange22 = worddocument.Tables[iTable].Cell(iCount, 2).Range;
                    //        wordcellrange22.Text = "дата поступления";
                    //        wordcellrange22.Font.Size = 8;

                    //        Word.Range wordcellrange32 = worddocument.Tables[iTable].Cell(iCount, 3).Range;
                    //        wordcellrange32.Text = "номер входящий";
                    //        wordcellrange32.Font.Size = 8;

                    //        Word.Range wordcellrange42 = worddocument.Tables[iTable].Cell(iCount, 4).Range;
                    //        wordcellrange42.Text = "срок контроля";
                    //        wordcellrange42.Font.Size = 8;

                    //        iCount++;
                    //        iCountNum++;
                    //    }
                    //}

                    // Добавим параграф.
                    object oMissingStartC = System.Reflection.Missing.Value;
                    worddocument.Paragraphs.Add(ref oMissingStartC);

                    wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordparagraph.Range.Font.Size = 10;
                    wordparagraph.Range.Text = "с истекшим сроком : " + person.КоличествоПросроченныхДокументов;
                    wordparagraph.Range.Font.Bold = 0;

                    if (person.ПросроченныеДокументы.Length > 0)
                    {
                        // Добавим параграф.
                        object oMissingStartD = System.Reflection.Missing.Value;
                        worddocument.Paragraphs.Add(ref oMissingStartD);

                        // Добавим таблицу.
                        int iPar = worddocument.Paragraphs.Count;

                        wordparagraph = worddocument.Paragraphs[iPar];
                        Word.Range wordrange = wordparagraph.Range;

                        //Добавляем таблицу содержащую документы с истёкшим сроком в начало параграфа
                        Object defaultTableBehavior =
                        Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                        Object autoFitBehavior =
                        Word.WdAutoFitBehavior.wdAutoFitWindow;

                        Word.Table wordtable1 = worddocument.Tables.Add(wordrange, person.ПросроченныеДокументы.Length * 2, 4,
                        ref defaultTableBehavior, ref autoFitBehavior);

                        Object style = "Классическая таблица 1";
                        wordtable1.set_Style(ref style);
                        //Далее можно добавлять выделение первых
                        //и последних строк и столбцов
                        wordtable1.ApplyStyleFirstColumn = false;
                        wordtable1.ApplyStyleHeadingRows = false;
                        wordtable1.ApplyStyleLastRow = false;
                        wordtable1.ApplyStyleLastColumn = false;

                        // счетчик таблиц.
                        iTable++;

                        int iCount = 1;
                        int iCountNum = 1;

                        //==== Здесь описываем документы чере foreach
                        foreach (DataRow r in person.ПросроченныеДокументы)
                        {
                            // Выведим информацию.
                            Word.Range wordcellrange1 = worddocument.Tables[iTable].Cell(iCount, 1).Range;
                            wordcellrange1.Text = iCountNum.ToString().Trim();
                            wordcellrange1.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                            Word.Range wordcellrange2 = worddocument.Tables[iTable].Cell(iCount, 2).Range;
                            wordcellrange2.Text = Convert.ToDateTime(r["ДатаПоступ"]).ToShortDateString();
                            wordcellrange2.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                            Word.Range wordcellrange3 = worddocument.Tables[iTable].Cell(iCount, 3).Range;
                            wordcellrange3.Text = r["НомерВход"].ToString().Trim();
                            wordcellrange3.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                            Word.Range wordcellrange4 = worddocument.Tables[iTable].Cell(iCount, 4).Range;
                            wordcellrange4.Text = Convert.ToDateTime(r["СрокВыполнения"]).ToShortDateString();
                            wordcellrange4.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                            iCount++;

                            // Добавим подписи внизу.
                            Word.Range wordcellrange12 = worddocument.Tables[iTable].Cell(iCount, 1).Range;
                            wordcellrange12.Text = "";
                            wordcellrange12.Font.Size = 8;


                            Word.Range wordcellrange22 = worddocument.Tables[iTable].Cell(iCount, 2).Range;
                            wordcellrange22.Text = "дата поступления";
                            wordcellrange22.Font.Size = 8;

                            Word.Range wordcellrange32 = worddocument.Tables[iTable].Cell(iCount, 3).Range;
                            wordcellrange32.Text = "номер входящий";
                            wordcellrange32.Font.Size = 8;

                            Word.Range wordcellrange42 = worddocument.Tables[iTable].Cell(iCount, 4).Range;
                            wordcellrange42.Text = "срок контроля";
                            wordcellrange42.Font.Size = 8;

                            iCount++;
                            iCountNum++;
                        }

                        //// Добавим пустой параграф.
                        //object oMissingStartE = System.Reflection.Missing.Value;
                        //worddocument.Paragraphs.Add(ref oMissingStartE);

                        // Добавим пустой параграф.
                        object oMissingStartF = System.Reflection.Missing.Value;
                        worddocument.Paragraphs.Add(ref oMissingStartF);

                        wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        wordparagraph.Range.Font.Size = 10;
                        wordparagraph.Range.Text = "с неистекшим сроком : " + person.КоличествоНеПросроченныхДокументов.ToString().Trim();
                        wordparagraph.Range.Font.Bold = 0;

                        // Добавим пустой параграф.
                        object oMissingStartF1 = System.Reflection.Missing.Value;
                        worddocument.Paragraphs.Add(ref oMissingStartF1);


                        // Добавим пустой параграф.
                        object oMissingStartG = System.Reflection.Missing.Value;
                        worddocument.Paragraphs.Add(ref oMissingStartG);

                        if (person.НеПрсороченныеДокументы.Length > 0)
                        {
                            int iPar2 = worddocument.Paragraphs.Count;

                            wordparagraph = worddocument.Paragraphs[iPar2];
                            Word.Range wordrange2 = wordparagraph.Range;

                            //Добавляем таблицу содержащую документы с истёкшим сроком в начало параграфа
                            Object defaultTableBehavior2 =
                            Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                            Object autoFitBehavior2 =
                            Word.WdAutoFitBehavior.wdAutoFitWindow;

                            Word.Table wordtable2 = worddocument.Tables.Add(wordrange2, person.НеПрсороченныеДокументы.Length * 2, 4,
                            ref defaultTableBehavior, ref autoFitBehavior);

                            Object style2 = "Классическая таблица 1";
                            wordtable2.set_Style(ref style);
                            //Далее можно добавлять выделение первых
                            //и последних строк и столбцов
                            wordtable2.ApplyStyleFirstColumn = false;
                            wordtable2.ApplyStyleHeadingRows = false;
                            wordtable2.ApplyStyleLastRow = false;
                            wordtable2.ApplyStyleLastColumn = false;

                            // счетчик таблиц.
                            iTable++;

                            int iCount2 = 1;
                            int iCountNum2 = 1;

                            foreach (DataRow r in person.НеПрсороченныеДокументы)
                            {
                                // Выведим информацию.
                                Word.Range wordcellrange1 = worddocument.Tables[iTable].Cell(iCount2, 1).Range;
                                wordcellrange1.Text = iCountNum2.ToString().Trim();
                                wordcellrange1.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                                Word.Range wordcellrange2 = worddocument.Tables[iTable].Cell(iCount2, 2).Range;
                                wordcellrange2.Text = Convert.ToDateTime(r["ДатаПоступ"]).ToShortDateString();
                                wordcellrange2.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                                Word.Range wordcellrange3 = worddocument.Tables[iTable].Cell(iCount2, 3).Range;
                                wordcellrange3.Text = r["НомерВход"].ToString().Trim();
                                wordcellrange3.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                                Word.Range wordcellrange4 = worddocument.Tables[iTable].Cell(iCount2, 4).Range;
                                wordcellrange4.Text = Convert.ToDateTime(r["СрокВыполнения"]).ToShortDateString();
                                wordcellrange4.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                                iCount2++;

                                // Добавим подписи внизу.
                                Word.Range wordcellrange12 = worddocument.Tables[iTable].Cell(iCount2, 1).Range;
                                wordcellrange12.Text = "";
                                wordcellrange12.Font.Size = 8;


                                Word.Range wordcellrange22 = worddocument.Tables[iTable].Cell(iCount2, 2).Range;
                                wordcellrange22.Text = "дата поступления";
                                wordcellrange22.Font.Size = 8;

                                Word.Range wordcellrange32 = worddocument.Tables[iTable].Cell(iCount2, 3).Range;
                                wordcellrange32.Text = "номер входящий";
                                wordcellrange32.Font.Size = 8;

                                Word.Range wordcellrange42 = worddocument.Tables[iTable].Cell(iCount2, 4).Range;
                                wordcellrange42.Text = "срок контроля";
                                wordcellrange42.Font.Size = 8;

                                iCount2++;
                                iCountNum2++;
                            }
                        }
                    }
                    else
                    {
                        // Добавим пустой параграф.
                        object oMissingStartF = System.Reflection.Missing.Value;
                        worddocument.Paragraphs.Add(ref oMissingStartF);

                        wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                        wordparagraph.Range.Font.Size = 10;
                        wordparagraph.Range.Text = "с неистекшим сроком : " + person.КоличествоНеПросроченныхДокументов.ToString().Trim();
                        wordparagraph.Range.Font.Bold = 0;

                        // Добавим пустой параграф.
                        object oMissingStartF1 = System.Reflection.Missing.Value;
                        worddocument.Paragraphs.Add(ref oMissingStartF1);


                        // Добавим пустой параграф.
                        object oMissingStartG = System.Reflection.Missing.Value;
                        worddocument.Paragraphs.Add(ref oMissingStartG);

                        if (person.НеПрсороченныеДокументы.Length > 0)
                        {
                            int iPar2 = worddocument.Paragraphs.Count;

                            wordparagraph = worddocument.Paragraphs[iPar2];
                            Word.Range wordrange2 = wordparagraph.Range;

                            //Добавляем таблицу содержащую документы с истёкшим сроком в начало параграфа
                            Object defaultTableBehavior2 =
                            Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                            Object autoFitBehavior2 =
                            Word.WdAutoFitBehavior.wdAutoFitWindow;

                            Word.Table wordtable2 = worddocument.Tables.Add(wordrange2, person.НеПрсороченныеДокументы.Length * 2, 4,
                            ref defaultTableBehavior2, ref autoFitBehavior2);

                            Object style2 = "Классическая таблица 1";
                            wordtable2.set_Style(ref style2);
                            //Далее можно добавлять выделение первых
                            //и последних строк и столбцов
                            wordtable2.ApplyStyleFirstColumn = false;
                            wordtable2.ApplyStyleHeadingRows = false;
                            wordtable2.ApplyStyleLastRow = false;
                            wordtable2.ApplyStyleLastColumn = false;

                            // счетчик таблиц.
                            iTable++;

                            int iCount2 = 1;
                            int iCountNum2 = 1;

                            foreach (DataRow r in person.НеПрсороченныеДокументы)
                            {
                                // Выведим информацию.
                                Word.Range wordcellrange1 = worddocument.Tables[iTable].Cell(iCount2, 1).Range;
                                wordcellrange1.Text = iCountNum2.ToString().Trim();
                                wordcellrange1.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                                Word.Range wordcellrange2 = worddocument.Tables[iTable].Cell(iCount2, 2).Range;
                                wordcellrange2.Text = Convert.ToDateTime(r["ДатаПоступ"]).ToShortDateString();
                                wordcellrange2.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                                Word.Range wordcellrange3 = worddocument.Tables[iTable].Cell(iCount2, 3).Range;
                                wordcellrange3.Text = r["НомерВход"].ToString().Trim();
                                wordcellrange3.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                                Word.Range wordcellrange4 = worddocument.Tables[iTable].Cell(iCount2, 4).Range;
                                wordcellrange4.Text = Convert.ToDateTime(r["СрокВыполнения"]).ToShortDateString();
                                wordcellrange4.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                                iCount2++;

                                // Добавим подписи внизу.
                                Word.Range wordcellrange12 = worddocument.Tables[iTable].Cell(iCount2, 1).Range;
                                wordcellrange12.Text = "";
                                wordcellrange12.Font.Size = 8;


                                Word.Range wordcellrange22 = worddocument.Tables[iTable].Cell(iCount2, 2).Range;
                                wordcellrange22.Text = "дата поступления";
                                wordcellrange22.Font.Size = 8;

                                Word.Range wordcellrange32 = worddocument.Tables[iTable].Cell(iCount2, 3).Range;
                                wordcellrange32.Text = "номер входящий";
                                wordcellrange32.Font.Size = 8;

                                Word.Range wordcellrange42 = worddocument.Tables[iTable].Cell(iCount2, 4).Range;
                                wordcellrange42.Text = "срок контроля";
                                wordcellrange42.Font.Size = 8;

                                iCount2++;
                                iCountNum2++;
                            }
                        }
                    }
                }

                    


                
                /*

                // Счетчик параграфов.
                int par = 3;

                DataSet dsTest = this.ds;

                DataTable tab = this.ds.Tables["Получатели"];

                int iTable = 0;

                foreach (DataRow row in this.ds.Tables["Получатели"].Rows)
                {
                    DataRow[] rows = this.ds.Tables["Документы"].Select("id_Получателя =" + Convert.ToInt16(row["id_Получателя"]));
                    if (rows.Length == 0)
                    {
                        continue;
                    }

                    //теперь задаём форматирование следующего абзаца
                    //Добавляем в документ несколько параграфов
                    object oMissingStart = System.Reflection.Missing.Value;
                    // Добавим пустой параграф.
                    worddocument.Paragraphs.Add(ref oMissingStart);

                    // Добавим параграф.
                    wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph.Range.Font.Size = 10;
                    wordparagraph.Range.Text = "Получатель :" + row["ОписаниеПолучателя"].ToString().Trim();
                    wordparagraph.Range.Font.Bold = 0;
                    wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                    par++;
                    object oMissing2 = System.Reflection.Missing.Value;
                    // Добавим пустой параграф.
                    worddocument.Paragraphs.Add(ref oMissing2);
                    // Добьавим параграф.
                    wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordparagraph.Range.Font.Size = 10;
                    wordparagraph.Range.Text = "Всего документов на контроле : " + row["КолвоДокументовНаКонтроле"].ToString().Trim();
                    wordparagraph.Range.Font.Bold = 0;

                    par++;
                    object oMissing3 = System.Reflection.Missing.Value;
                    // Добавим пустой параграф.
                    worddocument.Paragraphs.Add(ref oMissing3);
                    // Добавим параграф.
                    wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                    wordparagraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    wordparagraph.Range.Font.Size = 10;
                    wordparagraph.Range.ParagraphFormat.LeftIndent = 5;

                    wordparagraph.Range.Text = "В том числе документы с истекшим сроком исполнения :" + row["КолвоПросроченныхДокументов"].ToString().Trim();
                    wordparagraph.Range.Font.Bold = 0;

                    int asd = par++;

                    object oMissing4 = System.Reflection.Missing.Value;

                   // Добавим пустой параграф.
                    worddocument.Paragraphs.Add(ref oMissing4);
                    wordparagraph = worddocument.Paragraphs.Add(ref oMissing4);

                    int iPar = worddocument.Paragraphs.Count;

                    wordparagraph = worddocument.Paragraphs[iPar];
                    Word.Range wordrange = wordparagraph.Range;
                    //Добавляем таблицу в начало второго параграфа
                    Object defaultTableBehavior =
                    Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                    Object autoFitBehavior =
                    Word.WdAutoFitBehavior.wdAutoFitWindow;

                    Word.Table wordtable1 = worddocument.Tables.Add(wordrange, rows.Length * 2, 4,
                    ref defaultTableBehavior, ref autoFitBehavior);

                    Object style = "Классическая таблица 1";
                    wordtable1.set_Style(ref style);
                    //Далее можно добавлять выделение первых
                    //и последних строк и столбцов
                    wordtable1.ApplyStyleFirstColumn = false;
                    wordtable1.ApplyStyleHeadingRows = false;
                    wordtable1.ApplyStyleLastRow = false;
                    wordtable1.ApplyStyleLastColumn = false;

                    // счетчик таблиц.
                    iTable++;

                    int iCount = 1;

                    //==== Здесь описываем документы чере foreach
                        foreach (DataRow r in rows)
                        {
                        //if (row.Cells["НомерПП"].Value.ToString() != "")
                        //{
                            // Выведим информацию.
                            Word.Range wordcellrange1 = worddocument.Tables[iTable].Cell(iCount, 1).Range;
                            wordcellrange1.Text = r["НомерПП"].ToString().Trim();
                            wordcellrange1.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                            Word.Range wordcellrange2 = worddocument.Tables[iTable].Cell(iCount, 2).Range;
                            wordcellrange2.Text = Convert.ToDateTime(r["ДатаПоступления"]).ToShortDateString();
                            wordcellrange2.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                            Word.Range wordcellrange3 = worddocument.Tables[iTable].Cell(iCount, 3).Range;
                            wordcellrange3.Text = r["НомерВходящий"].ToString().Trim();
                            wordcellrange3.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                            Word.Range wordcellrange4 = worddocument.Tables[iTable].Cell(iCount, 4).Range;
                            wordcellrange4.Text = Convert.ToDateTime(r["ДатаКонтроля"]).ToShortDateString();
                            wordcellrange4.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                            iCount++;

                            // Добавим подписи внизу.
                            Word.Range wordcellrange12 = worddocument.Tables[iTable].Cell(iCount, 1).Range;
                            wordcellrange12.Text = "";
                            wordcellrange12.Font.Size = 8;


                            Word.Range wordcellrange22 = worddocument.Tables[iTable].Cell(iCount, 2).Range;
                            wordcellrange22.Text = "дата поступления";
                            wordcellrange22.Font.Size = 8;

                            Word.Range wordcellrange32 = worddocument.Tables[iTable].Cell(iCount, 3).Range;
                            wordcellrange32.Text = "номер входящий";
                            wordcellrange32.Font.Size = 8;

                            Word.Range wordcellrange42 = worddocument.Tables[iTable].Cell(iCount, 4).Range;
                            wordcellrange42.Text = "срок контроля";
                            wordcellrange42.Font.Size = 8;

                            iCount++;
                        }
                    }

                    par++;
                 */

                //DataTable tab = this.ds.Tables["Получатели"];
                //this.comboBox1.DataSource = this.ds.Tables["Получатели"];
                //this.comboBox1.DisplayMember = "ОписаниеПолучателя";
                //this.comboBox1.ValueMember = "id_Получателя";

                ////теперь задаём форматирование следующего абзаца
                ////Добавляем в документ несколько параграфов
                //object oMissing = System.Reflection.Missing.Value;
                //// Добавим пустой параграф.
                //worddocument.Paragraphs.Add(ref oMissing);

                //// Добавим параграф.
                //wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                //worddocument.Paragraphs[3].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                //worddocument.Paragraphs[3].Range.Font.Size = 10;
                //worddocument.Paragraphs[3].Range.Text = "Получатель :" + this.comboBox1.Text.Trim();
                //worddocument.Paragraphs[3].Range.Font.Bold = 0;

                //wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                //worddocument.Paragraphs[4].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                //worddocument.Paragraphs[4].Range.Font.Size = 10;
                //worddocument.Paragraphs[4].Range.Text =  this.label1.Text;
                //worddocument.Paragraphs[4].Range.Font.Bold = 0;

                //wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                //worddocument.Paragraphs[5].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                //worddocument.Paragraphs[5].Range.Font.Size = 10;
                //worddocument.Paragraphs[5].Range.ParagraphFormat.LeftIndent = 5;

                //worddocument.Paragraphs[5].Range.Text = "В том числе документы с истекшим сроком исполнения :" + this.dataGridView1.Rows.Count.ToString().Trim();
                //worddocument.Paragraphs[5].Range.Font.Bold = 0;

                //wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                //wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                ////Переходим к первому добавленному параграфу
                //wordparagraph = worddocument.Paragraphs[7];
                //Word.Range wordrange = wordparagraph.Range;
                ////Добавляем таблицу в 6 параграф
                //Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                //Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
                //Word.Table wordtable1 = worddocument.Tables.Add(wordrange, this.dataGridView1.Rows.Count *2, 4, ref defaultTableBehavior, ref autoFitBehavior);
                //wordtable1.Borders.Enable = 0;

                //int iCount = 1;

                //foreach (DataGridViewRow row in this.dataGridView1.Rows)
                //{
                //    if (row.Cells["НомерПП"].Value.ToString() != "")
                //    {
                //        // Выведим информацию.
                //        Word.Range wordcellrange1 = worddocument.Tables[1].Cell(iCount, 1).Range;
                //        wordcellrange1.Text = row.Cells["НомерПП"].Value.ToString().Trim();
                //        wordcellrange1.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                //        Word.Range wordcellrange2 = worddocument.Tables[1].Cell(iCount, 2).Range;
                //        wordcellrange2.Text = Convert.ToDateTime(row.Cells["ДатаПоступления"].Value).ToShortDateString();
                //        wordcellrange2.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                //        Word.Range wordcellrange3 = worddocument.Tables[1].Cell(iCount, 3).Range;
                //        wordcellrange3.Text = row.Cells["НомерВходящий"].Value.ToString().Trim();
                //        wordcellrange3.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                //        Word.Range wordcellrange4 = worddocument.Tables[1].Cell(iCount, 4).Range;
                //        wordcellrange4.Text = Convert.ToDateTime(row.Cells["ДатаКонтроля"].Value).ToShortDateString();
                //        wordcellrange4.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                //        iCount++;

                //        // Добавим подписи внизу.
                //        Word.Range wordcellrange12 = worddocument.Tables[1].Cell(iCount, 1).Range;
                //        wordcellrange12.Text = "";
                //        wordcellrange12.Font.Size = 8;


                //        Word.Range wordcellrange22 = worddocument.Tables[1].Cell(iCount, 2).Range;
                //        wordcellrange22.Text = "дата поступления";
                //        wordcellrange22.Font.Size = 8;

                //        Word.Range wordcellrange32 = worddocument.Tables[1].Cell(iCount, 3).Range;
                //        wordcellrange32.Text = "номер входящий";
                //        wordcellrange32.Font.Size = 8;

                //        Word.Range wordcellrange42 = worddocument.Tables[1].Cell(iCount, 4).Range;
                //        wordcellrange42.Text = "срок контроля";
                //        wordcellrange42.Font.Size = 8; 

                //        iCount++;
                //    }
                //}
            //}
            //catch (Exception ex)
            //{
            //    Text = ex.Message;
            //}
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}