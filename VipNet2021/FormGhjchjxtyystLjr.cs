using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;
using Word = Microsoft.Office.Interop.Word;

namespace RegKor
{
    public partial class FormПросроченныеДок : Form
    {
        private List<ДокументыИстекшийСрок> list;

        /// <summary>
        /// Просроченные документы.
        /// </summary>
        public List<ДокументыИстекшийСрок> ListDoc
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

        DataTable tabDate;

        public DataTable TabDate
        {
            get
            {
                return tabDate;
            }
            set
            {
                tabDate = value;
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

        public FormПросроченныеДок()
        {
            InitializeComponent();
        }

        private void FormПросроченныеДок_Load(object sender, EventArgs e)
        {
            //string query = "SELECT     convert(VARCHAR,dbo.Карточка.номерПП) + '/' + dbo.Карточка.НомерВход as 'НомерВход',  dbo.Карточка.ДатаПоступ, dbo.Корреспонденты.ОписаниеКорреспондента, " +
            //             " dbo.Карточка.КраткоеСодержание, dbo.Карточка.ДатаИсхода, dbo.Карточка.НомерИсход, dbo.Карточка.СрокВыполнения, " +
            //             " dbo.Карточка.Резолюция " +
            //             "FROM         dbo.Карточка INNER JOIN " +
            //             "  dbo.Корреспонденты ON dbo.Карточка.id_корреспондента = dbo.Корреспонденты.id_корреспондента " +
            //             " where ВДело = 'False' and СрокВыполнения < CONVERT(DATE,GETDATE()) ";

            //GetDataTable getTable = new GetDataTable(query);
            //DataTable tab = getTable.DataTable("Выборка");

            this.dataGridView1.DataSource = this.TabDate;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            this.TopMost = false;

            // Выведим на печать.
            try
            {
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

                // Установим горизонтальную ориентацию страницы.
                worddocument.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;

                //Получаем ссылки на параграфы документа
                wordparagraphs = worddocument.Paragraphs;
                //Будем работать с первым параграфом
                wordparagraph = (Word.Paragraph)wordparagraphs[1];
                //Выводим текст в первый параграф
                wordparagraph.Range.Text = "Документы с истекшими сроками исполнения от " + DateTime.Today.ToShortDateString();
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

                //теперь задаём форматирование следующего абзаца
                //Добавляем в документ несколько параграфов
                object oMissing = System.Reflection.Missing.Value;
                // Добавим пустой параграф.
                worddocument.Paragraphs.Add(ref oMissing);

                wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                //Переходим к первому добавленному параграфу
                wordparagraph = worddocument.Paragraphs[4];
                wordparagraph.Range.Font.Size = 9;
                Word.Range wordrange = wordparagraph.Range;

                //Добавляем таблицу в 6 параграф
                Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
                Word.Table wordtable1 = worddocument.Tables.Add(wordrange, this.dataGridView1.Rows.Count+1, 9, ref defaultTableBehavior, ref autoFitBehavior);
                wordtable1.Borders.Enable = 1;

                int iCountHead = 2;

                // Выведим информацию.
                Word.Range wordcellrange1 = worddocument.Tables[1].Cell(iCountHead, 1).Range;
                wordcellrange1.Text = "№ п.п";
                wordcellrange1.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                Word.Range wordcellrange2 = worddocument.Tables[1].Cell(iCountHead, 2).Range;

                wordcellrange2.Text = "№ входящий.";
                wordcellrange2.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange3 = worddocument.Tables[1].Cell(iCountHead, 3).Range;

                wordcellrange3.Text = "Дата поступления";
                wordcellrange3.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange4 = worddocument.Tables[1].Cell(iCountHead, 4).Range;
                wordcellrange4.Text = "Корреспондент";
                wordcellrange4.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange5 = worddocument.Tables[1].Cell(iCountHead, 5).Range;
                wordcellrange5.Text = "Краткое содержание";
                wordcellrange5.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange6 = worddocument.Tables[1].Cell(iCountHead, 6).Range;
                wordcellrange6.Text = "Дата исходящая";
                wordcellrange6.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange7 = worddocument.Tables[1].Cell(iCountHead, 7).Range;
                wordcellrange7.Text = "Номер исходящий";
                wordcellrange7.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange8 = worddocument.Tables[1].Cell(iCountHead, 8).Range;
                wordcellrange8.Text = "Срок исполнения";
                wordcellrange8.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange9 = worddocument.Tables[1].Cell(iCountHead, 9).Range;
                wordcellrange9.Text = "Исполнитель";
                wordcellrange9.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                int iCount = 3;

                foreach (DataGridViewRow row in this.dataGridView1.Rows)
                {
                    if (row.Cells["НомерВход"].Value.ToString().Length > 0)
                    {
                        // Выведим информацию.
                        Word.Range wordcellrange1t = worddocument.Tables[1].Cell(iCount, 1).Range;
                        wordcellrange1t.Text = (iCount - 2).ToString().Trim();
                        wordcellrange1t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                        Word.Range wordcellrange2t = worddocument.Tables[1].Cell(iCount, 2).Range;

                        wordcellrange2t.Text = row.Cells["НомерВход"].Value.ToString();
                        wordcellrange2t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange3t = worddocument.Tables[1].Cell(iCount, 3).Range;

                        DateTime dt = Convert.ToDateTime(row.Cells["ДатаПоступ"].Value);

                        wordcellrange3t.Text = dt.ToShortDateString().Trim();
                        wordcellrange3t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange4t = worddocument.Tables[1].Cell(iCount, 4).Range;
                        wordcellrange4t.Text = row.Cells["ОписаниеКорреспондента"].Value.ToString().Trim();
                        wordcellrange4t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange5t = worddocument.Tables[1].Cell(iCount, 5).Range;
                        wordcellrange5t.Text = row.Cells["КраткоеСодержание"].Value.ToString().Trim();
                        wordcellrange5t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange6t = worddocument.Tables[1].Cell(iCount, 6).Range;
                        wordcellrange6t.Text = Convert.ToDateTime(row.Cells["ДатаИсхода"].Value).ToShortDateString();
                        wordcellrange6t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange7t = worddocument.Tables[1].Cell(iCount, 7).Range;
                        wordcellrange7t.Text = row.Cells["НомерИсход"].Value.ToString().Trim();
                        wordcellrange7t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange8t = worddocument.Tables[1].Cell(iCount, 8).Range;
                        wordcellrange8t.Text = Convert.ToDateTime(row.Cells["СрокВыполнения"].Value).ToShortDateString().Trim();
                        wordcellrange8t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange9t = worddocument.Tables[1].Cell(iCount, 9).Range;
                        wordcellrange9t.Text = row.Cells["Резолюция"].Value.ToString().Trim();
                        wordcellrange9t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    }

                        iCount++;
                      
                }
            }
            catch (Exception ex)
            {
                Text = ex.Message;
            }

            // Объеденим первы 2 ячейки вертикально.
            object begCell = worddocument.Tables[1].Cell(1, 1).Range.Start;
            object endCell = worddocument.Tables[1].Cell(2, 1).Range.End;
            Word.Range wordcellrange = worddocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            wordapp.Selection.Cells.Merge();

            // Объеденим горизонтаьно первую строку.
            object begCell2 = worddocument.Tables[1].Cell(1, 2).Range.Start;
            object endCell2 = worddocument.Tables[1].Cell(1, 9).Range.End;
            wordcellrange = worddocument.Range(ref begCell2, ref endCell2);
            wordcellrange.Select();
            wordapp.Selection.Cells.Merge();

            worddocument.Tables[1].Cell(1, 2).Range.Text = "Документы";

        }

        


    }
}