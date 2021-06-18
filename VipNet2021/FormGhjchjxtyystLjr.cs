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
    public partial class Form��������������� : Form
    {
        private List<���������������������> list;

        /// <summary>
        /// ������������ ���������.
        /// </summary>
        public List<���������������������> ListDoc
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

        // ���������� ��������� Word.
        private Word.Application wordapp;

        // ��������.
        private Word.Paragraphs wordparagraphs;
        private Word.Paragraph wordparagraph;

        // Word ���������.
        private Word.Documents worddocuments;
        private Word.Document worddocument;

        public Form���������������()
        {
            InitializeComponent();
        }

        private void Form���������������_Load(object sender, EventArgs e)
        {
            //string query = "SELECT     convert(VARCHAR,dbo.��������.�������) + '/' + dbo.��������.��������� as '���������',  dbo.��������.����������, dbo.��������������.����������������������, " +
            //             " dbo.��������.�����������������, dbo.��������.����������, dbo.��������.����������, dbo.��������.��������������, " +
            //             " dbo.��������.��������� " +
            //             "FROM         dbo.�������� INNER JOIN " +
            //             "  dbo.�������������� ON dbo.��������.id_�������������� = dbo.��������������.id_�������������� " +
            //             " where ����� = 'False' and �������������� < CONVERT(DATE,GETDATE()) ";

            //GetDataTable getTable = new GetDataTable(query);
            //DataTable tab = getTable.DataTable("�������");

            this.dataGridView1.DataSource = this.TabDate;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            this.TopMost = false;

            // ������� �� ������.
            try
            {
                //������� ������ Word - ����������� ������� Word
                wordapp = new Word.Application();
                //������ ��� �������
                wordapp.Visible = true;

                Object template = Type.Missing;
                Object newTemplate = false;
                Object documentType = Word.WdNewDocumentType.wdNewBlankDocument;
                Object visible = true;
                ////������� �������� 1
                //wordapp.Documents.Add(
                //ref template, ref newTemplate, ref documentType, ref visible);

                //������� �������� 2 worddocument � ������ ������ ����������� ������ 
                worddocument =
                wordapp.Documents.Add(
                 ref template, ref newTemplate, ref documentType, ref visible);

                // ��������� �������������� ���������� ��������.
                worddocument.PageSetup.Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape;

                //�������� ������ �� ��������� ���������
                wordparagraphs = worddocument.Paragraphs;
                //����� �������� � ������ ����������
                wordparagraph = (Word.Paragraph)wordparagraphs[1];
                //������� ����� � ������ ��������
                wordparagraph.Range.Text = "��������� � ��������� ������� ���������� �� " + DateTime.Today.ToShortDateString();
                //������ �������������� ������ � ���������
                //wordparagraph.Range.Font.Color = Word.WdColor.wdColorBlue;
                wordparagraph.Range.Font.Size = 14;
                wordparagraph.Range.Font.Name = "Times New Roman";
                //wordparagraph.Range.Font.Italic = 1;
                wordparagraph.Range.Font.Bold = 1;
                //wordparagraph.Range.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                //wordparagraph.Range.Font.UnderlineColor = Word.WdColor.wdColorDarkRed;
                //wordparagraph.Range.Font.StrikeThrough=1; ����� ������������
                //������������
                wordparagraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                //������ ����� �������������� ���������� ������
                //��������� � �������� ��������� ����������
                object oMissing = System.Reflection.Missing.Value;
                // ������� ������ ��������.
                worddocument.Paragraphs.Add(ref oMissing);

                wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                wordparagraph = worddocument.Paragraphs.Add(ref oMissing);
                //��������� � ������� ������������ ���������
                wordparagraph = worddocument.Paragraphs[4];
                wordparagraph.Range.Font.Size = 9;
                Word.Range wordrange = wordparagraph.Range;

                //��������� ������� � 6 ��������
                Object defaultTableBehavior = Word.WdDefaultTableBehavior.wdWord9TableBehavior;
                Object autoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitWindow;
                Word.Table wordtable1 = worddocument.Tables.Add(wordrange, this.dataGridView1.Rows.Count+1, 9, ref defaultTableBehavior, ref autoFitBehavior);
                wordtable1.Borders.Enable = 1;

                int iCountHead = 2;

                // ������� ����������.
                Word.Range wordcellrange1 = worddocument.Tables[1].Cell(iCountHead, 1).Range;
                wordcellrange1.Text = "� �.�";
                wordcellrange1.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                Word.Range wordcellrange2 = worddocument.Tables[1].Cell(iCountHead, 2).Range;

                wordcellrange2.Text = "� ��������.";
                wordcellrange2.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange3 = worddocument.Tables[1].Cell(iCountHead, 3).Range;

                wordcellrange3.Text = "���� �����������";
                wordcellrange3.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange4 = worddocument.Tables[1].Cell(iCountHead, 4).Range;
                wordcellrange4.Text = "�������������";
                wordcellrange4.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange5 = worddocument.Tables[1].Cell(iCountHead, 5).Range;
                wordcellrange5.Text = "������� ����������";
                wordcellrange5.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange6 = worddocument.Tables[1].Cell(iCountHead, 6).Range;
                wordcellrange6.Text = "���� ���������";
                wordcellrange6.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange7 = worddocument.Tables[1].Cell(iCountHead, 7).Range;
                wordcellrange7.Text = "����� ���������";
                wordcellrange7.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange8 = worddocument.Tables[1].Cell(iCountHead, 8).Range;
                wordcellrange8.Text = "���� ����������";
                wordcellrange8.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                Word.Range wordcellrange9 = worddocument.Tables[1].Cell(iCountHead, 9).Range;
                wordcellrange9.Text = "�����������";
                wordcellrange9.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                int iCount = 3;

                foreach (DataGridViewRow row in this.dataGridView1.Rows)
                {
                    if (row.Cells["���������"].Value.ToString().Length > 0)
                    {
                        // ������� ����������.
                        Word.Range wordcellrange1t = worddocument.Tables[1].Cell(iCount, 1).Range;
                        wordcellrange1t.Text = (iCount - 2).ToString().Trim();
                        wordcellrange1t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle; //.wdLineStyleTriple;

                        Word.Range wordcellrange2t = worddocument.Tables[1].Cell(iCount, 2).Range;

                        wordcellrange2t.Text = row.Cells["���������"].Value.ToString();
                        wordcellrange2t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange3t = worddocument.Tables[1].Cell(iCount, 3).Range;

                        DateTime dt = Convert.ToDateTime(row.Cells["����������"].Value);

                        wordcellrange3t.Text = dt.ToShortDateString().Trim();
                        wordcellrange3t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange4t = worddocument.Tables[1].Cell(iCount, 4).Range;
                        wordcellrange4t.Text = row.Cells["����������������������"].Value.ToString().Trim();
                        wordcellrange4t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange5t = worddocument.Tables[1].Cell(iCount, 5).Range;
                        wordcellrange5t.Text = row.Cells["�����������������"].Value.ToString().Trim();
                        wordcellrange5t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange6t = worddocument.Tables[1].Cell(iCount, 6).Range;
                        wordcellrange6t.Text = Convert.ToDateTime(row.Cells["����������"].Value).ToShortDateString();
                        wordcellrange6t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange7t = worddocument.Tables[1].Cell(iCount, 7).Range;
                        wordcellrange7t.Text = row.Cells["����������"].Value.ToString().Trim();
                        wordcellrange7t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange8t = worddocument.Tables[1].Cell(iCount, 8).Range;
                        wordcellrange8t.Text = Convert.ToDateTime(row.Cells["��������������"].Value).ToShortDateString().Trim();
                        wordcellrange8t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;

                        Word.Range wordcellrange9t = worddocument.Tables[1].Cell(iCount, 9).Range;
                        wordcellrange9t.Text = row.Cells["���������"].Value.ToString().Trim();
                        wordcellrange9t.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    }

                        iCount++;
                      
                }
            }
            catch (Exception ex)
            {
                Text = ex.Message;
            }

            // ��������� ����� 2 ������ �����������.
            object begCell = worddocument.Tables[1].Cell(1, 1).Range.Start;
            object endCell = worddocument.Tables[1].Cell(2, 1).Range.End;
            Word.Range wordcellrange = worddocument.Range(ref begCell, ref endCell);
            wordcellrange.Select();
            wordapp.Selection.Cells.Merge();

            // ��������� ������������ ������ ������.
            object begCell2 = worddocument.Tables[1].Cell(1, 2).Range.Start;
            object endCell2 = worddocument.Tables[1].Cell(1, 9).Range.End;
            wordcellrange = worddocument.Range(ref begCell2, ref endCell2);
            wordcellrange.Select();
            wordapp.Selection.Cells.Merge();

            worddocument.Tables[1].Cell(1, 2).Range.Text = "���������";

        }

        


    }
}