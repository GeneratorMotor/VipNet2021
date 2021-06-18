using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Windows.Forms;

namespace RegKor.Classess
{
    public class WordPrint
    {
        private string nameReport = string.Empty;

        public WordPrint(string ��������������)
        {
            nameReport = ��������������;
        }

        public void Print(I����� tableData)
        {

            FileInfo fnDel = new FileInfo(System.Windows.Forms.Application.StartupPath + @"\���������\�����.doc");
            fnDel.Delete();

            string fName = nameReport;

            try
            {
                //��������� ������ � ����� ���������
                FileInfo fn = new FileInfo(System.Windows.Forms.Application.StartupPath + @"\������\�����.doc");
                fn.CopyTo(System.Windows.Forms.Application.StartupPath + @"\���������\" + fName + ".doc", true);
            }
            catch
            {
                MessageBox.Show("�������� � ��� ��� ������ ������� � ���� ����������. �������� ���� �������.");
                return;
            }

            string filName = System.Windows.Forms.Application.StartupPath + @"\���������\" + fName + ".doc";



            //System.Diagnostics.Process.Start("C:/asdasd.xls");

            //������ ����� Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //��������� ��������
            Microsoft.Office.Interop.Word.Document doc = null;

            object fileName = filName;
            object falseValue = false;
            object trueValue = true;
            object missing = Type.Missing;
            object writePasswordDocument = "12A86Asd";

            doc = app.Documents.Open(ref fileName, ref missing, ref trueValue,
            ref missing, ref missing, ref missing, ref missing, ref writePasswordDocument,
            ref missing, ref missing, ref missing, ref missing, ref trueValue,
            ref missing, ref missing, ref missing);

            //NAMEREPORT
            ////����� ��������
            object wdrepl = WdReplace.wdReplaceAll;
            //object searchtxt = "GreetingLine";
            object searchtxt = "NAMEREPORT";
            object newtxt = (object)nameReport;
            //object frwd = true;
            object frwd = false;
            doc.Content.Find.Execute(ref searchtxt, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd, ref missing, ref missing, ref newtxt, ref wdrepl, ref missing, ref missing,
            ref missing, ref missing);

            // ������ ��� ������.
            �������������������� report = (��������������������)tableData;

            //�������� �������
            object bookNaziv = "�������";
            Range wrdRng = doc.Bookmarks.get_Item(ref  bookNaziv).Range;

            object behavior = Microsoft.Office.Interop.Word.WdDefaultTableBehavior.wdWord8TableBehavior;
            object autobehavior = Microsoft.Office.Interop.Word.WdAutoFitBehavior.wdAutoFitWindow;


            Microsoft.Office.Interop.Word.Table table = doc.Tables.Add(wrdRng, 1, 10, ref behavior, ref autobehavior);
            table.Range.ParagraphFormat.SpaceAfter = 6;

            table.Columns[1].Width = 40;
            table.Columns[2].Width = 100;
            table.Columns[3].Width = 60;
            table.Columns[4].Width = 60;
            table.Columns[5].Width = 120;
            table.Columns[6].Width = 60;
            table.Columns[7].Width = 60;
            table.Columns[8].Width = 80;
            table.Columns[9].Width = 100;
            table.Columns[10].Width = 80;
            table.Borders.Enable = 1; // ����� - �������� �����
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 10;

            //������� �����
            int i = 1;

            // ���������� ����� � ������.
            int iCountRow = report.DataGridView1.Rows.Count;

            // ����� �����.
            table.Cell(i, 1).Range.Text = "� �.�";
            table.Cell(i, 2).Range.Text = "�������������";
            table.Cell(i, 3).Range.Text = "���� ���������";
            table.Cell(i, 4).Range.Text = "����� ���������";
            table.Cell(i, 5).Range.Text = "������� ����������";
            table.Cell(i, 6).Range.Text = "���� ��������";
            table.Cell(i, 7).Range.Text = "����� ��������";
            table.Cell(i, 8).Range.Text = "���� ����������";
            table.Cell(i, 9).Range.Text = "��������� ����������";
            table.Cell(i, 10).Range.Text = "�����������";

            Object beforeRow1 = Type.Missing;
            table.Rows.Add(ref beforeRow1);

            i++;
            foreach (DataGridViewRow r in report.DataGridView1.Rows)
            {
                    if (i <= iCountRow)
                    {
                        int k = i - 1;
                        table.Cell(i, 1).Range.Text = (k).ToString();
                        table.Cell(i, 2).Range.Text = r.Cells["�������������"].Value.ToString().Trim();
                        table.Cell(i, 3).Range.Text = r.Cells["����������"].Value.ToString().Trim();
                        table.Cell(i, 4).Range.Text = r.Cells["����������"].Value.ToString().Trim();
                        table.Cell(i, 5).Range.Text = r.Cells["�����������������"].Value.ToString().Trim();
                        table.Cell(i, 6).Range.Text = r.Cells["����������"].Value.ToString().Trim();
                        table.Cell(i, 7).Range.Text = r.Cells["����� ��������"].Value.ToString().Trim();
                        table.Cell(i, 8).Range.Text = r.Cells["��������������"].Value.ToString().Trim();
                        table.Cell(i, 9).Range.Text = r.Cells["�������������������"].Value.ToString().Trim();
                        table.Cell(i, 10).Range.Text = r.Cells["�����������"].Value.ToString().Trim();
                    }
                //}

                Object beforeRow2 = Type.Missing;
                table.Rows.Add(ref beforeRow2);

                i++;
            }

            //������ ��������� ������
            //table.Rows[i].Delete();

            //������� ������������ ��������
            app.Visible = true;


        }
    }
}
