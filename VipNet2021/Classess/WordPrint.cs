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

        public WordPrint(string названиеОтчета)
        {
            nameReport = названиеОтчета;
        }

        public void Print(IОтчет tableData)
        {

            FileInfo fnDel = new FileInfo(System.Windows.Forms.Application.StartupPath + @"\Документы\Отчет.doc");
            fnDel.Delete();

            string fName = nameReport;

            try
            {
                //Скопируем шаблон в папку Документы
                FileInfo fn = new FileInfo(System.Windows.Forms.Application.StartupPath + @"\Шаблон\Отчет.doc");
                fn.CopyTo(System.Windows.Forms.Application.StartupPath + @"\Документы\" + fName + ".doc", true);
            }
            catch
            {
                MessageBox.Show("Возможно у вас уже открыт договор с этим льготником. Закройте этот договор.");
                return;
            }

            string filName = System.Windows.Forms.Application.StartupPath + @"\Документы\" + fName + ".doc";



            //System.Diagnostics.Process.Start("C:/asdasd.xls");

            //Создаём новый Word.Application
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

            //Загружаем документ
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
            ////Номер договора
            object wdrepl = WdReplace.wdReplaceAll;
            //object searchtxt = "GreetingLine";
            object searchtxt = "NAMEREPORT";
            object newtxt = (object)nameReport;
            //object frwd = true;
            object frwd = false;
            doc.Content.Find.Execute(ref searchtxt, ref missing, ref missing, ref missing, ref missing, ref missing, ref frwd, ref missing, ref missing, ref newtxt, ref wdrepl, ref missing, ref missing,
            ref missing, ref missing);

            // Данные для отчета.
            ОтчетОВходДокументах report = (ОтчетОВходДокументах)tableData;

            //Вставить таблицу
            object bookNaziv = "таблица";
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
            table.Borders.Enable = 1; // Рамка - сплошная линия
            table.Range.Font.Name = "Times New Roman";
            table.Range.Font.Size = 10;

            //счётчик строк
            int i = 1;

            // Количество строк в отчете.
            int iCountRow = report.DataGridView1.Rows.Count;

            // Пишем шапку.
            table.Cell(i, 1).Range.Text = "№ п.п";
            table.Cell(i, 2).Range.Text = "Корреспондент";
            table.Cell(i, 3).Range.Text = "Дата исходящая";
            table.Cell(i, 4).Range.Text = "Номер исходящий";
            table.Cell(i, 5).Range.Text = "Краткое содержание";
            table.Cell(i, 6).Range.Text = "Дата входящая";
            table.Cell(i, 7).Range.Text = "Номер входящий";
            table.Cell(i, 8).Range.Text = "Срок исполнения";
            table.Cell(i, 9).Range.Text = "Результат исполнения";
            table.Cell(i, 10).Range.Text = "Исполнитель";

            Object beforeRow1 = Type.Missing;
            table.Rows.Add(ref beforeRow1);

            i++;
            foreach (DataGridViewRow r in report.DataGridView1.Rows)
            {
                    if (i <= iCountRow)
                    {
                        int k = i - 1;
                        table.Cell(i, 1).Range.Text = (k).ToString();
                        table.Cell(i, 2).Range.Text = r.Cells["Корреспондент"].Value.ToString().Trim();
                        table.Cell(i, 3).Range.Text = r.Cells["ДатаИсхода"].Value.ToString().Trim();
                        table.Cell(i, 4).Range.Text = r.Cells["НомерИсход"].Value.ToString().Trim();
                        table.Cell(i, 5).Range.Text = r.Cells["КраткоеСодержание"].Value.ToString().Trim();
                        table.Cell(i, 6).Range.Text = r.Cells["ДатаПоступ"].Value.ToString().Trim();
                        table.Cell(i, 7).Range.Text = r.Cells["Номер входящий"].Value.ToString().Trim();
                        table.Cell(i, 8).Range.Text = r.Cells["СрокВыполнения"].Value.ToString().Trim();
                        table.Cell(i, 9).Range.Text = r.Cells["РезультатВыполнения"].Value.ToString().Trim();
                        table.Cell(i, 10).Range.Text = r.Cells["Исполнитель"].Value.ToString().Trim();
                    }
                //}

                Object beforeRow2 = Type.Missing;
                table.Rows.Add(ref beforeRow2);

                i++;
            }

            //удалим последную строку
            //table.Rows[i].Delete();

            //откроем получившейся документ
            app.Visible = true;


        }
    }
}
