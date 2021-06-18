using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace RegKor.Classess
{
    /// <summary>
    /// Класс для создания лог файла.
    /// </summary>
    public static class Log
    {
        /// <summary>
        /// Записываем информацию в лог.
        /// </summary>
        /// <param name="filePatch"></param>
        /// <param name="logText"></param>
        public static void WriteLine(string filePatch, string logText)
        {
            if (File.Exists(filePatch) == false)
            {
                using(FileStream fs = File.Create(filePatch))
                using (TextWriter writ = new StreamWriter(fs))
                {
                    //writ.BaseStream.Seek(fs.Length, SeekOrigin.End);//запись в конец файла
                    writ.WriteLine(logText);
                }
            }
            else
            {
                // ОТкроем loq файл и запишем в него log.
                //using (FileStream fs = File.Open(filePatch,FileMode.Open))
                using (TextWriter writ = File.AppendText(filePatch))
                {
                    //writ.BaseStream.Seek(fs.Length, SeekOrigin.End);//запись в конец файла
                    writ.Write(logText + ",");
                }
                
            }
        }
    }
}
