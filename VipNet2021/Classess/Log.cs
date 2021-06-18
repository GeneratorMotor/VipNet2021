using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace RegKor.Classess
{
    /// <summary>
    /// ����� ��� �������� ��� �����.
    /// </summary>
    public static class Log
    {
        /// <summary>
        /// ���������� ���������� � ���.
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
                    //writ.BaseStream.Seek(fs.Length, SeekOrigin.End);//������ � ����� �����
                    writ.WriteLine(logText);
                }
            }
            else
            {
                // ������� loq ���� � ������� � ���� log.
                //using (FileStream fs = File.Open(filePatch,FileMode.Open))
                using (TextWriter writ = File.AppendText(filePatch))
                {
                    //writ.BaseStream.Seek(fs.Length, SeekOrigin.End);//������ � ����� �����
                    writ.Write(logText + ",");
                }
                
            }
        }
    }
}
