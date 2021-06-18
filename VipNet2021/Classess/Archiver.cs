using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Threading;
using SevenZip;


namespace RegKor.Classess
{
    public static class Archiver
    {
        /// <summary>
        /// ������ ����� archiveName, ������� �������� ���� fileNames
        /// </summary>
        /// <param name="archiver"></param>
        /// <param name="fileNames"></param>
        /// <param name="archiveName"></param>
        //public static void AddToArchive(string archiver, string fileNames,string archiveName, string tempDir)
        //{
        //    //try
        //    //{
        //        // ��������������� ��������
        //        if (!File.Exists(archiver))
        //            throw new Exception("��������� 7z �� ���� \"" + archiver +
        //            "\" �� ������");

        //        // ��������� ��������� ������ 7z
        //        ProcessStartInfo startInfo = new ProcessStartInfo();
        //        startInfo.FileName = archiver;
        //        // �������� � ����� � ������������ �������
        //        //startInfo.Arguments = " a -mx9 ";// +tempDir;

        //        startInfo.Arguments = " a -mx9 -omX=2";// +tempDir;
        //        // ��� ������
        //        startInfo.Arguments += "\"" + archiveName + "\"";
        //        // ����� ��� ���������
        //        startInfo.Arguments += " \"" + fileNames + "\"";
        //        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
        //        int sevenZipExitCode = 0;
        //        using (Process sevenZip = Process.Start(startInfo))
        //        {
        //            sevenZip.WaitForExit();
        //            sevenZipExitCode = sevenZip.ExitCode;
        //        }
        //        // ���� � ������� ���� �� ����������,
        //        //������� ��� ��� ����� 1 �������
        //        if (sevenZipExitCode != 0 && sevenZipExitCode != 1)
        //        {
        //            using (Process sevenZip = Process.Start(startInfo))
        //            {
        //                Thread.Sleep(1000);
        //                sevenZip.WaitForExit();
        //                switch (sevenZip.ExitCode)
        //                {
        //                    case 0: return; // ��� ������ � ��������������
        //                    case 1: return; // ���� ����������� ��������������
        //                    case 2: throw new Exception("��������� ������");
        //                    case 7: throw new Exception("������ � ��������� ������");
        //                    case 8:
        //                    throw new Exception("������������ ������ ��� ���������� ��������");
        //                    case 225:
        //                    throw new Exception("������������ ������� ���������� ��������");
        //                    default: throw new Exception("��������� 7z ������ ������������������� ��� ������: " + sevenZip.ExitCode.ToString());
        //                }
        //            }
        //        }

        //    //}
        //    //catch (Exception e)
        //    //{
        //    //    throw new Exception("SevenZip.AddToArchive: " + e.Message);
        //    //}
        //}

        /// <summary>
        /// ���������� ����������� 
        /// </summary>
        /// <param name="archiver">���� � ����������</param>
        /// <param name="fileNames"></param>
        /// <param name="archiveName"></param>
        /// <param name="tempDir">���������� ���� ��������� �����</param>
        public static void AddToArchive(string archiver, string fileNames, string archiveName, string tempDir)
        {
            // ���� � ����� ������� ����� �����.
            string source_folder = fileNames;

            // ��� ������.
            string archive_name = archiveName;

            // ���� � ����� 7zip.dll
            string library_source =  archiver;

            if (File.Exists(library_source))//���� ���������� 7zip ����������
            {
                //����������� �������� ���� � ������, ���������� ��� ������, � ��� �� ������� 7z.dll
                SevenZipExtractor.SetLibraryPath(library_source); //���������� ���������� 7zip
                SevenZipCompressor compressor = new SevenZipCompressor(); //��������� ���������� ����������
                compressor.ArchiveFormat = OutArchiveFormat.Zip; //�������� ������ ������. ������ "Zip" ����� ��������� "SevenZip".
                compressor.CompressionLevel = CompressionLevel.Ultra; // ������ ����� ������
                compressor.CompressionMode = CompressionMode.Create; //�������������� ���������
                compressor.TempFolderPath =  System.IO.Path.GetTempPath(); //����������� ��������� �����
                compressor.CompressDirectory(source_folder, archive_name, false); //��� ������� ������


                
                //string sTest = System.IO.Path.GetTempFileName();

                
            }


        }
    }
}
