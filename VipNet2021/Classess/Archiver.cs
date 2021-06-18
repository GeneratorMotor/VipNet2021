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
        /// Создаёт архив archiveName, который содержит файл fileNames
        /// </summary>
        /// <param name="archiver"></param>
        /// <param name="fileNames"></param>
        /// <param name="archiveName"></param>
        //public static void AddToArchive(string archiver, string fileNames,string archiveName, string tempDir)
        //{
        //    //try
        //    //{
        //        // Предварительные проверки
        //        if (!File.Exists(archiver))
        //            throw new Exception("Архиватор 7z по пути \"" + archiver +
        //            "\" не найден");

        //        // Формируем параметры вызова 7z
        //        ProcessStartInfo startInfo = new ProcessStartInfo();
        //        startInfo.FileName = archiver;
        //        // добавить в архив с максимальным сжатием
        //        //startInfo.Arguments = " a -mx9 ";// +tempDir;

        //        startInfo.Arguments = " a -mx9 -omX=2";// +tempDir;
        //        // имя архива
        //        startInfo.Arguments += "\"" + archiveName + "\"";
        //        // файлы для запаковки
        //        startInfo.Arguments += " \"" + fileNames + "\"";
        //        startInfo.WindowStyle = ProcessWindowStyle.Hidden;
        //        int sevenZipExitCode = 0;
        //        using (Process sevenZip = Process.Start(startInfo))
        //        {
        //            sevenZip.WaitForExit();
        //            sevenZipExitCode = sevenZip.ExitCode;
        //        }
        //        // Если с первого раза не получилось,
        //        //пробуем еще раз через 1 секунду
        //        if (sevenZipExitCode != 0 && sevenZipExitCode != 1)
        //        {
        //            using (Process sevenZip = Process.Start(startInfo))
        //            {
        //                Thread.Sleep(1000);
        //                sevenZip.WaitForExit();
        //                switch (sevenZip.ExitCode)
        //                {
        //                    case 0: return; // Без ошибок и предупреждений
        //                    case 1: return; // Есть некритичные предупреждения
        //                    case 2: throw new Exception("Фатальная ошибка");
        //                    case 7: throw new Exception("Ошибка в командной строке");
        //                    case 8:
        //                    throw new Exception("Недостаточно памяти для выполнения операции");
        //                    case 225:
        //                    throw new Exception("Пользователь отменил выполнение операции");
        //                    default: throw new Exception("Архиватор 7z вернул недокументированный код ошибки: " + sevenZip.ExitCode.ToString());
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
        /// Архивирует дирректорию 
        /// </summary>
        /// <param name="archiver">Путь к архиватору</param>
        /// <param name="fileNames"></param>
        /// <param name="archiveName"></param>
        /// <param name="tempDir">Директория куда сохранить архив</param>
        public static void AddToArchive(string archiver, string fileNames, string archiveName, string tempDir)
        {
            // Путь к папке которую нужно сжать.
            string source_folder = fileNames;

            // Имя архива.
            string archive_name = archiveName;

            // Путь к файлу 7zip.dll
            string library_source =  archiver;

            if (File.Exists(library_source))//Если библиотека 7zip существует
            {
                //Попробовать написать путь к архиву, посмотреть что выйдет, а так же скачать 7z.dll
                SevenZipExtractor.SetLibraryPath(library_source); //Подгружаем библиотеку 7zip
                SevenZipCompressor compressor = new SevenZipCompressor(); //Объявляем переменную архиватора
                compressor.ArchiveFormat = OutArchiveFormat.Zip; //Выбираем формат архива. Вместо "Zip" можно поставить "SevenZip".
                compressor.CompressionLevel = CompressionLevel.Ultra; // ультра режим сжатия
                compressor.CompressionMode = CompressionMode.Create; //подтверждаются настройки
                compressor.TempFolderPath =  System.IO.Path.GetTempPath(); //объявляется временная папка
                compressor.CompressDirectory(source_folder, archive_name, false); //сам процесс сжатия


                
                //string sTest = System.IO.Path.GetTempFileName();

                
            }


        }
    }
}
