using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace RegKor.Classess
{
    /// <summary>
    /// Класс определяет режим работы программы до 2012 года или 2012 год и позже
    /// </summary>
    class ВерсияРаботы
    {
        /// <summary>
        /// Записываем в файл значение флага true или false
        /// </summary>
        /// <param name="версияРаботыБазы"></param>
        public void ВерсияБазы(bool версияРаботыБазы)
        {
            //@"..\report\KontrolMessage.rpt"
            FileStream file = File.Create(@"..\File\BinaryVersion.dat");

            //Запишем данные в файл
            BinaryWriter write = new BinaryWriter(file);
            write.Write(версияРаботыБазы);
            write.Close();
            file.Close();
        }

        /// <summary>
        /// Считываем из файла значение флага true или false
        /// </summary>
        /// <returns></returns>
        public bool УзнатьВерсиюРаботы()
        {
            //Получим файл
            FileStream file = File.Open(@"..\File\BinaryVersion.dat", FileMode.Open);

            //Прочтём данные из файла
            BinaryReader read = new BinaryReader(file);

            bool версия = read.ReadBoolean();
            read.Close();

            file.Close();
            return версия;
        }
    }
}
