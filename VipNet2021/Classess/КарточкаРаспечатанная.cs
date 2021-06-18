using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RegKor.Classess
{
    class КарточкаРаспечатанная
    {
        private string _НомерКарточки1;
        private string _НомерКарточки2;
        private string _текущийГод;
        private string _будущийГод;


        /// <summary>
        /// Хранит первое января будущего года
        /// </summary>
        public string БудущийГод
        {
            get
            {
                return _будущийГод;
            }
            set
            {
                _будущийГод = value;
            }
        }

        /// <summary>
        /// Хранит первое января текущего года
        /// </summary>
        public string ТекущийГод
        {
            get
            {
                return _текущийГод;
            }
            set
            {
                _текущийГод = value;
            }
        }


        /// <summary>
        /// хранит номер первой каротчки в отчёте
        /// </summary>
        public string ПервыйНомерКарточки
        {
            get
            {
                return _НомерКарточки1;
            }
            set
            {
                _НомерКарточки1 = value;
            }
        }

        /// <summary>
        /// Хранит постледний номер карточки в отчёте
        /// </summary>
        public string КрайнийНомерКарточки
        {
            get
            {
                return _НомерКарточки2;
            }
            set
            {
                _НомерКарточки2 = value;
            }
        }




        /// <summary>
        /// Формирует карточку
        /// </summary>
        /// <returns></returns>
        //public DataSet ПолучитьДанные()
        public List<Карточка> ПолучитьДанные()
        {
            string arrayStringFirst = НомерКарточки(ПервыйНомерКарточки);
            string arrayStringLast = НомерКарточки(КрайнийНомерКарточки);

           
            //Переменная хранит строку запроса к БД
            string Query = string.Empty;

            if (arrayStringFirst != null && arrayStringLast == null)
            {
                //получаем номер карточки
                string номерВхода = arrayStringFirst;

                //стирока запроса
                Query = "select * from dbo.Выборка " +
                        "where номерПП = " + номерВхода + " and ДатаПоступ >= '" + ТекущийГод + "' and ДатаПоступ <= '" + БудущийГод + "' ";
            }

            if (arrayStringFirst != null && arrayStringLast != null)
            {

                //string номерВхода_F = arrayStringFirst[0];
                string номерПП_F = arrayStringFirst;

                //string номерВход_L = arrayStringLast[0];
                string номерПП_L = arrayStringLast;

                Query = "select * from dbo.Выборка " +
                        "where номерПП >= " + номерПП_F + " and номерПП <= " + номерПП_L + " and ДатаПоступ >= '" + ТекущийГод + "' and ДатаПоступ <= '" + БудущийГод + "' ";
            }

            ПодключитьБД строкаПодключения = new ПодключитьБД();
            SqlConnection con = new SqlConnection(строкаПодключения.СтрокаПодключения());

            //SqlDataAdapter da = new SqlDataAdapter(Query, con);
            SqlCommand com = new SqlCommand(Query, con);

            con.Open();
            SqlDataReader read = com.ExecuteReader();

            //DS1.ViewКарточкаDataTable tab = new DS1.ViewКарточкаDataTable();

            DS1.ВыборкаDataTable tab = new DS1.ВыборкаDataTable();
            //DataTable tab = new DataTable();

            List<Карточка> list = new List<Карточка>();

                ////В связи с тем что номер исходящий представляет собой склейку из двух полей НомерПП/ИсходящийНомер
            //Заполним таблицу в ручную
            while (read.Read())
            {
                Карточка карточка = new Карточка();
                карточка.id_карточки = read["id_карточки"].ToString();
                карточка.ОписаниеДокумента = read["ОписаниеДокумента"].ToString();

                карточка.ОписаниеКорреспондента = read["ОписаниеКорреспондента"].ToString();
                карточка.НомерИсход = read["НомерИсход"].ToString();

                карточка.НомерВход = read["НомерВход"].ToString();
                DateTime dtИсход = Convert.ToDateTime(read["ДатаИсхода"].ToString());
                карточка.ДатаИсхода = ConvertDate(dtИсход); //dtИсход.ToShortDateString();

                карточка.ДатаПоступ = ConvertDate(Convert.ToDateTime(read["ДатаПоступ"].ToString()));


                карточка.КраткоеСодержание = read["КраткоеСодержание"].ToString();

                //DateTime dt = Convert.ToDateTime(read["СрокВыполнения"].ToString());

                //карточка.СрокВыполнения = dt.ToShortDateString();
                карточка.Резолюция = read["Резолюция"].ToString();

                карточка.ВДело = read["ВДело"].ToString();
                карточка.НаКонтроле = read["НаКонтроле"].ToString();

                карточка.РезультатВыполнения = read["РезультатВыполнения"].ToString();

                // Проверим карточка в деле или нет.
                bool flag = Convert.ToBoolean(read["ВДело"]);
                if (flag != true)
                {

                    DateTime dt = Convert.ToDateTime(read["СрокВыполнения"].ToString());
                    //string[] dateArry = dt.ToShortDateString().Split(".");

                    // Сформируем дату в виде 01-января-2014

                    

                    карточка.СрокВыполнения = ConvertDate(dt);
                }
                else
                {
                    карточка.СрокВыполнения = "";
                }

                list.Add(карточка);

                //DataRow row = tab.NewRow();
             
                //string ДатаИсхода = read["ДатаИсхода"].ToString().Substring(0, 10);//.Remove(10,9);
                //string ДатаПоступления = read["ДатаПоступ"].ToString();//.Remove(10, 9);

                //int iTest = Convert.ToInt32(read["id_карточки"]);

                //row["id_карточки"] = read["id_карточки"];

                //row["ДатаИсхода"] = ДатаИсхода;
                //row["ДатаПоступ"] = ДатаПоступления;

                //row["НомерВход"] = read["НомерВход"];
                //row["НомерИсход"] = read["НомерИсход"];

                //row["ОписаниеДокумента"] = read["ОписаниеДокумента"];
                //row["ОписаниеКорреспондента"] = read["ОписаниеКорреспондента"];

                //row["КраткоеСодержание"] = read["КраткоеСодержание"];
                //row["Резолюция"] = read["Резолюция"];

                //row["РезультатВыполнения"] = read["РезультатВыполнения"];
                ////row["СрокВыполнения"] = read["СрокВыполнения"];

                //row["НаКонтроле"] = read["НаКонтроле"];


                //row["ВДело"] = read["ВДело"];

                //bool flag = Convert.ToBoolean(read["ВДело"]);
                //if (flag != true)
                //{
                //    //row["НаКонтроле"] = read["НаКонтроле"];
                //    row["СрокВыполнения"] = read["СрокВыполнения"];

                //}
                //else
                //{
                //    //row["НаКонтроле"] = "";
                //    row["СрокВыполнения"] = null;
                //}
                //row["номерПП"] = read["номерПП"];
                
                //tab.Rows.Add(row);
            }

            //DataSet ds = new DataSet();
            //ds.Tables.Add(tab);

            con.Close();
            //return ds;
            return list;

        }

        /// <summary>
        /// Возвращает номер ПП
        /// </summary>
        /// <param name="номерКарточки"></param>
        /// <returns></returns>
        private string НомерКарточки(string номерКарточки)
        {
            string sNum = string.Empty;
            //string[] arrayString = null;
            if (номерКарточки != "")
            {
                sNum = номерКарточки;
                return sNum;
            }
            else
            {
                sNum = null;
            }
            return sNum;

        }

        private string ConvertDate(DateTime dt)
        {
            string dat = dt.ToShortDateString();
            string[] dateArry = dat.Split('.');

            // Переменная для хранения названия месяца.
            string nameMotch = string.Empty;

            string montch = dateArry[1];
            switch (montch)
            {
                case "01":
                    nameMotch = "Январь";
                    break;
                case "02":
                    nameMotch = "Февраль";
                    break;
                case "03":
                    nameMotch = "Март";
                    break;
                case "04":
                    nameMotch = "Апрель";
                    break;
                case "05":
                    nameMotch = "Май";
                    break;
                case "06":
                    nameMotch = "Июнь";
                    break;
                case "07":
                    nameMotch = "Июль";
                    break;
                case "08":
                    nameMotch = "Август";
                    break;
                case "09":
                    nameMotch = "Сентябрь";
                    break;
                case "10":
                    nameMotch = "Октябрь";
                    break;
                case "11":
                    nameMotch = "Ноябрь";
                    break;
                case "12":
                    nameMotch = "Декабрь";
                    break;
                     
            }


            string date = dateArry[0] + "-" + nameMotch + "-" + dateArry[2];
            return date;

        }

       


    }
}
