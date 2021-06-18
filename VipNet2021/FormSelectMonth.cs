using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace RegKor
{
    public partial class FormSelectMonth : Form
    {
        string началоћес€ца;
        string год;
        string весь√од;
        string окончаниећес€ца;

        /// <summary>
        ///  ¬озвращает первый день временного диапазона
        /// </summary>
        public string Getѕервыйƒень
        {
            get
            {
                return началоћес€ца;
            }
            set
            {
                началоћес€ца = value;
            }
        }

        /// <summary>
        /// ¬озвращает последний день мес€ца.
        /// </summary>
        public string Get райнийƒень
        {
            get
            {
                return окончаниећес€ца;
            }
            set
            {
                окончаниећес€ца = value;
            }
        }

        /// <summary>
        /// ¬есь год.
        /// </summary>
        public string ¬есь√од
        {
            get
            {
                return весь√од;
            }
            set
            {
                весь√од = value;
            }
        }

        /// <summary>
        /// ¬цыбранный год.
        /// </summary>
        public string ¬ыбранный√од
        {
            get
            {
                return год;
            }
            set
            {
                год = value;
            }
        }

        public FormSelectMonth()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.Text != "")
            {

                // ѕолучим год.
                int _год = Convert.ToInt32(this.¬ыбранный√од);

                // ѕолучим номер мес€ца.
                string названиећес€ца = this.comboBox1.Text;
                string нумћес = string.Empty;

                switch (названиећес€ца)
                {
                    case "январь":
                        нумћес = "01";
                        break;
                    case "‘евраль":
                        нумћес = "02";
                        break;
                    case "ћарт":
                        нумћес = "03";
                        break;
                    case "јпрель":
                        нумћес = "04";
                        break;
                    case "ћай":
                        нумћес = "05";
                        break;
                    case "»юнь":
                        нумћес = "06";
                        break;
                    case "»юль":
                        нумћес = "07";
                        break;
                    case "јвгуст":
                        нумћес = "08";
                        break;
                    case "—ент€брь":
                        нумћес = "09";
                        break;
                    case "ќкт€брь":
                        нумћес = "10";
                        break;
                    case "Ќо€брь":
                        нумћес = "11";
                        break;
                    case "ƒекабрь":
                        нумћес = "12";
                        break;
                }

                if (названиећес€ца != "¬есь √од")
                {

                    // ѕолучим количество дней в мес€це.
                    int мес€ц¬ыбранный = Convert.ToInt32(нумћес);

                    int num√од = _год + 1;

                    int countDay = DateTime.DaysInMonth(num√од, мес€ц¬ыбранный);

                    // ¬озвратим дату начала мес€ца.
                    this.Getѕервыйƒень = "01." + нумћес + "." + num√од.ToString().Trim();

                    this.Get райнийƒень = countDay.ToString().Trim() + "." + нумћес + "." + num√од.ToString().Trim();
                }
                else
                {
                    int num√од = _год + 1;

                    this.Getѕервыйƒень = "01.01." + num√од.ToString().Trim();

                    this.Get райнийƒень = "31.12." + num√од.ToString().Trim();

                }
            }
            else
            {
                MessageBox.Show("Ќе выбран мес€ц");
                this.Close();
            }

        }
    }
}