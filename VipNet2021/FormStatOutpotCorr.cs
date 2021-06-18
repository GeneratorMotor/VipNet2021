using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using RegKor.Classess;

namespace RegKor
{
    public partial class FormStatOutpotCorr : Form
    {
        private RangeDate rd;

        // ������� � �������.
        private DataTable rez;

        // ������� � �������������.
        private DataTable rezIsp;

        /// <summary>
        /// ������ �������� ���.
        /// </summary>
        public RangeDate �����������
        {
            get
            {
                return rd;
            }
            set
            {
                rd = value;
            }
        }

        public FormStatOutpotCorr()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
           
            ����������������������������������� getData = new �����������������������������������(rd);

            DataTable tPerson = getData.�����������();

            List<��������������������������> list = new List<��������������������������>();

            // ����� ������.
            �������������������������� itemHead = new ��������������������������();
            itemHead.������� = "� �.�.";
            itemHead.�������������������������� = "������������ ��������������";
            itemHead.������������������������ = "���������� ��������� ����������";
            itemHead.����������������� = "�������� ��������";
            itemHead.VipNet = "VipNet";
            itemHead.Fax = "����";
            itemHead.EMail = "e-mail";
            itemHead.����������� = "�����������";

            list.Add(itemHead);

            // ��������� �� ������������.
            foreach (DataRow row in tPerson.Rows)
            {
                // ������� �������� ����� ��� �������� ������������.
                DataRow[] rows = rez.Select("������������������ = '" + row["������������������"].ToString().Trim() + "' ");

                int iCount = 1;
                foreach (DataRow r in rows)
                {
                    �������������������������� itm = new ��������������������������();
                    itm.������� = iCount.ToString();
                    itm.�������������������������� = r["����������������������"].ToString().Trim();

                    itm.������������������������ = r["�������������������������"].ToString().Trim();
                    itm.����������������� = r["������"].ToString().Trim();
                    itm.EMail = r["����"].ToString().Trim();
                    itm.VipNet = r["���"].ToString().Trim();
                    itm.Fax = r["����"].ToString().Trim();
                    itm.����������� = r["������������������"].ToString().Trim();

                    list.Add(itm);

                    iCount++;
                }

               

                DataRow rCount = getData.��������������������������(row["������������������"].ToString().Trim()).Rows[0];

                // �������� ������ ����� ��� �����������.
                �������������������������� itCount = new ��������������������������();
                itCount.������� = "����� �� ����������� " + rCount["������������������"].ToString().Trim();
                itCount.����������� = "----------";
                itCount.������������������������ = rCount["������������������"].ToString().Trim();
                itCount.����������������� = rCount["������"].ToString().Trim();
                itCount.EMail = rCount["����"].ToString().Trim();
                itCount.VipNet = rCount["���"].ToString().Trim();
                itCount.Fax = rCount["����"].ToString().Trim();
                itCount.����������� = rCount["������������������"].ToString().Trim();

                list.Add(itCount);
            }

            int iCount2 = 1;

            int count������������ = 0;
            int count��������� = 0;
            int count������� = 0;
            int count������ = 0;
            int count������� = 0;

            DataTable dtRowsCorrespondent = getData.����������������������();
            // ������� ����� �� �� �������� �� ���������������.
            foreach (DataRow r in dtRowsCorrespondent.Rows)
            {
                �������������������������� itm = new ��������������������������();
                itm.������� = iCount2.ToString();
                itm.�������������������������� = r["����������������������"].ToString().Trim();

                itm.������������������������ = r["�������������������������"].ToString().Trim();

                if (DBNull.Value != r["�������������������������"])
                {
                    count������������ += Convert.ToInt32(r["�������������������������"]);
                }

                itm.����������������� = r["������"].ToString().Trim();

                if (DBNull.Value != r["������"])
                {
                    count��������� += Convert.ToInt32(r["������"]);
                }

                itm.EMail = r["����"].ToString().Trim();

                if (DBNull.Value != r["����"])
                {
                    count������� += Convert.ToInt32(r["����"]);
                }

                itm.VipNet = r["���"].ToString().Trim();

                if (DBNull.Value != r["���"])
                {
                    count������ += Convert.ToInt32(r["���"]);
                }

                itm.Fax = r["����"].ToString().Trim();

                if (DBNull.Value != r["����"])
                {
                    count������� += Convert.ToInt32(r["����"]);
                }

                itm.����������� = "-----";

                list.Add(itm);

                iCount2++;
            }

            �������������������������� count = new ��������������������������();
            count.������� = "����� � ����� �� ��������";
            count.�������������������������� = "-----";
            count.������������������������ = count������������.ToString();
            count.����������������� = count���������.ToString();
            count.EMail = count�������.ToString();
            count.VipNet = count������.ToString();
            count.Fax = count�������.ToString();

            list.Add(count);

            Report����������������������������� reportStatistic = new Report�����������������������������("���������� �� ��������� ��������������� c " + this.�����������.DataStart.ToShortDateString() + " �� " + this.�����������.DataEnd.ToShortDateString() + " �� ���� �. ��������");

            PrintReportStaticOutputDoc report = new PrintReportStaticOutputDoc(reportStatistic);
            report.ListDate = list;
            report.Execute();


            //ExcelPrint excel = new ExcelPrint(" ���������� �� ��������� ��������������� c " + this.�����������.DataStart.ToShortDateString() + " �� " + this.�����������.DataEnd.ToShortDateString() + " �� ���� �. ��������");

            //// �������� ��������� � ������� ���������� �����.
            ////excel.Print���������������������������������(list);

            //// �������� ������ ���� ������������� ����� �� ���� �����.
            //excel.SaveFileCSV(list);

            //// ��������� �� ��� � Excel.
            this.Close();
        }

        private void FormStatOutpotCorr_Load(object sender, EventArgs e)
        {
            ����������������������������������� getData = new �����������������������������������(rd);
            
            rez = getData.����������DataGridView();

            this.dataGridView1.DataSource = rez;

            this.dataGridView1.Columns["����"].HeaderText = "e-mail";

        }
    }
}