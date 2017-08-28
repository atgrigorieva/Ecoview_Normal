using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using SWF = System.Windows.Forms;

namespace Ecoview_Normal
{
    public partial class Ecoview : Form
    {
        public string edconctr;
        public string SposobZadan;
        public string Zavisimoct;
        public string aproksim;
        public int selet_rezim;
        public bool IzmerenieOpen = false;
        public bool IzmerCreate = false;
        public bool IzmerCreate1 = false;
        public double[][] massGEMultiAbs;
        public double[][] massGEMultiT;
        public int countButtonClick;
        public string filepath;
        public Microsoft.Office.Interop.Excel.Workbook workBook;
        public Microsoft.Office.Interop.Excel.Worksheet workSheet;
        public string WL_grad1;
        public int IzmerFr_count;
        public double CellOpt;
        public Ecoview(int selet_rezim1)
        {
            InitializeComponent();
            this.selet_rezim = selet_rezim1;

            switch (selet_rezim)
            {
                case 1:
                    tabControl2.TabPages.Remove(tabPage3);
                    tabControl2.TabPages.Remove(tabPage4);
                    tabControl2.TabPages.Remove(tabPage2);
                    tabControl2.TabPages.Remove(tabPage5);
                    tabControl2.TabPages.Remove(tabPage6);
                    tabControl2.TabPages.Remove(tabPage7);
                    tabControl2.TabPages.Remove(tabPage8);
                    tabControl2.TabPages.Remove(tabPage9);
                    this.Text = "Eciview Normal v1.0 Фотометрический режим";
                    tabControl2.SelectedIndex = 2;
                    tabControl2.SelectTab(tabPage1);
                    break;
                case 2:
                    tabControl2.TabPages.Remove(tabPage1);
                    tabControl2.TabPages.Remove(tabPage2);
                    tabControl2.TabPages.Remove(tabPage5);
                    tabControl2.TabPages.Remove(tabPage6);
                    tabControl2.TabPages.Remove(tabPage7);
                    tabControl2.TabPages.Remove(tabPage8);
                    tabControl2.TabPages.Remove(tabPage9);
                    this.Text = "Eciview Normal v1.0 Количественный режим";
                    tabControl2.SelectedIndex = 0;
                    tabControl2.SelectTab(tabPage3);
                    button13.Enabled = false;
                    break;
                case 3:
                    tabControl2.TabPages.Remove(tabPage3);
                    tabControl2.TabPages.Remove(tabPage4);
                    tabControl2.TabPages.Remove(tabPage1);
                    tabControl2.TabPages.Remove(tabPage5);
                    tabControl2.TabPages.Remove(tabPage6);
                    tabControl2.TabPages.Remove(tabPage7);
                    tabControl2.TabPages.Remove(tabPage8);
                    tabControl2.TabPages.Remove(tabPage9);
                    this.Text = "Eciview Normal v1.0 Многоволновой режим";
                    tabControl2.SelectedIndex = 3;
                    button13.Enabled = false;
                    tabControl2.SelectTab(tabPage2);
                    break;
                case 4:
                    tabControl2.TabPages.Remove(tabPage3);
                    tabControl2.TabPages.Remove(tabPage4);
                    tabControl2.TabPages.Remove(tabPage2);
                    tabControl2.TabPages.Remove(tabPage1);
                    tabControl2.TabPages.Remove(tabPage6);
                    tabControl2.TabPages.Remove(tabPage7);
                    tabControl2.TabPages.Remove(tabPage8);
                    tabControl2.TabPages.Remove(tabPage9);
                    this.Text = "Eciview Normal v1.0 Кинетический режим";
                    tabControl2.SelectedIndex = 4;
                    button13.Enabled = false;
                    tabControl2.SelectTab(tabPage5);
                    break;
                case 9:
                    tabControl2.TabPages.Remove(tabPage4);
                    tabControl2.TabPages.Remove(tabPage5);
                    tabControl2.TabPages.Remove(tabPage2);
                    tabControl2.TabPages.Remove(tabPage1);
                    tabControl2.TabPages.Remove(tabPage6);
                    tabControl2.TabPages.Remove(tabPage7);
                    tabControl2.TabPages.Remove(tabPage8);
                    this.Text = "Eciview Normal v1.0 Работа в Excel";
                    tabControl2.SelectedIndex = 9;

                    tabControl2.SelectTab(tabPage9);
                    break;
            }
            

        }
        public bool nonPort; //Порт включен(выключен)
        public SerialPort newPort; //SerialPort
        public string portsName; //Имя порта
        public string GW1_2;
        public string versionPribor; //версия прибора
        public string wavelength1 = Convert.ToString(0);
        public string[] RDstring;
        public int countSA;
        public string GE5_1_0 = "";
        public int indata_zero;
        public bool ComPort, ComPodkl, StopSpectr, StopAgro, USE_KO;
        //public string this.Text = "";
        public string address_lab;
        public string name_lab;
        public bool OpenIzmer;
        public bool OpenIzmer1;
        public string Description;
        public string DateTime;
        public string Ispolnitel;
        public string direction;
        public string code;
        public SWF.TextBox[] textBox = new SWF.TextBox[20];
        public SWF.TextBox[] textBoxCO = new SWF.TextBox[20];
        public int NoCoIzmer;
        public string Veshestvo1;
        public string WidthCuvette;
        public string ND;
        public int Days;
        public string CountSeriya, CountInSeriya;
        public string BottomLine, TopLine;
        public double k0, k1, k2;
        public int NoCaIzm, NoCaIzm1, NoCaSer, NoCaSer1;
        public double start = 0.0, cancel = 0.0, interval, delay;
        public double[] scan_massSA;
        public double[] scan_mass;

        private void button12_Click(object sender, EventArgs e)
        {
            Calibrovka calibrovka = new Calibrovka(this);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            Izmerenie izmeren = new Izmerenie(this);
        }

        public double[] massWL;

        private void button15_Click(object sender, EventArgs e)
        {
            if (IzmerenieFR_Table.RowCount <= 26)
            {
                if (IzmerenieFR_Table.RowCount > 1)
                {
                    IzmerFr_count = IzmerenieFR_Table.RowCount - 1;
                    IzmerenieFR_Table.Rows.Add();
                    IzmerenieFR_Table.Rows[IzmerFr_count].Cells["N"].Value = IzmerFr_count + 1;
                    IzmerenieFR_Table.Rows[IzmerFr_count].Cells["Walve"].Value = GWNew.Text;
                    IzmerenieFR_Table.Rows[IzmerFr_count].Cells["KOne"].Value = "0.0";
                }
                else
                {
                    MessageBox.Show("Создайте новое Измерение");
                }
            }
            else
            {
                MessageBox.Show("Строк не более 26");
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (IzmerenieFR_Table.RowCount > 1)
            {
                if (IzmerenieFR_Table.RowCount > 2)
                {
                    if (IzmerenieFR_Table.CurrentCell.RowIndex != IzmerenieFR_Table.RowCount - 1)
                    {
                        IzmerenieFR_Table.Rows.RemoveAt(IzmerenieFR_Table.CurrentCell.RowIndex);
                        for (int i = 0; i < IzmerenieFR_Table.RowCount - 1; i++)
                        {
                            IzmerenieFR_Table.Rows[i].Cells[0].Value = i + 1;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Удаление запрещено!");
                    }
                }

                else
                {
                    MessageBox.Show("Количество образцов не может быть меньше 1 !");
                }


            }
            else
            {
                MessageBox.Show("Таблица не содержит строк!");
            }
        }

        public double[] massGE;
        public string[][,] countScan;
        public int countscan = 0;

        private void button11_Click(object sender, EventArgs e)
        {
            switch (selet_rezim)
            {
                case 2:
                 /* if (tabControl2.SelectedIndex == 0)
                    {
                        if (Table1.CurrentCell.ColumnIndex >= 3 && Table1.CurrentCell.ReadOnly != true)
                        {

                            if (Table1.CurrentCell.Value != "" && Table1.CurrentCell.Value != null)
                            {
                                CellOpt = Convert.ToDouble(Table1.CurrentCell.Value.ToString());
                                CelloptCopy = CellOpt;
                            }

                            ZapicInTable1();

                        }
                    }
                    else
                    {
                        if (Table2.CurrentCell.ColumnIndex >= 2 && Table2.CurrentCell.ReadOnly != true)
                        {
                            if (Table2.CurrentCell.Value != "" && Table2.CurrentCell.Value != null)
                            {
                                CellOpt = Convert.ToDouble(Table2.CurrentCell.Value.ToString());
                                CelloptCopy = CellOpt;
                            }

                            ZapicInTable2();
                        }
                    }*/
                    break;
                case 1:
                    if (IzmerenieFR_Table.CurrentCell.ColumnIndex == 5)
                    {
                        IzmerenieFR_Table_Zapic();
                    }
                    else
                    {
                        if (IzmerenieFR_Table.CurrentCell.ReadOnly == true)
                        {
                            MessageBox.Show("Изменения запрещены");
                        }
                    }
                    break;
                case 6:
                    if (tabControl2.SelectedIndex == 0)
                    {
                        if (Table1.CurrentCell.ColumnIndex >= 3 && Table1.CurrentCell.ReadOnly != true)
                        {

                            if (Table1.CurrentCell.Value != "" && Table1.CurrentCell.Value != null)
                            {
                              //CellOpt = Convert.ToDouble(Table1.CurrentCell.Value.ToString());
                            }

                          //ZapicInTable1();

                        }
                    }
                    else
                    {
                        StopAgro = true;
                    }
                    break;
                case 5:
                    StopSpectr = true;
                    break;
                case 4:
                    if (timer2.Enabled == true)
                    {
                        timer2.Enabled = true;
                        timer2.Stop();
                        MinMax();
                        countscan = 0;
                        //  TableKinetica1.Rows.Add();
                        button14.Enabled = true;
                        button11.Enabled = false;
                        label27.Visible = true;
                        label28.Visible = false;
                        Podskazka.Text = "Сохраните измерение";
                        label33.Visible = false;
                        button6.Enabled = true;
                        button5.Enabled = true;
                        button3.Enabled = true;
                        button7.Enabled = true;
                        button12.Enabled = true;
                        button8.Enabled = true;

                    }

                    break;
                case 3:
                    StopSpectr = true;
                    break;
                case 9:
                    StopSpectr = true;
                    break;
            }
            button1.Enabled = true;
        }

        public double timeLeft;

        public void IzmerenieFR_Table_Zapic()
        {
            InputBox _InputBox = new InputBox(this);
            _InputBox.ShowDialog();

            IzmerenieFR_Table.CurrentCell.Value = string.Format("{0:0.0}", CellOpt);
            CellOpt = 0;
            IzmerenieFR_Table.Rows[IzmerenieFR_Table.CurrentRow.Index].Cells[6].Value = string.Format("{0:0.0000}",
                Convert.ToDouble(IzmerenieFR_Table.Rows[IzmerenieFR_Table.CurrentRow.Index].Cells[3].Value)
                * Convert.ToDouble(IzmerenieFR_Table.Rows[IzmerenieFR_Table.CurrentRow.Index].Cells[5].Value));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Conection conection = new Conection(this);
            button13.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PortClose portclose = new PortClose(this);
        }

        private void button17_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView5.ColumnCount - 2; i++)
            {
                if (dataGridView5.Columns["Abs " + i].HeaderText == "Abs " + textBoxCO[i].Text + " нм")
                {
                    dataGridView5.Columns["Abs " + i].HeaderText = "%T " + textBoxCO[i].Text + " нм";
                }
                else
                {
                    dataGridView5.Columns["Abs " + i].HeaderText = "Abs " + textBoxCO[i].Text + " нм";
                }
            }
            for (int j = 0; j < dataGridView5.Rows.Count - 1; j++)
            {
                for (int i = 0; i < dataGridView5.ColumnCount - 2; i++)
                {
                    if (dataGridView5.Columns["Abs " + i].HeaderText == "Abs " + textBoxCO[i].Text + " нм")
                    {
                        if (massGEMultiAbs[j][i].ToString() != null)
                        {
                            dataGridView5.Rows[j].Cells["Abs " + i].Value = string.Format("{0:0.0000}", massGEMultiAbs[j][i]);
                        }
                        else
                        {
                            dataGridView5.Rows[j].Cells["Abs " + i].Value = null;
                        }
                    }
                    else
                    {
                        if (massGEMultiT[j][i].ToString() != null)
                        {
                            dataGridView5.Rows[j].Cells["Abs " + i].Value = string.Format("{0:0.00}", massGEMultiT[j][i]);
                        }
                        else
                        {
                            dataGridView5.Rows[j].Cells["Abs " + i].Value = null;
                        }
                    }
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (ComPodkl == true)
            {
                WalveNew();
            }
            else
            {
                MessageBox.Show("Подключитесь к прибору!");
            }
        }
        public void WalveNew()
        {
            NewWalve _NewWalve = new NewWalve(this);
            _NewWalve.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            PriborInformation informationPribor = new PriborInformation(this);
            informationPribor.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            CreateDimension createDemension = new CreateDimension(this);
        }

        private void Ecoview_Load(object sender, EventArgs e)
        {

        }
        public void TableKinetica(object sender, EventArgs e)
        {
            label56.Text = string.Format("{0:0.0}", timeLeft);
            timeLeft = timeLeft - Convert.ToDouble(interval);
            Application.DoEvents();
            //  MessageBox.Show("Интервал: " + interval*1000);
            TableKinetica1.Rows.Add();
            Array.Resize<double>(ref massWL, massWL.Length + 1);
            Array.Resize<double>(ref massGE, massGE.Length + 1);

            string GE5Izmer = "";
            string GE5_1_1 = "";
            while (GE5Izmer == "")
            {
                // SW_Scan();
                GE5Izmer = "";
                GE5_1_1 = "";
                newPort.Write("SA " + countSA + "\r");
                string indata = newPort.ReadExisting();
                string indata_0;
                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {
                        indata_bool = false;
                    }

                    else
                    {
                        indata = newPort.ReadExisting();
                    }
                }
                newPort.Write("GE 1\r");
               
                indata_0 = "";
                for (int i = 0; i <= 5000000; i++)
                {
                    indata = newPort.ReadExisting();
                    if (indata_0.Contains("\r>"))
                    {
                        break;
                    }
                    indata_0 += indata;
                }

               
                indata_bool = true;
              
                Regex regex = new Regex(@"\W");
                Regex regex1 = new Regex(@"\D");
                GE5Izmer = regex.Replace(indata_0, "");
                GE5Izmer = regex1.Replace(GE5Izmer, "");
            }
           
            GEText.Text = GE5Izmer;
            double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
            double OptPlot1 = 0;
           
            OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) / (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));
            double OptPlot1_1 = OptPlot1;
            Application.DoEvents();
            massWL[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
            TableKinetica1.Rows[countscan].Cells[0].Value = string.Format("{0:0.0}", Convert.ToDouble(interval) * countscan);
            if (TableKinetica1.Columns[1].HeaderText == "Abs")
            {
                TableKinetica1.Rows[countscan].Cells[1].Value = string.Format("{0:0.0000}", OptPlot1_1);
                massGE[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);
                TableKinetica1.Rows[countscan].Cells[2].Value =
                    string.Format("{0:0.0}",
                    ((Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) * 100));
            }
            else
            {
                TableKinetica1.Rows[countscan].Cells[2].Value = string.Format("{0:0.0000}", OptPlot1_1);
                TableKinetica1.Rows[countscan].Cells[1].Value =
                    string.Format("{0:0.0}",
                    ((Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) * 100));
              
                massGE[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);

            }
            massWL[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
            massGE[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);
            Array.Sort(massWL);
            Array.Sort(massGE);
            double x1 = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
            double y1 = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);
            /*ScanChart.Series[0].Points.AddXY(x1, y1);
            ScanChart.Series[0].ChartType = SeriesChartType.Point;
            ScanChart.ChartAreas[0].AxisY.Crossing = 0;
            ScanChart.ChartAreas[0].AxisX.Crossing = 0;*/

            chart3.Series[countButtonClick].Points.AddXY(x1, y1);
            chart3.Series[countButtonClick].ChartType = SeriesChartType.Line;
            if (TableKinetica1.Rows[countscan].Cells[1].Value != null && TableKinetica1.Rows[countscan].Cells[2].Value != null)
            {
                chart3.ChartAreas[0].AxisX.Minimum = cancel;
                chart3.ChartAreas[0].AxisX.Maximum = start;

            }
            chart3.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart3.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart3.ChartAreas[0].AxisX.Title = TableKinetica1.Columns[0].HeaderText;
            chart3.ChartAreas[0].AxisY.Title = TableKinetica1.Columns[1].HeaderText;
            countscan++;
            //del = GoodMorning;


            if (timeLeft == 0.0)
            {
                timer2.Enabled = false;
                timer2.Stop();
                MinMax();
                button14.Enabled = true;
                button11.Enabled = false;
                label33.Visible = false;
                label27.Visible = true;
                Podskazka.Text = "Сохраните градуировку!";
                button1.Enabled = true;
                label56.Text = "00";
                countscan = 0;

                button6.Enabled = true;
                button5.Enabled = true;
                button3.Enabled = true;
                button7.Enabled = true;
                button12.Enabled = true;
                button8.Enabled = true;

            }

        }
        public void MinMax()
        {
            double max = 0.0;
            double min = 0.0;
            countscan = 0;
            for (int i = 0; i < TableKinetica1.Rows.Count; i++)
            {
                double x1 = 0;
                double y1 = 0;
                if (i == 0)
                {
                    if (Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) > Convert.ToDouble(TableKinetica1.Rows[i + 1].Cells[1].Value))
                    {
                        max = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                        dataGridView3.Rows.Add(TableKinetica1.Rows[i].Cells[0].Value, TableKinetica1.Rows[i].Cells[1].Value, TableKinetica1.Rows[i].Cells[2].Value);
                        min = max;
                        x1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[0].Value);
                        y1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                        chart3.Series[0].Points.AddXY(x1, y1);
                        chart3.Series[0].Points[countscan].Label = Convert.ToString(x1);
                        chart3.Series[0].Points[countscan].Color = System.Drawing.Color.DarkViolet;
                        chart3.Series[0].ChartType = SeriesChartType.Point;
                        countscan++;
                    }
                    else
                    {
                        min = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                        dataGridView4.Rows.Add(TableKinetica1.Rows[i].Cells[0].Value, TableKinetica1.Rows[i].Cells[1].Value, TableKinetica1.Rows[i].Cells[2].Value);
                        max = min;
                        x1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[0].Value);
                        y1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                        chart3.Series[0].Points.AddXY(x1, y1);
                        chart3.Series[0].Points[countscan].Label = Convert.ToString(x1);
                        chart3.Series[0].Points[countscan].Color = System.Drawing.Color.DarkOrchid;
                        chart3.Series[0].ChartType = SeriesChartType.Point;
                        countscan++;
                    }

                }
                else {
                    if (i + 1 != TableKinetica1.Rows.Count)
                    {
                        if (Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) > Convert.ToDouble(TableKinetica1.Rows[i - 1].Cells[1].Value)
                            &&
                            Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) >= Convert.ToDouble(TableKinetica1.Rows[i + 1].Cells[1].Value)

                            )
                        {
                            max = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                            min = max;
                            dataGridView3.Rows.Add(TableKinetica1.Rows[i].Cells[0].Value, TableKinetica1.Rows[i].Cells[1].Value, TableKinetica1.Rows[i].Cells[2].Value);
                            x1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[0].Value);
                            y1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                            chart3.Series[0].Points.AddXY(x1, y1);
                            chart3.Series[0].Points[countscan].Label = Convert.ToString(x1);
                            chart3.Series[0].Points[countscan].Color = System.Drawing.Color.DarkViolet;
                            chart3.Series[0].ChartType = SeriesChartType.Point;
                            countscan++;
                        }
                        if ((Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) < Convert.ToDouble(TableKinetica1.Rows[i - 1].Cells[1].Value)
                            &&
                            Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value) <= Convert.ToDouble(TableKinetica1.Rows[i + 1].Cells[1].Value))

                            )
                        {
                            min = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                            dataGridView4.Rows.Add(TableKinetica1.Rows[i].Cells[0].Value, TableKinetica1.Rows[i].Cells[1].Value, TableKinetica1.Rows[i].Cells[2].Value);
                            max = min;
                            x1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[0].Value);
                            y1 = Convert.ToDouble(TableKinetica1.Rows[i].Cells[1].Value);
                            chart3.Series[0].Points.AddXY(x1, y1);
                            chart3.Series[0].Points[countscan].Label = Convert.ToString(x1);
                            chart3.Series[0].Points[countscan].Color = System.Drawing.Color.Teal;
                            chart3.Series[0].ChartType = SeriesChartType.Point;
                            countscan++;
                        }
                    }
                }

            }


        }
        public void IzmerenieFr_izmer()
        {
            int startIndexCell = 3;
            int endIndexCell = 6;
            int rowIndex = IzmerenieFR_Table.CurrentRow.Index;

            bool doNotWrite = false;
            string SWAnalis = WL_grad1;
            string GE5Izmer = "";
            string GE5_1_1 = "";
            newPort.Write("SA " + countSA + "\r");
            string indata = newPort.ReadExisting();
            string indata_0;
            bool indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else
                {
                    indata = newPort.ReadExisting();

                }
            }


            // double[] GEIZMERmass = new double[10];
            double GEIZMERmass = 0;
            for (int j = 0; j < 10; j++)
            {
                newPort.Write("GE 1\r");
                indata_0 = "";
                for (int i = 0; i <= 5000000; i++)
                {
                    indata = newPort.ReadExisting();
                    if (indata_0.Contains("\r>"))
                    {
                        break;
                    }
                    indata_0 += indata;
                }

                indata_bool = true;
                GE5Izmer = "";
                Regex regex = new Regex(@"\W");
                Regex regex1 = new Regex(@"\D");
                GE5Izmer = regex.Replace(indata_0, "");
                GE5Izmer = regex1.Replace(GE5Izmer, "");
                // GEIZMERmass[j] = Convert.ToDouble(GE5Izmer);
                GEIZMERmass += Convert.ToDouble(GE5Izmer);
            }

            GEIZMERmass = GEIZMERmass / 10;
            GE5Izmer = "";
            GE5Izmer = Convert.ToString(GEIZMERmass);
            //GE5Izmer = Convert.ToString(GEIZMERmass.Max());


            GEText.Text = GE5Izmer;

            double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
            double OptPlot1 = 0;

            OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) /
                (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));


            IzmerenieFR_Table.Rows[rowIndex].Cells[1].Value = "Образец " + (rowIndex + 1);
            double OptPlot1_1 = OptPlot1;
            IzmerenieFR_Table.Rows[rowIndex].Cells[2].Value = string.Format("{0:0.0}", GWNew.Text);
            if (IzmerenieFR_Table.CurrentCell.ColumnIndex != 5)
            {
                if ((IzmerenieFR_Table.CurrentCell.ReadOnly != true && rowIndex != IzmerenieFR_Table.Rows.Count - 1) || IzmerenieFR_Table.CurrentCell.ColumnIndex == 3)
                {
                    IzmerenieFR_Table.Rows[rowIndex].Cells[3].Value = string.Format("{0:0.0000}", OptPlot1_1);
                    string k1 = Convert.ToString(IzmerenieFR_Table.Rows[rowIndex].Cells[5].Value);
                    k1 = k1.Replace(".", ",");
                    IzmerenieFR_Table.Rows[rowIndex].Cells[4].Value = string.Format("{0:0.00}",
                         (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                        ((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA]))) * 100);


                    IzmerenieFR_Table.Rows[rowIndex].Cells[6].Value = string.Format("{0:0.0000}", (OptPlot1_1 * Convert.ToDouble(k1)));
                    int curentIndex = IzmerenieFR_Table.CurrentCell.ColumnIndex;
                    if (curentIndex != IzmerenieFR_Table.ColumnCount - 1 || rowIndex != IzmerenieFR_Table.Rows.Count - 1)
                    {
                        if (rowIndex != IzmerenieFR_Table.Rows.Count - 2)
                        {
                            IzmerenieFR_Table.CurrentCell = this.IzmerenieFR_Table[curentIndex, rowIndex + 1];
                        }
                        else
                        {
                            MessageBox.Show("Измерения были проведены!");
                        }
                    }

                }
                else
                {
                    MessageBox.Show("Запись запрещена!");
                }
            }
            else
            {
                MessageBox.Show("Производить измерения в данную ячейку запрещено! Только ручное изменение!");
            }
        }
        public void TimerTick1(object sender, EventArgs e)
        {
            label53.Text = Convert.ToString(delay);
            delay--;
            if (delay < 0.0)
            {
                //  label33.Visible = false;
                timer1.Stop();
                timer1.Enabled = false;
                timeLeft = Convert.ToInt32(start);
                timer2.Start();
                TableKinetica1.Rows.Clear();
                TableKinetica(sender, e);
                TableKinetica1.Rows.Clear();
                timer2.Enabled = true;
                button11.Enabled = true;
                countscan = 0;
            }

        }
        public void ChartGraf()
        {


            if (TableKinetica1.Columns[1].HeaderText == "Abs")
            {
                //Array.Sort(massGE);
                chart3.ChartAreas[0].AxisY.Minimum = 0;
                chart3.ChartAreas[0].AxisY.Maximum = 3;
                chart3.ChartAreas[0].AxisX.Minimum = 0;
                chart3.ChartAreas[0].AxisX.Maximum = start;
                dataGridView1.Columns[1].HeaderText = "Abs";
                dataGridView1.Columns[2].HeaderText = "%T";
                dataGridView2.Columns[1].HeaderText = "Abs";
                dataGridView2.Columns[2].HeaderText = "%T";
                TableKinetica1.Columns[2].HeaderText = "%T";
            }
            else
            {
                //Array.Sort(massGE);
                chart3.ChartAreas[0].AxisY.Minimum = 0;
                chart3.ChartAreas[0].AxisY.Maximum = 125;
                chart3.ChartAreas[0].AxisX.Minimum = 0;
                chart3.ChartAreas[0].AxisX.Maximum = start;
                TableKinetica1.Columns[2].HeaderText = "Abs";
                dataGridView1.Columns[1].HeaderText = "%T";
                dataGridView1.Columns[2].HeaderText = "Abs";
                dataGridView2.Columns[1].HeaderText = "%T";
                dataGridView2.Columns[2].HeaderText = "Abs";
            }

        }
        public void TableKinetica()
        {

            label56.Text = string.Format("{0:0.0}", timeLeft);
            timeLeft = timeLeft - Convert.ToDouble(interval);
            Application.DoEvents();
            //  MessageBox.Show("Интервал: " + interval*1000);
            TableKinetica1.Rows.Add();
            Array.Resize<double>(ref massWL, massWL.Length + 1);
            Array.Resize<double>(ref massGE, massGE.Length + 1);

            string GE5Izmer = "";
            string GE5_1_1 = "";
            while (GE5Izmer == "")
            {
                // SW_Scan();
                GE5Izmer = "";
                GE5_1_1 = "";
                newPort.Write("SA " + countSA + "\r");
                string indata = newPort.ReadExisting();
                string indata_0;
                bool indata_bool = true;
                while (indata_bool == true)
                {
                    if (indata.Contains(">"))
                    {
                        indata_bool = false;
                    }

                    else
                    {
                        indata = newPort.ReadExisting();
                    }
                }
                newPort.Write("GE 1\r");
    
                indata_0 = "";
                for (int i = 0; i <= 5000000; i++)
                {
                    indata = newPort.ReadExisting();
                    if (indata_0.Contains("\r>"))
                    {
                        break;
                    }
                    indata_0 += indata;
                }

                indata_bool = true;

                Regex regex = new Regex(@"\W");
                Regex regex1 = new Regex(@"\D");
                GE5Izmer = regex.Replace(indata_0, "");
                GE5Izmer = regex1.Replace(GE5Izmer, "");
            }

            GEText.Text = GE5Izmer;
            double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
            double OptPlot1 = 0;
         
            OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) / (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));
            double OptPlot1_1 = OptPlot1;
            Application.DoEvents();
            massWL[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
            TableKinetica1.Rows[countscan].Cells[0].Value = string.Format("{0:0.0}", Convert.ToDouble(interval) * countscan);
            if (TableKinetica1.Columns[1].HeaderText == "Abs")
            {
                TableKinetica1.Rows[countscan].Cells[1].Value = string.Format("{0:0.0000}", OptPlot1_1);
                massGE[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);
                TableKinetica1.Rows[countscan].Cells[2].Value =
                    string.Format("{0:0.0}",
                    ((Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) * 100));
            }
            else
            {
                TableKinetica1.Rows[countscan].Cells[2].Value = string.Format("{0:0.0000}", OptPlot1_1);
                TableKinetica1.Rows[countscan].Cells[1].Value =
                    string.Format("{0:0.0}",
                    ((Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) * 100));
                /** listBox1.Items.Add(GE5Izmer);*/
                massGE[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);

            }
            massWL[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
            massGE[countscan] = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);
            Array.Sort(massWL);
            Array.Sort(massGE);
            double x1 = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[0].Value);
            double y1 = Convert.ToDouble(TableKinetica1.Rows[countscan].Cells[1].Value);


            chart3.Series[countButtonClick].Points.AddXY(x1, y1);
            chart3.Series[countButtonClick].ChartType = SeriesChartType.Line;
            if (TableKinetica1.Rows[countscan].Cells[1].Value != null && TableKinetica1.Rows[countscan].Cells[2].Value != null)
            {
                chart3.ChartAreas[0].AxisX.Minimum = cancel;
                chart3.ChartAreas[0].AxisX.Maximum = start;

            }
            chart3.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart3.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
            chart3.ChartAreas[0].AxisX.Title = TableKinetica1.Columns[0].HeaderText;
            chart3.ChartAreas[0].AxisY.Title = TableKinetica1.Columns[1].HeaderText;
            countscan++;
            //del = GoodMorning;


            if (timeLeft == 0.0)
            {
                timer2.Enabled = false;
                timer2.Stop();
                MinMax();
                button14.Enabled = true;
                button11.Enabled = false;
                label33.Visible = false;
                label27.Visible = true;
                Podskazka.Text = "Сохраните градуировку!";
                button1.Enabled = true;
                label56.Text = "00";
                countscan = 0;

                button6.Enabled = true;
                button5.Enabled = true;
                button3.Enabled = true;
                button7.Enabled = true;
                button12.Enabled = true;
                button8.Enabled = true;

            }

        }
        public void Play_Ecxel()
        {

            try
            {
                string GE5Izmer = "";

                int cell = Convert.ToInt32(workSheet.Application.ActiveCell.Column);
                int row = Convert.ToInt32(workSheet.Application.ActiveCell.Row);

                newPort.Write("SA " + countSA + "\r");
                string indata = newPort.ReadExisting();

                string indata_0;
                bool indata_bool = true;
                while (indata_bool == true)
                {

                    if (indata.Contains(">"))
                    {

                        indata_bool = false;

                    }

                    else
                    {
                        indata = newPort.ReadExisting();

                    }
                }
                indata_bool = true;
                newPort.Write("GE 1\r");
                indata_0 = "";

                for (int i = 0; i <= 5000000; i++)
                {
                    indata = newPort.ReadExisting();
                    if (indata_0.Contains("\r>"))
                    {
                        break;
                    }
                    indata_0 += indata;
                }
                GE5Izmer = "";
                Regex regex = new Regex(@"\W");
                Regex regex1 = new Regex(@"\D");
                GE5Izmer = regex.Replace(indata_0, "");
                GE5Izmer = regex1.Replace(GE5Izmer, "");

                double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
                double OptPlot1 = 0;

                OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));
                double OptPlot1_1 = OptPlot1;

                workSheet.Cells[row, cell] = string.Format("{0:0.0000}", OptPlot1_1);
                // timePlay++;
            }
            catch
            {
                MessageBox.Show("Вы закрыли файл или файл недоступен для записи!");
            }



        }
        public void SW_MultiScan()
        {
            SW_MultiScan sw_multiscan = new SW_MultiScan(this);
        }
        public void TableMultiScan()
        {
            countscan = 0;
            while ((countscan != dataGridView5.ColumnCount - 2) && (StopSpectr != true))
            {
                Application.DoEvents();
                string GE5Izmer = "";
                string GE5_1_1 = "";
                while (GE5Izmer == "")
                {
                    SW_MultiScan();
                    GE5Izmer = "";
                    GE5_1_1 = "";
                    newPort.Write("SA " + scan_massSA[countscan] + "\r");
                    string indata = newPort.ReadExisting();
                    string indata_0;
                    bool indata_bool = true;
                    while (indata_bool == true)
                    {
                        if (indata.Contains(">"))
                        {
                            indata_bool = false;
                        }

                        else
                        {
                            indata = newPort.ReadExisting();
                        }
                    }

                    newPort.Write("GE 1\r");
                   
                    indata_0 = "";
                    for (int i = 0; i <= 5000000; i++)
                    {
                        indata = newPort.ReadExisting();
                        if (indata_0.Contains("\r>"))
                        {
                            break;
                        }
                        indata_0 += indata;
                    }

                    //  indata_0 = "";
                    indata_bool = true;
                   
                    Regex regex = new Regex(@"\W");
                    Regex regex1 = new Regex(@"\D");
                    GE5Izmer = regex.Replace(indata_0, "");
                    GE5Izmer = regex1.Replace(GE5Izmer, "");
                }
                //MessageBox.Show("Измерение");
                GEText.Text = GE5Izmer;
                double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(scan_mass[countscan]) * 100;
                double OptPlot1 = 0;

                OptPlot1 = Math.Log10((Convert.ToDouble(scan_mass[countscan]) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));
                double OptPlot1_1 = OptPlot1;
                Application.DoEvents();
                dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells["Abs " + countscan].Value = string.Format("{0:0.0000}", OptPlot1_1);
                massGEMultiAbs[dataGridView5.Rows.Count - 2][countscan] =
                    Convert.ToDouble(dataGridView5.Rows[dataGridView5.Rows.Count - 2].Cells["Abs " + countscan].Value);

                massGEMultiT[dataGridView5.Rows.Count - 2][countscan] = (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])) /
                    (Convert.ToDouble(scan_mass[countscan]) - Convert.ToDouble(RDstring[countSA])) * 100;
                countscan++;
                Application.DoEvents();
            }
            button14.Enabled = true;
            button11.Enabled = false;
            Application.DoEvents();
            if (StopSpectr == true)
            {
                MessageBox.Show("Измерение было прервано!");
            }
        }
        int cellnull;
        bool Izmerenie1;
        double Asred1;
        public void Graduirovka()
        {
            double sum = 0.0;
            int startIndexCell = 2;
            int endIndexCell = startIndexCell + NoCaIzm;
            int rowIndex = Table1.CurrentRow.Index;

            bool doNotWrite = false;
            string SWAnalis = WL_grad1;
            string GE5Izmer = "";
            string GE5_1_1 = "";
            newPort.Write("SA " + countSA + "\r");
            string indata = newPort.ReadExisting();
            string indata_0;
            bool indata_bool = true;
            while (indata_bool == true)
            {

                if (indata.Contains(">"))
                {

                    indata_bool = false;

                }

                else
                {
                    indata = newPort.ReadExisting();

                }
            }

            newPort.Write("GE 1\r");
            indata_0 = "";
            for (int i = 0; i <= 5000000; i++)
            {
                indata = newPort.ReadExisting();
                if (indata_0.Contains("\r>"))
                {
                    break;
                }
                indata_0 += indata;
            }


            indata_bool = true;

            Regex regex = new Regex(@"\W");
            Regex regex1 = new Regex(@"\D");
            GE5Izmer = regex.Replace(indata_0, "");
            GE5Izmer = regex1.Replace(GE5Izmer, "");
            GEText.Text = GE5Izmer;

            double Aser = Convert.ToDouble(GE5Izmer) / Convert.ToDouble(GE5_1_0) * 100;
            double OptPlot1 = 0;

            OptPlot1 = Math.Log10((Convert.ToDouble(GE5_1_0) - Convert.ToDouble(RDstring[countSA])) / (Convert.ToDouble(GE5Izmer) - Convert.ToDouble(RDstring[countSA])));



            double OptPlot1_1 = OptPlot1;
            if (Table1.CurrentCell.ReadOnly != true)
            {
                Table1.CurrentCell.Value = string.Format("{0:0.0000}", OptPlot1_1);
                Table1.CurrentCell.Style.BackColor = System.Drawing.Color.White;

                int curentIndex = Table1.CurrentCell.ColumnIndex;
                if (curentIndex != Table1.ColumnCount - 1 || rowIndex != Table1.Rows.Count - 2)
                {
                    if (rowIndex != Table1.Rows.Count - 2)
                    {
                        Table1.CurrentCell = this.Table1[curentIndex, rowIndex + 1];

                    }
                    else
                    {
                        Table1.CurrentCell = this.Table1[curentIndex + 1, 0];
                    }

                }

            }
            else
            {
                MessageBox.Show("Запись запрещена!");
            }
            GAText.Text = string.Format("{0:0.00}", Aser);
            for (int j = 0; j < Table1.Rows.Count - 1; j++)
            {
                {
                    for (int i = 3; i < Table1.Rows[j].Cells.Count; i++)
                    {
                        if (Table1.Rows[j].Cells[i].Value == null)
                        {
                            doNotWrite = true;

                            for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                            {
                                if (Table1.Rows[rowIndex].Cells[l].Value == null)
                                {
                                    cellnull++;
                                    // Table1.Rows[rowIndex].Cells[Table1.CurrentCell.ColumnIndex + 1].Selected = true;

                                }
                            }
                        }


                    }
                }
            }
            if (!doNotWrite)
            {
                if (NoCaSer == 1)
                {
                    radioButton1.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                    radioButton3.Enabled = false;
                    radioButton2.Enabled = false;
                }
                if (NoCaSer == 2)
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                    radioButton3.Enabled = false;
                }
                if (NoCaSer >= 3)
                {
                    radioButton1.Enabled = true;
                    radioButton2.Enabled = true;
                    radioButton3.Enabled = true;
                    radioButton4.Enabled = true;
                    radioButton5.Enabled = true;
                }
                sum = 0.0;
                

                for (int l = startIndexCell + NoCaIzm; l <= endIndexCell; ++l)
                {
                    if (Table1.Rows[rowIndex].Cells[l].Value == null)
                    {
                        cellnull++;
                    }

                    else
                    {
                        for (int j = 0; j < Table1.Rows.Count - 1; j++)
                        {

                            for (int i1 = startIndexCell + 1; i1 <= endIndexCell; ++i1)
                            {
                                sum += Convert.ToDouble(Table1.Rows[j].Cells[i1].Value);
                                Asred1 = sum / NoCaIzm;
                                // MessageBox.Show(Convert.ToString(Asred1));
                                Table1.Rows[j].Cells["Asred"].Value = string.Format("{0:0.0000}", Asred1);

                            }
                            sum = 0.0;
                        }
                    }
                    Izmerenie1 = true;
                }
                for (int m = 0; m < Table1.Rows.Count - 1; m++)
                {
                    for (int ml = 0; ml < Table1.Rows[m].Cells.Count; ml++)
                    {
                        if (Table1.Rows[m].Cells[ml].Value == null)
                        { doNotWrite = true; }
                    }
                }
                if (!doNotWrite)
                {
                    while (true)
                    {
                        int ml = Table1.Columns.Count - 1;//С какого столбца начать
                        if (Table1.Columns.Count == 3 + NoCaIzm)
                            break;
                        Table1.Columns.RemoveAt(ml);
                    }
                    functionAsred();

                }

            }
        }
        public int circle;
        public double XY, SUMMY2, SUMMX;
        public void functionAsred()
        {
            while (true)
            {
                int ml = Table1.Columns.Count - 1;//С какого столбца начать
                if (Table1.Columns.Count == 3 + NoCaIzm)
                    break;
                Table1.Columns.RemoveAt(ml);
            }
            groupBox3.Enabled = true;
            groupBox2.Enabled = true;

            if (radioButton1.Checked == true)
            {
                Lineinaya0 lineinaya0 = new Lineinaya0(this);
               
            }
            else
            {
                if (radioButton2.Checked == true)
                {

                   // lineinaya();
                }
                else
                {
                   // kvadratichnaya();
                }
            }
            Podskazka.Text = "Сохраните градуировку!";
            label27.Visible = true;
            label24.Visible = false;
            label25.Visible = false;
            label26.Visible = false;
            label28.Visible = false;
            label33.Visible = false;
            label14.Enabled = true;
            RR.Enabled = true;
            SKO.Enabled = true;
            label21.Enabled = true;
            label22.Enabled = true;
            AgroText0.Enabled = true;
            AgroText1.Enabled = true;
            AgroText2.Enabled = true;
        }
    }
}
