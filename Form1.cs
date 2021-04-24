using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Ports;
using DevComponents.Instrumentation;
using Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using ZedGraph;
using System.Threading;
namespace safrun
{
    
    public partial class Form1 : Form
    {
        GraphPane mypanehız = new GraphPane();
        GraphPane mypaneakım = new GraphPane();
        PointPairList listPointhız = new PointPairList();
        PointPairList listPointakım = new PointPairList();
        LineItem myCurvehız;
        LineItem myCurveakım;
        
        public Form1()
        {
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            this.hızgösterge.BackColor = System.Drawing.Color.Transparent;
            this.maxsicthermo.BackColor = System.Drawing.Color.Transparent;
            this.sıc1thermo.BackColor = System.Drawing.Color.Transparent;
            this.sıc2thermo.BackColor = System.Drawing.Color.Transparent;
            this.sıc3thermo.BackColor = System.Drawing.Color.Transparent;
            this.akımgösterge.BackColor = System.Drawing.Color.Transparent;
            this.topgergösterge.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox1.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox2.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox3.BackColor = System.Drawing.Color.Transparent;
            this.pictureBox4.BackColor = System.Drawing.Color.Transparent;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label7.BackColor = System.Drawing.Color.Transparent;
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label9.BackColor = System.Drawing.Color.Transparent;
            this.label10.BackColor = System.Drawing.Color.Transparent;
            this.label11.BackColor = System.Drawing.Color.Transparent;
            this.label13.BackColor = System.Drawing.Color.Transparent;
            this.label12.BackColor = System.Drawing.Color.Transparent;
            this.label14.BackColor = System.Drawing.Color.Transparent;
            this.label15.BackColor = System.Drawing.Color.Transparent;
            this.label16.BackColor = System.Drawing.Color.Transparent;
            this.label17.BackColor = System.Drawing.Color.Transparent;
            this.label18.BackColor = System.Drawing.Color.Transparent;
            this.label19.BackColor = System.Drawing.Color.Transparent;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label25.BackColor = System.Drawing.Color.Transparent;
            this.label26.BackColor = System.Drawing.Color.Transparent;
            this.label27.BackColor = System.Drawing.Color.Transparent;
            this.label28.BackColor = System.Drawing.Color.Transparent;
            this.label20.BackColor = System.Drawing.Color.Transparent;
            this.label29.BackColor = System.Drawing.Color.Transparent;
            this.label30.BackColor = System.Drawing.Color.Transparent;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label6.BackColor = System.Drawing.Color.Transparent;
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.BBS.BackColor = System.Drawing.Color.Transparent;
            this.BBS1.BackColor = System.Drawing.Color.Transparent;
            this.label31.BackColor = System.Drawing.Color.Transparent;
            this.label32.BackColor = System.Drawing.Color.Transparent;
            this.label33.BackColor = System.Drawing.Color.Transparent;
            this.label34.BackColor = System.Drawing.Color.Transparent;
            this.label35.BackColor = System.Drawing.Color.Transparent;
            this.label36.BackColor = System.Drawing.Color.Transparent;
            this.label37.BackColor = System.Drawing.Color.Transparent;
            this.label38.BackColor = System.Drawing.Color.Transparent;
            this.label39.BackColor = System.Drawing.Color.Transparent;
            this.label40.BackColor = System.Drawing.Color.Transparent;
            this.label41.BackColor = System.Drawing.Color.Transparent;
            this.label42.BackColor = System.Drawing.Color.Transparent;
            this.label43.BackColor = System.Drawing.Color.Transparent;
            this.label44.BackColor = System.Drawing.Color.Transparent;
            this.label45.BackColor = System.Drawing.Color.Transparent;
            this.label46.BackColor = System.Drawing.Color.Transparent;
            this.label47.BackColor = System.Drawing.Color.Transparent;
            this.label48.BackColor = System.Drawing.Color.Transparent;
            this.label49.BackColor = System.Drawing.Color.Transparent;
            this.label50.BackColor = System.Drawing.Color.Transparent;
            this.label51.BackColor = System.Drawing.Color.Transparent;
            this.label52.BackColor = System.Drawing.Color.Transparent;
            this.label53.BackColor = System.Drawing.Color.Transparent;
            this.label54.BackColor = System.Drawing.Color.Transparent;
            this.label55.BackColor = System.Drawing.Color.Transparent;
            this.label56.BackColor = System.Drawing.Color.Transparent;
            this.label57.BackColor = System.Drawing.Color.Transparent;
            this.label58.BackColor = System.Drawing.Color.Transparent;
            this.label59.BackColor = System.Drawing.Color.Transparent;
            this.label60.BackColor = System.Drawing.Color.Transparent;
            this.label61.BackColor = System.Drawing.Color.Transparent;
            this.label62.BackColor = System.Drawing.Color.Transparent;
            this.label63.BackColor = System.Drawing.Color.Transparent;
            this.label64.BackColor = System.Drawing.Color.Transparent;
            this.textBox8.BackColor = System.Drawing.Color.Transparent;
            this.textBox9.BackColor = System.Drawing.Color.Transparent;
            this.textBox11.BackColor = System.Drawing.Color.Transparent;
            this.textBox12.BackColor = System.Drawing.Color.Transparent;
            this.textBox13.BackColor = System.Drawing.Color.Transparent;
            this.textBox14.BackColor = System.Drawing.Color.Transparent;
            this.textBox15.BackColor = System.Drawing.Color.Transparent;
            this.textBox16.BackColor = System.Drawing.Color.Transparent;
            this.textBox17.BackColor = System.Drawing.Color.Transparent;
            this.textBox18.BackColor = System.Drawing.Color.Transparent;
            this.textBox19.BackColor = System.Drawing.Color.Transparent;
            this.textBox20.BackColor = System.Drawing.Color.Transparent;
            this.textBox21.BackColor = System.Drawing.Color.Transparent;
            this.textBox22.BackColor = System.Drawing.Color.Transparent;
            this.textBox23.BackColor = System.Drawing.Color.Transparent;
            this.textBox24.BackColor = System.Drawing.Color.Transparent;
            this.textBox25.BackColor = System.Drawing.Color.Transparent;
            this.textBox26.BackColor = System.Drawing.Color.Transparent;
            this.textBox27.BackColor = System.Drawing.Color.Transparent;


            string[] Portlar = SerialPort.GetPortNames();
            foreach(string port in Portlar)
            {
                comboBoxEx1.Items.Add(port);
                comboBoxEx1.SelectedIndex = 0;
            }
            comboBoxEx2.Items.Add("4800");
            comboBoxEx2.Items.Add("9600");
            comboBoxEx2.SelectedIndex = 1;
            label2.Text = "Bağlantı Kapalı";
            GrafikHazirla();
        }
        
        private void GrafikHazirla()
        {
            mypanehız = zedGraphControl1.GraphPane;
            mypaneakım = zedGraphControl2.GraphPane;
            mypanehız.Title.Text = "Hız - Zaman Grafiği";
            mypaneakım.Title.Text = "Akım - Zaman Grafiği";
            mypanehız.XAxis.Title.Text = " t (s)";
            mypaneakım.XAxis.Title.Text = " t (s)";
            mypanehız.YAxis.Title.Text = "Çıkış Hız(Km/h)";
            mypaneakım.YAxis.Title.Text = "Çıkış Akım(W/h)";
            mypanehız.YAxis.Scale.Min = 0;
            mypaneakım.YAxis.Scale.Min = 0;
            mypanehız.YAxis.Scale.Max = 200;
            mypaneakım.YAxis.Scale.Max = 200;
            myCurvehız = mypanehız.AddCurve(null, listPointhız, Color.Red, SymbolType.None);
            myCurveakım = mypaneakım.AddCurve(null, listPointakım, Color.Blue, SymbolType.None);
            myCurvehız.Line.Width = 3;
            myCurveakım.Line.Width = 3;
        }
        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            serialPort1.PortName = comboBoxEx1.SelectedItem.ToString();
            serialPort1.Open();

        }
        double zaman = 0;
        int degerler0,degerler1,degerler2,degerler3,degerler4,degerler5,degerler6,degerler7,degerler8,degerler9,degerler10,degerler11;
        int degerler12,degerler13,degerler14,degerler15,degerler16,degerler17,degerler18,degerler19,degerler20,degerler21,degerler22;
        int degerler24,degerler25,degerler26,degerler27,degerler23;
        private void timer1_Tick(object sender, EventArgs e)
        {
            
            string sonuc = serialPort1.ReadExisting();
            string[] degerler = sonuc.Split(',');
            
            
           
            try
            {
                label5.Text = degerler[0];
                label6.Text = degerler[1];
                label7.Text = degerler[2];
                label8.Text = degerler[3];
                label9.Text = degerler[3];
                label10.Text = degerler[5];
                label11.Text = degerler[6];
                textBox8.Text = degerler[7];
                textBox9.Text = degerler[8];
                textBox10.Text = degerler[9];
                textBox11.Text = degerler[10];
                textBox12.Text = degerler[11];
                textBox13.Text = degerler[12];
                textBox14.Text = degerler[13];
                textBox15.Text = degerler[14];
                textBox16.Text = degerler[15];
                textBox17.Text = degerler[16];
                textBox18.Text = degerler[17];
                textBox19.Text = degerler[18];
                textBox20.Text = degerler[19];
                textBox21.Text = degerler[20];
                textBox22.Text = degerler[21];
                textBox23.Text = degerler[22];
                textBox24.Text = degerler[23];
                textBox25.Text = degerler[24];
                textBox26.Text = degerler[25];
                textBox27.Text = degerler[26];
                zaman += 0.05;
                listPointhız.Add(new PointPair(zaman, degerler0));
                listPointakım.Add(new PointPair(zaman, degerler1));
                mypanehız.XAxis.Scale.Max = zaman;
                mypaneakım.XAxis.Scale.Max = zaman;
                mypanehız.AxisChange();
                mypaneakım.AxisChange();
                zedGraphControl1.Refresh();
                zedGraphControl2.Refresh();

                degerler0 = Convert.ToInt32(label5.Text);
                degerler1 = Convert.ToInt32(label6.Text);
                degerler2 = Convert.ToInt32(label7.Text);
                degerler3 = Convert.ToInt32(label8.Text);
                degerler4 = Convert.ToInt32(label9.Text);
                degerler5 = Convert.ToInt32(label10.Text);
                degerler6 = Convert.ToInt32(label11.Text);
                degerler7 = Convert.ToInt32(label39.Text);
                degerler8 = Convert.ToInt32(label48.Text);
                degerler9 = Convert.ToInt32(label49.Text);
                degerler10 = Convert.ToInt32(label50.Text);
                degerler11 = Convert.ToInt32(label51.Text);
                degerler12 = Convert.ToInt32(label51.Text);
                degerler13 = Convert.ToInt32(label52.Text);
                degerler14 = Convert.ToInt32(label53.Text);
                degerler15 = Convert.ToInt32(label54.Text);
                degerler16 = Convert.ToInt32(label55.Text);
                degerler17 = Convert.ToInt32(label56.Text);
                degerler18 = Convert.ToInt32(label57.Text);
                degerler19 = Convert.ToInt32(label58.Text);
                degerler20 = Convert.ToInt32(label59.Text);
                degerler21 = Convert.ToInt32(label60.Text);
                degerler22 = Convert.ToInt32(label61.Text);
                degerler23 = Convert.ToInt32(label62.Text);
                degerler24 = Convert.ToInt32(label63.Text);
                degerler25 = Convert.ToInt32(label64.Text);
                degerler26 = Convert.ToInt32(label65.Text); 
                  degerler7 = Convert.ToInt32(textBox8.Text);
                  degerler8 = Convert.ToInt32(textBox9.Text);
                  degerler9 = Convert.ToInt32(textBox10.Text);
                  degerler10 = Convert.ToInt32(textBox11.Text);
                  degerler11 = Convert.ToInt32(textBox12.Text);
                  degerler12 = Convert.ToInt32(textBox13.Text);
                  degerler13 = Convert.ToInt32(textBox14.Text);
                  degerler14 = Convert.ToInt32(textBox15.Text);
                  degerler15 = Convert.ToInt32(textBox16.Text);
                  degerler16 = Convert.ToInt32(textBox17.Text);
                  degerler17 = Convert.ToInt32(textBox18.Text);
                  degerler18 = Convert.ToInt32(textBox19.Text);
                  degerler19 = Convert.ToInt32(textBox20.Text);
                  degerler20 = Convert.ToInt32(textBox21.Text);
                  degerler21 = Convert.ToInt32(textBox22.Text);
                  degerler22 = Convert.ToInt32(textBox23.Text);
                  degerler23 = Convert.ToInt32(textBox24.Text);
                  degerler24 = Convert.ToInt32(textBox25.Text);
                  degerler25 = Convert.ToInt32(textBox26.Text);
                  degerler26 = Convert.ToInt32(textBox27.Text);

                 hızgösterge.SetPointerValue("Scale1", "Pointer1",degerler0);
                maxsicthermo.SetPointerValue("Scale1", "Pointer1", degerler1);
                sıc1thermo.SetPointerValue("Scale1", "Pointer1", degerler2);
                sıc2thermo.SetPointerValue("Scale1", "Pointer1", degerler3);
                sıc3thermo.SetPointerValue("Scale1","Pointer1",degerler4);
                akımgösterge.SetPointerValue("Scale1","Pointer1", degerler5);
                topgergösterge.SetPointerValue("Scale1", "Pointer1",degerler6);

                Pil1bar.Value = degerler7;
                Pil2bar.Value = degerler8;
                Pil3bar.Value = degerler9;
                Pil4bar.Value = degerler10;
                Pil5bar.Value = degerler11;
                Pil6bar.Value = degerler12;
                Pil7bar.Value = degerler13;
                Pil8bar.Value = degerler14;
                Pil9bar.Value = degerler15;
                Pil10bar.Value = degerler16;
                Pil11bar.Value = degerler17;
                Pil12bar.Value = degerler18;
                Pil13bar.Value = degerler19;
                Pil14bar.Value = degerler20;
                Pil15bar.Value = degerler21;
                Pil16bar.Value = degerler22;
                Pil17bar.Value = degerler23;
                Pil18bar.Value = degerler24;
                Pil19bar.Value = degerler25;
                Pil20bar.Value = degerler26;
                



                serialPort1.DiscardInBuffer();
                }
            catch (Exception ex)
                {
                
                MessageBox.Show(ex.Message);
                    timer1.Stop();

                }
            
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            timer1.Start();
            if (serialPort1.IsOpen == false)
            {
                if (comboBoxEx1.Text == "")
                {


                    return;
                    serialPort1.PortName = comboBoxEx1.Text;
                    serialPort1.BaudRate = Convert.ToInt16(comboBoxEx2.Text);
                    try
                    {
                        serialPort1.Open();
                        label2.Text = "Bağlantı Açık";
                    }
                    catch (Exception hata)
                    {
                        MessageBox.Show("Hata:" + hata.Message);

                    }
                }
                else
                {
                    label2.Text = "Bağlantı Kuruldu";
                }
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            timer1.Stop();
            serialPort1.DiscardInBuffer();
            if (serialPort1.IsOpen == true)
            {
                serialPort1.Close();
                label2.Text = "Bağlantı Kapalı";
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (serialPort1.IsOpen == true)
            {
                serialPort1.Close();
            }

        }
        int i;
        int j;
        private void button1_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            Workbook wb = excel.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet ws = (Worksheet)excel.ActiveSheet;
            ws.Cells[1, 1] = "Zaman";
            ws.Cells[1, 2] = "Hız";
            ws.Cells[1, 3] = "Akim";
            ws.Cells[1, 4] = "Maximum Sıcaklık";
            ws.Cells[1, 5] = "Toplam Gerilim";
            ws.Cells[1, 6] = "Sicaklik 1";
            ws.Cells[1, 7] = "Sicaklik 2";
            ws.Cells[1, 8] = "Sicaklik 3";
            ws.Cells[1, 9] = "Pil1";
            ws.Cells[1, 10] = "Pil2";
            ws.Cells[1, 11] = "Pil3";
            ws.Cells[1, 12] = "Pil4";
            ws.Cells[1, 13] = "Pil5";
            ws.Cells[1, 14] = "Pil6";
            ws.Cells[1, 15] = "Pil7";
            ws.Cells[1, 16] = "Pil8";
            ws.Cells[1, 17] = "Pil9";
            ws.Cells[1, 18] = "Pil10";
            ws.Cells[1, 19] = "Pil11";
            ws.Cells[1, 20] = "Pil12";
            ws.Cells[1, 21] = "Pil13";
            ws.Cells[1, 22] = "Pil14";
            ws.Cells[1, 23] = "Pil15";
            ws.Cells[1, 24] = "Pil16";
            ws.Cells[1, 25] = "Pil17";
            ws.Cells[1, 26] = "Pil18";
            ws.Cells[1, 27] = "Pil19";
            ws.Cells[1, 28] = "Pil20";
            try
            {
                for (i = 2; i < i + 1; i++)
                {
                    for (j = 1; j < 29; j++)
                    {
                        string sonuc = serialPort1.ReadExisting();
                        string[] degerler = sonuc.Split();
                        string saniye = DateTime.Now.Second.ToString();
                        string dakika = DateTime.Now.Minute.ToString();
                        string saat = DateTime.Now.Hour.ToString();

                        if (j == 1)
                        {
                            ws.Cells[i, j] = saat + ":" + dakika + ":" + saniye;

                        }
                        else if (j == 2)
                        {
                            ws.Cells[i, j] = degerler[0];

                        }
                        else if (j == 3)
                        {
                            ws.Cells[i, j] = degerler[1];

                        }
                        else if (j == 4)
                        {
                            ws.Cells[i, j] = degerler[2];

                        }
                        else if (j == 5)
                        {
                            ws.Cells[i, j] = degerler[3];

                        }
                        else if (j == 6)
                        {
                            ws.Cells[i, j] = degerler[4];

                        }
                        else if (j == 7)
                        {
                            ws.Cells[i, j] = degerler[5];

                        }
                        else if (j == 8)
                        {
                            ws.Cells[i, j] = degerler[6];

                        }

                        else if (j == 9)
                        {
                            ws.Cells[i, j] = degerler[7];

                        }
                        else if (j == 10)
                        {
                            ws.Cells[i, j] = degerler[8];

                        }
                        else if (j == 11)
                        {
                            ws.Cells[i, j] = degerler[9];

                        }
                        else if (j == 12)
                        {
                            ws.Cells[i, j] = degerler[10];

                        }
                        else if (j == 13)
                        {
                            ws.Cells[i, j] = degerler[11];

                        }
                        else if (j == 14)
                        {
                            ws.Cells[i, j] = degerler[12];

                        }
                        else if (j == 15)
                        {
                            ws.Cells[i, j] = degerler[13];

                        }
                        else if (j == 16)
                        {
                            ws.Cells[i, j] = degerler[14];

                        }
                        else if (j == 17)
                        {
                            ws.Cells[i, j] = degerler[15];

                        }
                        else if (j == 18)
                        {
                            ws.Cells[i, j] = degerler[16];

                        }
                        else if (j == 19)
                        {
                            ws.Cells[i, j] = degerler[17];

                        }
                        else if (j == 20)
                        {
                            ws.Cells[i, j] = degerler[18];

                        }
                        else if (j == 21)
                        {
                            ws.Cells[i, j] = degerler[19];

                        }
                        else if (j == 22)
                        {
                            ws.Cells[i, j] = degerler[20];

                        }
                        else if (j == 23)
                        {
                            ws.Cells[i, j] = degerler[21];

                        }
                        else if (j == 24)
                        {
                            ws.Cells[i, j] = degerler[22];

                        }
                        else if (j == 25)
                        {
                            ws.Cells[i, j] = degerler[23];

                        }
                        else if (j == 26)
                        {
                            ws.Cells[i, j] = degerler[24];

                        }
                        else if (j == 27)
                        {
                            ws.Cells[i, j] = degerler[25];

                        }
                        else if (j == 28)
                        {
                            ws.Cells[i, j] = degerler[26];

                        }


                    }
                }
            }
            catch (Exception)
            {

            }
        }

    }
    }

