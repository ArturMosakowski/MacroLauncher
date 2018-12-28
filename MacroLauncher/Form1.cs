using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;

namespace MacroLauncher
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage12;
            tabPage12.Focus();
            this.Size = new Size(950,550);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage1;
            tabPage1.Focus();
            button1.BackColor = Color.LightSteelBlue;

            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage2;
            tabPage2.Focus();
            button2.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage3;
            tabPage3.Focus();
            button4.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage4;
            tabPage4.Focus();
            button5.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage5;
            tabPage5.Focus();
            button6.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage6;
            tabPage6.Focus();
            button7.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage7;
            tabPage7.Focus();
            button8.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage8;
            tabPage8.Focus();
            button9.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage9;
            tabPage9.Focus();
            button10.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage10;
            tabPage10.Focus();
            button11.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button12.BackColor = Color.SteelBlue;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            tabControl1.SelectedTab = tabPage11;
            tabPage11.Focus();
            button12.BackColor = Color.LightSteelBlue;

            button1.BackColor = Color.SteelBlue;
            button2.BackColor = Color.SteelBlue;
            button4.BackColor = Color.SteelBlue;
            button5.BackColor = Color.SteelBlue;
            button6.BackColor = Color.SteelBlue;
            button7.BackColor = Color.SteelBlue;
            button8.BackColor = Color.SteelBlue;
            button9.BackColor = Color.SteelBlue;
            button10.BackColor = Color.SteelBlue;
            button11.BackColor = Color.SteelBlue;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox2.Lines.ElementAt(0);

                count1 = textBox1.Lines.Count();
                if (textBox1.Lines.Count() != textBox2.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox2.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox3.Lines.ElementAt(0);

                count1 = textBox1.Lines.Count();
                if (textBox1.Lines.Count() != textBox3.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox3.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox4.Lines.ElementAt(0);

                count1 = textBox1.Lines.Count();
                if (textBox1.Lines.Count() != textBox4.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox4.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox5.Lines.ElementAt(0);

                count1 = textBox1.Lines.Count();
                if (textBox1.Lines.Count() != textBox5.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox5.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox6.Lines.ElementAt(0);

                count1 = textBox1.Lines.Count();
                if (textBox1.Lines.Count() != textBox6.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox6.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox7.Lines.ElementAt(0);

                count1 = textBox1.Lines.Count();
                if (textBox1.Lines.Count() != textBox7.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox7.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            int count1 = 0;
            count1 = textBox1.Lines.Count();
            try
            {
                string firstrow1 = textBox2.Lines.ElementAt(0);
                if (textBox1.Lines.Count() != textBox2.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox2.Text += System.Environment.NewLine + firstrow1;
                    }
                }
            }
            catch { }
            try
            {
                string firstrow2 = textBox3.Lines.ElementAt(0);
                if (textBox1.Lines.Count() != textBox3.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox3.Text += System.Environment.NewLine + firstrow2;
                    }
                }
            }
            catch { }
            try
            {
                string firstrow3 = textBox4.Lines.ElementAt(0);
                if (textBox1.Lines.Count() != textBox4.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox4.Text += System.Environment.NewLine + firstrow3;
                    }
                }
            }
            catch { }
            try
            {
                string firstrow4 = textBox5.Lines.ElementAt(0);
                if (textBox1.Lines.Count() != textBox5.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox5.Text += System.Environment.NewLine + firstrow4;
                    }
                }
            }
            catch { }
            try
            {
                string firstrow5 = textBox6.Lines.ElementAt(0);
                if (textBox1.Lines.Count() != textBox6.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox6.Text += System.Environment.NewLine + firstrow5;
                    }
                }
            }
            catch { }
            try
            {
                string firstrow6 = textBox7.Lines.ElementAt(0);
                if (textBox1.Lines.Count() != textBox7.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox7.Text += System.Environment.NewLine + firstrow6;
                    }
                }
            }
            catch { }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            if (textBox1.Lines.Count() != textBox8.Lines.Count())
            {
                string warning = "Ilość pozycji w kolumnie Artykuł jest różna od ilości pozycji w kolumnie Wartość CS. Proszę uzupełnić dane.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }else
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();

                saveFileDialog1.Filter = "csv files (*.csv)|*.csv";
                saveFileDialog1.FileName = "TABELA DO IMPORTU CCS XX.csv";
                saveFileDialog1.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                saveFileDialog1.FilterIndex = 2;
                saveFileDialog1.RestoreDirectory = true;
                string text1 = "artykuł;region;sklep;format;typ ceny/marży;data pocz.;data koń.;wartość CS lub M%";
                string text2;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    using (Stream s = File.Open(saveFileDialog1.FileName, FileMode.Create))
                    using (StreamWriter writer = new StreamWriter(s))
                    {
                        writer.Write(text1);
                        writer.Write(System.Environment.NewLine);
                        for (int i = 0; i < textBox1.Lines.Count(); i++)
                        {
                            text2 = textBox1.Lines.ElementAt(i);
                            try
                            {
                                text2 += ";" + textBox2.Lines.ElementAt(i);
                            }
                            catch
                            {
                                text2 += ";";
                            }
                            try
                            {
                                text2 += ";" + textBox3.Lines.ElementAt(i);
                            }
                            catch
                            {
                                text2 += ";";
                            }
                            try
                            {
                                text2 += ";" + textBox4.Lines.ElementAt(i);
                            }
                            catch
                            {
                                text2 += ";";
                            }
                            try
                            {
                                text2 += ";" + textBox5.Lines.ElementAt(i);
                            }
                            catch
                            {
                                text2 += ";";
                            }
                            try
                            {
                                text2 += ";" + textBox6.Lines.ElementAt(i);
                            }
                            catch
                            {
                                text2 += ";";
                            }
                            try
                            {
                                text2 += ";" + textBox7.Lines.ElementAt(i);
                            }
                            catch
                            {
                                text2 += ";";
                            }
                            try
                            {
                                text2 += ";" + textBox8.Lines.ElementAt(i);
                            }
                            catch
                            {
                                text2 += ";";
                            }

                            writer.Write(text2);
                            writer.Write(System.Environment.NewLine);
                        }
                    }
                }
            };
        }

        private void button43_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox10.Lines.ElementAt(0);

                count1 = textBox11.Lines.Count();
                if (textBox11.Lines.Count() != textBox10.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox10.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
            try
            {
                int count2 = 0;
                string firstrow2 = textBox9.Lines.ElementAt(0);

                count2 = textBox11.Lines.Count();
                if (textBox11.Lines.Count() != textBox9.Lines.Count())
                {
                    for (int i = 1; i < count2; i++)
                    {
                        textBox9.Text += System.Environment.NewLine + firstrow2;
                    }
                }
            }
            catch { }
        }

        private void button42_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox10.Lines.ElementAt(0);

                count1 = textBox11.Lines.Count();
                if (textBox11.Lines.Count() != textBox10.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox10.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button41_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox9.Lines.ElementAt(0);

                count1 = textBox11.Lines.Count();
                if (textBox11.Lines.Count() != textBox9.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox9.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            textBox11.Clear();
            textBox10.Clear();
            textBox9.Clear();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            if ((textBox11.Lines.Count() == textBox10.Lines.Count()) && (textBox11.Lines.Count() == textBox9.Lines.Count()) && (textBox11.Lines.Count() > 0))
            {
                string[] art01 = new string[3000];
                string[] code01 = new string[3000];
                string[] crc01 = new string[3000];

                for (int i = 0; i < 3000; i++)
                {
                    art01[i] = "";
                    code01[i] = "";
                    crc01[i] = "";
                }
                for (int j = 0; j < textBox11.Lines.Count(); j++)
                {
                    art01[j] = textBox11.Lines.ElementAt(j);
                    code01[j] = textBox10.Lines.ElementAt(j);
                    crc01[j] = textBox9.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox11.Lines.Count(); k++)
                {
                    System.Windows.Forms.SendKeys.SendWait(art01[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait(code01[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait(crc01[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{F5}");
                    Thread.Sleep(300);
                }
                Application.Restart();
            }
            else
            {
                string warning = "Błąd wprowadzania danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button44_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox13.Lines.ElementAt(0);

                count1 = textBox14.Lines.Count();
                if (textBox13.Lines.Count() != textBox14.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox13.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            textBox14.Clear();
            textBox13.Clear();
        }

        private void button24_Click(object sender, EventArgs e)
        {
            if ((textBox14.Lines.Count() == textBox13.Lines.Count()) && (textBox14.Lines.Count() > 0))
            {
                string[] art02 = new string[3000];
                string[] code02 = new string[3000];

                for (int i = 0; i < 3000; i++)
                {
                    art02[i] = "";
                    code02[i] = "";
                }
                for (int j = 0; j < textBox14.Lines.Count(); j++)
                {
                    art02[j] = textBox14.Lines.ElementAt(j);
                    code02[j] = textBox13.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox14.Lines.Count(); k++)
                {
                    System.Windows.Forms.SendKeys.SendWait(art02[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait(code02[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{F5}");
                    Thread.Sleep(300);
                }
                Application.Restart();
            }
            else
            {
                string warning = "Błąd wprowadzania danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            textBox15.Clear();
            textBox12.Clear();
        }

        private void button26_Click(object sender, EventArgs e)
        {
            if ((textBox15.Lines.Count() == textBox12.Lines.Count()) && (textBox15.Lines.Count() > 0))
            {
                string[] art03 = new string[3000];
                string[] status03 = new string[3000];
                for (int i = 0; i < 3000; i++)
                {
                    art03[i] = "";
                    status03[i] = "";
                }
                for (int j = 0; j < textBox15.Lines.Count(); j++)
                {
                    art03[j] = textBox15.Lines.ElementAt(j);
                    status03[j] = textBox12.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox15.Lines.Count(); k++)
                {
                    System.Windows.Forms.SendKeys.SendWait(art03[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F6}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait(status03[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F1}");
                    Thread.Sleep(500);
                }
                Application.Restart();
            }
            else
            {
                string warning = "Ilość numerów artykułów nie jest zgodna z ilością statusów bądź nie wprowadzono żadnych danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            textBox19.Clear();
            textBox18.Clear();
            textBox17.Clear();
            textBox16.Clear();
        }

        private void button28_Click(object sender, EventArgs e)
        {
            if ((textBox19.Lines.Count() == textBox18.Lines.Count()) && (textBox19.Lines.Count() == textBox17.Lines.Count()) && (textBox19.Lines.Count() == textBox16.Lines.Count()) && (textBox19.Lines.Count() > 0))
            {
                string[] art04 = new string[3000];
                string[] cena04 = new string[3000];
                string[] data04 = new string[3000];
                string[] dostawca04 = new string[3000];

                for (int i = 0; i < 3000; i++)
                {
                    art04[i] = "";
                    cena04[i] = "";
                    data04[i] = "";
                    dostawca04[i] = "";
                }
                for (int j = 0; j < textBox19.Lines.Count(); j++)
                {
                    art04[j] = textBox19.Lines.ElementAt(j);
                    cena04[j] = textBox18.Lines.ElementAt(j);
                    data04[j] = textBox17.Lines.ElementAt(j);
                    dostawca04[j] = textBox16.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox19.Lines.Count(); k++)
                {
                    System.Windows.Forms.SendKeys.SendWait(art04[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait(dostawca04[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{F6}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait(data04[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait(cena04[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F1}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{F12}");
                    Thread.Sleep(500);
                }
                Application.Restart();
            }
            else
            {
                string warning = "Błąd wprowadzania danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            textBox22.Clear();
            textBox21.Clear();
            textBox20.Clear();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            if ((textBox22.Lines.Count() == textBox21.Lines.Count()) && (textBox22.Lines.Count() == textBox20.Lines.Count()) && (textBox22.Lines.Count() > 0))
            {
                string[] art05 = new string[3000];
                string[] status05 = new string[3000];
                string[] dostawca05 = new string[3000];

                for (int i = 0; i < 3000; i++)
                {
                    art05[i] = "";
                    status05[i] = "";
                    dostawca05[i] = "";
                }
                for (int j = 0; j < textBox22.Lines.Count(); j++)
                {
                    art05[j] = textBox22.Lines.ElementAt(j);
                    status05[j] = textBox21.Lines.ElementAt(j);
                    dostawca05[j] = textBox20.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox22.Lines.Count(); k++)
                {
                    System.Windows.Forms.SendKeys.SendWait(art05[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait(dostawca05[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait(status05[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F1}");
                    Thread.Sleep(500);
                }
                Application.Restart();
            }
            else
            {
                string warning = "Błąd wprowadzania danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            textBox24.Clear();
            textBox23.Clear();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            if ((textBox24.Lines.Count() == textBox23.Lines.Count()) && (textBox24.Lines.Count() > 0))
            {
                string[] art06 = new string[3000];
                string[] kod06 = new string[3000];
                for (int i = 0; i < 3000; i++)
                {
                    art06[i] = "";
                    kod06[i] = "";
                }
                for (int j = 0; j < textBox24.Lines.Count(); j++)
                {
                    art06[j] = textBox24.Lines.ElementAt(j);
                    kod06[j] = textBox23.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox24.Lines.Count(); k++)
                {
                    System.Windows.Forms.SendKeys.SendWait(art06[k].ToString());
                    Thread.Sleep(150);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(150);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(150);
                    System.Windows.Forms.SendKeys.SendWait("+{F3}");
                    Thread.Sleep(150);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(50);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(50);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(50);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(50);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(50);
                    System.Windows.Forms.SendKeys.SendWait(kod06[k].ToString());
                    Thread.Sleep(150);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(150);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(150);
                    System.Windows.Forms.SendKeys.SendWait("+{F1}");
                    Thread.Sleep(250);
                }
                Application.Restart();
            }
            else
            {
                string warning = "Ilość numerów artykułów nie jest zgodna z ilością kodów marki bądź nie wprowadzono żadnych danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            textBox26.Clear();
            textBox25.Clear();
        }

        private void button34_Click(object sender, EventArgs e)
        {
            if ((textBox26.Lines.Count() == textBox25.Lines.Count()) && (textBox26.Lines.Count() > 0))
            {
                string[] art07 = new string[3000];
                string[] opis07 = new string[3000];
                for (int i = 0; i < 3000; i++)
                {
                    art07[i] = "";
                    opis07[i] = "";
                }
                for (int j = 0; j < textBox26.Lines.Count(); j++)
                {
                    art07[j] = textBox26.Lines.ElementAt(j);
                    opis07[j] = textBox25.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox26.Lines.Count(); k++)
                {
                    System.Windows.Forms.SendKeys.SendWait(art07[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    System.Windows.Forms.SendKeys.SendWait(opis07[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F1}");
                    Thread.Sleep(500);
                }
                Application.Restart();
            }
            else
            {
                string warning = "Ilość numerów artykułów nie jest zgodna z ilością opisów bądź nie wprowadzono żadnych danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            textBox29.Clear();
            textBox28.Clear();
            textBox27.Clear();
        }

        private void button36_Click(object sender, EventArgs e)
        {
            if ((textBox29.Lines.Count() == textBox28.Lines.Count()) && (textBox29.Lines.Count() == textBox27.Lines.Count()) && (textBox29.Lines.Count() > 0))
            {
                string[] art08 = new string[3000];
                string[] status08 = new string[3000];
                string[] data08 = new string[3000];
                for (int i = 0; i < 3000; i++)
                {
                    art08[i] = "";
                    status08[i] = "";
                    data08[i] = "";
                }
                for (int j = 0; j < textBox29.Lines.Count(); j++)
                {
                    art08[j] = textBox29.Lines.ElementAt(j);
                    status08[j] = textBox28.Lines.ElementAt(j);
                    data08[j] = textBox27.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox29.Lines.Count(); k++)
                {
                    System.Windows.Forms.SendKeys.SendWait(art08[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F6}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{F10}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait(status08[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait(data08[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{F3}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{F12}");
                    Thread.Sleep(500);
                }

                Application.Restart();
            }
            else
            {
                string warning = "Liczba wartości w poszczególnych kolumnach jest rózna bądź nie wprowadzono żadnych danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            textBox31.Clear();
            textBox30.Clear();
        }

        private void button38_Click(object sender, EventArgs e)
        {
            if ((textBox31.Lines.Count() == textBox30.Lines.Count()) && (textBox31.Lines.Count() > 0))
            {
                string[] art09 = new string[3000];
                string[] nom09 = new string[3000];
                for (int i = 0; i < 3000; i++)
                {
                    art09[i] = "";
                    nom09[i] = "";
                }
                for (int j = 0; j < textBox31.Lines.Count(); j++)
                {
                    art09[j] = textBox31.Lines.ElementAt(j);
                    nom09[j] = textBox30.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox31.Lines.Count(); k++)
                {
                    System.Windows.Forms.SendKeys.SendWait(art09[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F3}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait("{TAB}");
                    Thread.Sleep(100);
                    System.Windows.Forms.SendKeys.SendWait(nom09[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F1}");
                    Thread.Sleep(500);
                }
                Application.Restart();
            }
            else
            {
                string warning = "Ilość numerów artykułów nie jest zgodna z ilością pozycji w kolumnie nomenklatura bądź nie wprowadzono żadnych danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            textBox33.Clear();
            textBox32.Clear();
        }

        private void button40_Click(object sender, EventArgs e)
        {
            if ((textBox33.Lines.Count() == textBox32.Lines.Count()) && (textBox33.Lines.Count() > 0))
            {
                string[] art10 = new string[3000];
                string[] gama10 = new string[3000];
                for (int i = 0; i < 3000; i++)
                {
                    art10[i] = "";
                    gama10[i] = "";
                }
                for (int j = 0; j < textBox33.Lines.Count(); j++)
                {
                    art10[j] = textBox33.Lines.ElementAt(j);
                    gama10[j] = textBox32.Lines.ElementAt(j);
                }

                System.Windows.Forms.SendKeys.SendWait("%{TAB}");
                Thread.Sleep(500);
                for (int k = 0; k < textBox33.Lines.Count(); k++)
                {
                    /*
                    System.Windows.Forms.SendKeys.SendWait(art10[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F10}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait(gama10[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F12}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F1}");
                    Thread.Sleep(500);
                    */
                    System.Windows.Forms.SendKeys.SendWait(art10[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{+}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F10}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait(gama10[k].ToString());
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("{ENTER}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F12}");
                    Thread.Sleep(300);
                    System.Windows.Forms.SendKeys.SendWait("+{F1}");
                    Thread.Sleep(500);
                }
                Application.Restart();
            }
            else
            {
                string warning = "Ilość numerów artykułów nie jest zgodna z ilością pozycji w kolumnie gama lokalna bądź nie wprowadzono żadnych danych.";
                MessageBox.Show(warning, "Błąd wprowadzania danych.");
            }
        }

        private void button45_Click(object sender, EventArgs e)
        {
            
        }

        private void button45_Click_1(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox12.Lines.ElementAt(0);

                count1 = textBox15.Lines.Count();
                if (textBox12.Lines.Count() != textBox15.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox12.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button47_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox18.Lines.ElementAt(0);

                count1 = textBox19.Lines.Count();
                if (textBox18.Lines.Count() != textBox19.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox18.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button46_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox17.Lines.ElementAt(0);

                count1 = textBox19.Lines.Count();
                if (textBox17.Lines.Count() != textBox19.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox17.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button49_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox16.Lines.ElementAt(0);

                count1 = textBox19.Lines.Count();
                if (textBox16.Lines.Count() != textBox19.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox16.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button48_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox18.Lines.ElementAt(0);

                count1 = textBox19.Lines.Count();
                if (textBox18.Lines.Count() != textBox19.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox18.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }

            try
            {
                int count1 = 0;
                string firstrow = textBox17.Lines.ElementAt(0);

                count1 = textBox19.Lines.Count();
                if (textBox17.Lines.Count() != textBox19.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox17.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }

            try
            {
                int count1 = 0;
                string firstrow = textBox16.Lines.ElementAt(0);

                count1 = textBox19.Lines.Count();
                if (textBox16.Lines.Count() != textBox19.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox16.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button51_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox21.Lines.ElementAt(0);

                count1 = textBox22.Lines.Count();
                if (textBox21.Lines.Count() != textBox22.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox21.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button50_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox20.Lines.ElementAt(0);

                count1 = textBox22.Lines.Count();
                if (textBox20.Lines.Count() != textBox22.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox20.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button52_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox21.Lines.ElementAt(0);

                count1 = textBox22.Lines.Count();
                if (textBox21.Lines.Count() != textBox22.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox21.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }

            try
            {
                int count1 = 0;
                string firstrow = textBox20.Lines.ElementAt(0);

                count1 = textBox22.Lines.Count();
                if (textBox20.Lines.Count() != textBox22.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox20.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button53_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox23.Lines.ElementAt(0);

                count1 = textBox24.Lines.Count();
                if (textBox23.Lines.Count() != textBox24.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox23.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button55_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox28.Lines.ElementAt(0);

                count1 = textBox29.Lines.Count();
                if (textBox28.Lines.Count() != textBox29.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox28.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button54_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox27.Lines.ElementAt(0);

                count1 = textBox29.Lines.Count();
                if (textBox27.Lines.Count() != textBox29.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox27.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button56_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox28.Lines.ElementAt(0);

                count1 = textBox29.Lines.Count();
                if (textBox28.Lines.Count() != textBox29.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox28.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }

            try
            {
                int count1 = 0;
                string firstrow = textBox27.Lines.ElementAt(0);

                count1 = textBox29.Lines.Count();
                if (textBox27.Lines.Count() != textBox29.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox27.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button57_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox30.Lines.ElementAt(0);

                count1 = textBox31.Lines.Count();
                if (textBox30.Lines.Count() != textBox31.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox30.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button58_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox32.Lines.ElementAt(0);

                count1 = textBox33.Lines.Count();
                if (textBox32.Lines.Count() != textBox33.Lines.Count())
                {
                    for (int i = 1; i < count1; i++)
                    {
                        textBox32.Text += System.Environment.NewLine + firstrow;
                    }
                }
            }
            catch { }
        }

        private void button59_Click(object sender, EventArgs e)
        {
            try
            {
                int count1 = 0;
                string firstrow = textBox25.Lines.ElementAt(0);

                count1 = textBox26.Lines.Count();
                if (textBox25.Lines.Count() != textBox26.Lines.Count())
                {
                    textBox25.Text = firstrow + " " + textBox26.Lines.ElementAt(0);
                    for (int i = 1; i < count1; i++)
                    {
                        textBox25.Text += System.Environment.NewLine + firstrow + " " + textBox26.Lines.ElementAt(i);
                    }
                }
            }
            catch { }
        }
    }
}
