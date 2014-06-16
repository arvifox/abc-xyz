using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace excel2
{
    public partial class Form1 : Form
    {
        List<Good> ListOfGoods;
        int SumYear;
        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel.Application oe = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook oeb = oe.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Microsoft.Office.Interop.Excel.Worksheet oes;
                oes = (Microsoft.Office.Interop.Excel.Worksheet)oeb.Sheets[1];
                ListOfGoods = new List<Good>();
                for (int i = 4; true; i++)
                {
                    if (oes.Cells[i, 2].text == "")
                    {
                        break;
                    }
                    //Выбираем область таблицы. (в нашем случае просто ячейку)
                    //Microsoft.Office.Interop.Excel.Range range = ObjWorkSheet.get_Range(textBox1.Text + i.ToString(), textBox1.Text + i.ToString());
                    //Добавляем полученный из ячейки текст.
                    //richTextBox1.Text = richTextBox1.Text + range.Text.ToString() + "\n";
                    //это чтобы форма прорисовывалась (не подвисала)...
                    //Application.DoEvents();
                    ListViewItem newi = new ListViewItem(oes.Cells[i, 2].text);
                    newi.SubItems.Add(oes.Cells[i, 3].text);
                    newi.SubItems.Add(oes.Cells[i, 4].text);
                    newi.SubItems.Add(oes.Cells[i, 5].text);
                    newi.SubItems.Add(oes.Cells[i, 6].text);
                    listView1.Items.Add(newi);
                    Good ng = new Good();
                    ng.name = oes.Cells[i, 2].text;
                    ng.kv1 = Convert.ToInt32(oes.Cells[i, 3].text);
                    ng.kv2 = Convert.ToInt32(oes.Cells[i, 4].text);
                    ng.kv3 = Convert.ToInt32(oes.Cells[i, 5].text);
                    ng.kv4 = Convert.ToInt32(oes.Cells[i, 6].text);
                    ListOfGoods.Add(ng);
                }
                oe.Quit();
            }
        }

        public int CompareYear(Good p1, Good p2)
        {
            if (p1.year > p2.year) { return -1; } else { if (p1.year < p2.year) { return 1; } else { return 0; } }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SumYear = 0;
            for (int i = 0; i < ListOfGoods.Count; i++)
            {
                Good g = ListOfGoods[i];
                g.year = g.kv1 + g.kv2 + g.kv3 + g.kv4;
                SumYear = SumYear + g.year;
            }
            for (int i = 0; i < ListOfGoods.Count; i++)
            {
                Good g = ListOfGoods[i];
                g.pryear = Math.Round((double)g.year * 100 / SumYear, 2);
            }
            ListOfGoods.Sort(CompareYear);
            Good gg = ListOfGoods[0];
            gg.incall = gg.pryear;
            for (int i = 1; i < ListOfGoods.Count; i++)
            {
                Good prev = ListOfGoods[i - 1];
                Good g = ListOfGoods[i];
                g.incall = prev.incall + g.pryear;
            }
            Form2 f2 = new Form2();
            f2.Show();
            f2.listView1.Items.Clear();
            for (int i = 0; i < ListOfGoods.Count; i++)
            {
                Good g = ListOfGoods[i];
                ListViewItem ni = new ListViewItem(g.name);
                ni.SubItems.Add(Convert.ToString(g.year));
                ni.SubItems.Add(g.pryear.ToString("F2"));
                ni.SubItems.Add(g.incall.ToString("F2"));
                if (g.incall < 80)
                {
                    g.abc = "A";
                }
                else
                {
                    if (g.incall < 95)
                    {
                        g.abc = "B";
                    }
                    else
                    {
                        g.abc = "C";
                    }
                }
                ni.SubItems.Add(g.abc);
                f2.listView1.Items.Add(ni);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            double SumY=0.0, SumPay=0;
            for (int i = 0; i < ListOfGoods.Count; i++)
            {
                Good g = ListOfGoods[i];
                g.paystore = (double)Convert.ToDouble(textBox1.Text) * Convert.ToDouble(textBox3.Text) * g.year / Convert.ToDouble(textBox2.Text);
                SumY = SumY + g.year;
                SumPay = SumPay + g.paystore;
                ListViewItem ni = new ListViewItem(g.name);
                ni.SubItems.Add(Convert.ToString(g.year));
                ni.SubItems.Add(g.paystore.ToString("F1"));
                f3.listView1.Items.Add(ni);
            }
            ListViewItem nl = new ListViewItem("Итого");
            nl.SubItems.Add(SumY.ToString("F1"));
            nl.SubItems.Add(SumPay.ToString("F1"));
            f3.listView1.Items.Add(nl);
            double SumA = 0, SumB = 0, SumC = 0;
            for (int i = 0; i < ListOfGoods.Count; i++)
            {
                Good g = ListOfGoods[i];
                if (g.abc == "A")
                {
                    SumA = SumA + g.year;
                }
                else
                {
                    if (g.abc == "B")
                    {
                        SumB = SumB + g.year;
                    }
                    else
                    {
                        SumC = SumC + g.year;
                    }
                }
            }
            double pa = 0.0, pb = 0.0, pc = 0.0;
            pa = (double)Convert.ToDouble(textBox4.Text) * Convert.ToDouble(textBox3.Text) * SumA / Convert.ToDouble(textBox2.Text);
            pb = (double)Convert.ToDouble(textBox5.Text) * Convert.ToDouble(textBox3.Text) * SumB / Convert.ToDouble(textBox2.Text);
            pc = (double)Convert.ToDouble(textBox6.Text) * Convert.ToDouble(textBox3.Text) * SumC / Convert.ToDouble(textBox2.Text);
            ListViewItem nla = new ListViewItem("A");
            nla.SubItems.Add(SumA.ToString("F1"));
            nla.SubItems.Add(pa.ToString("F1"));
            f3.listView2.Items.Add(nla);
            ListViewItem nlb = new ListViewItem("B");
            nlb.SubItems.Add(SumB.ToString("F1"));
            nlb.SubItems.Add(pb.ToString("F1"));
            f3.listView2.Items.Add(nlb);
            ListViewItem nlc = new ListViewItem("C");
            nlc.SubItems.Add(SumC.ToString("F1"));
            nlc.SubItems.Add(pc.ToString("F1"));
            f3.listView2.Items.Add(nlc);
            ListViewItem nlall = new ListViewItem("All");
            nlall.SubItems.Add((SumA+SumB+SumC).ToString("F1"));
            nlall.SubItems.Add((pa+pb+pc).ToString("F1"));
            f3.listView2.Items.Add(nlall);
            f3.label3.Text = (SumPay - pa - pb - pc).ToString("F1");
            f3.Show();
        }

        public int CompareKvar(Good p1, Good p2)
        {
            if (p1.kvar > p2.kvar) { return 1; } else { if (p1.kvar < p2.kvar) { return -1; } else { return 0; } }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form4 f4 = new Form4();
            for (int i = 0; i < ListOfGoods.Count; i++)
            {
                Good g = ListOfGoods[i];
                g.avkv = (double)(g.kv1 + g.kv2 + g.kv3 + g.kv4) / 4;
                g.kvar = (double)100 / g.avkv * (Math.Sqrt((Math.Pow(g.kv1 - g.avkv,2) + Math.Pow(g.kv2 - g.avkv,2) + Math.Pow(g.kv3 - g.avkv,2) + Math.Pow(g.kv4 - g.avkv,2)) / 4));
            }
            ListOfGoods.Sort(CompareKvar);
            for (int i = 0; i < ListOfGoods.Count; i++)
            {
                Good g = ListOfGoods[i];
                ListViewItem ni = new ListViewItem(g.name);
                ni.SubItems.Add(Convert.ToString(g.year));
                ni.SubItems.Add(Convert.ToString(g.kv1));
                ni.SubItems.Add(Convert.ToString(g.kv2));
                ni.SubItems.Add(Convert.ToString(g.kv3));
                ni.SubItems.Add(Convert.ToString(g.kv4));
                ni.SubItems.Add(g.avkv.ToString("F1"));
                ni.SubItems.Add(g.kvar.ToString("F2"));
                if (g.kvar < 10)
                {
                    g.xyz = "X";
                }
                else
                {
                    if (g.kvar < 25)
                    {
                        g.xyz = "Y";
                    }
                    else
                    {
                        g.xyz = "Z";
                    }
                }
                ni.SubItems.Add(g.xyz);
                f4.listView1.Items.Add(ni);
            }
            f4.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form5 f5 = new Form5();
            for (int i = 0; i < ListOfGoods.Count; i++)
            {
                Good g = ListOfGoods[i];
                if ((g.abc == "A") && (g.xyz == "X"))
                {
                    ListViewItem ni = new ListViewItem(g.name);
                    f5.lvax.Items.Add(ni);
                }
                if ((g.abc == "A") && (g.xyz == "Y"))
                {
                    ListViewItem ni = new ListViewItem(g.name);
                    f5.lvay.Items.Add(ni);
                }
                if ((g.abc == "A") && (g.xyz == "Z"))
                {
                    ListViewItem ni = new ListViewItem(g.name);
                    f5.lvaz.Items.Add(ni);
                }
                if ((g.abc == "B") && (g.xyz == "X"))
                {
                    ListViewItem ni = new ListViewItem(g.name);
                    f5.lvbx.Items.Add(ni);
                }
                if ((g.abc == "B") && (g.xyz == "Y"))
                {
                    ListViewItem ni = new ListViewItem(g.name);
                    f5.lvby.Items.Add(ni);
                }
                if ((g.abc == "B") && (g.xyz == "Z"))
                {
                    ListViewItem ni = new ListViewItem(g.name);
                    f5.lvbz.Items.Add(ni);
                }
                if ((g.abc == "C") && (g.xyz == "X"))
                {
                    ListViewItem ni = new ListViewItem(g.name);
                    f5.lvcx.Items.Add(ni);
                }
                if ((g.abc == "C") && (g.xyz == "Y"))
                {
                    ListViewItem ni = new ListViewItem(g.name);
                    f5.lvcy.Items.Add(ni);
                }
                if ((g.abc == "C") && (g.xyz == "Z"))
                {
                    ListViewItem ni = new ListViewItem(g.name);
                    f5.lvcz.Items.Add(ni);
                }
            }
            f5.Show();
        }
    }
    public class Good
    {
        public string name;
        public int kv1, kv2, kv3, kv4;
        public int year;
        public double pryear;
        public double incall;
        public string abc;
        public string xyz;
        public double paystore;
        public double avkv;
        public double kvar;
    }
}
