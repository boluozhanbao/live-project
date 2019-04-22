using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VBAInExcel
{
    public partial class Form1 : Form
    {
        int cnt = 0;
        public Form1()
        {
            InitializeComponent();

            comboBox1.Items.Add("一等奖");
            comboBox1.SelectedIndex = 0;
            comboBox1.Items.Add("二等奖");
            comboBox1.Items.Add("三等奖");
            comboBox1.Items.Add("四等奖");
            comboBox1.Items.Add("五等奖");
            listBox1.Font = new Font(this.Font.FontFamily, 16);
            listBox2.Font = new Font(this.Font.FontFamily, 16);

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void button1_Click(object sender, EventArgs e)
        {
            //listBox2.Items.Add("success!!!");
            string Begintime;
            string Endtime;
            Begintime = dateTimePicker1.Value.Year.ToString() + "-" + dateTimePicker1.Value.Month.ToString() + "-" + dateTimePicker1.Value.Day.ToString();
            Endtime = dateTimePicker2.Value.Year.ToString() + "-" + dateTimePicker2.Value.Month.ToString() + "-" + dateTimePicker2.Value.Day.ToString();
            FileProcessing fp = new FileProcessing(textBox4.Text);
            //listBox2.Items.Add("success1!!!");
            fp.Read();
            //listBox2.Items.Add("success2!!!");
            Pprogram.filter(fp.messages, Begintime, Endtime, textBox1.Text);
            //listBox2.Items.Add("success3!!!");
            string[] person = Pprogram.GetArray();
            //listBox2.Items.Add("success4!!!");
            int n = Pprogram.Getnum();
            //listBox2.Items.Add(n);
            //listBox2.Items.Add("success5!!!");
            List<long> ans = new List<long>();
            //listBox2.Items.Add("success6!!!");
            string key = Pprogram.GetBTC();
            //listBox2.Items.Add("success7!!!");
            ans = Pprogram.GetWinners(0, n, cnt, key);
            //listBox2.Items.Add("success8!!!");
            foreach (long l in ans)
            {
                listBox2.Items.Add(person[l]);
            }    
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string str;
            str= comboBox1.Text + ":\r\n" + textBox3.Text + "*" + textBox2.Text + "\n";
            listBox1.Items.Add(str);
            cnt += Convert.ToInt32(textBox2.Text);
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
                listBox1.Items.Remove(listBox1.SelectedItem);
            if (listBox1.Items.Count == 0)
                return;
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.InitialDirectory = "C:\\";    //打开对话框后的初始目录
            fileDialog.Filter = "文本文件|*.txt|所有文件|*.*";
            fileDialog.RestoreDirectory = false;    //若为false，则打开对话框后为上次的目录。若为true，则为初始目录
            if (fileDialog.ShowDialog() == DialogResult.OK)
                textBox4.SelectedText = fileDialog.FileName;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
