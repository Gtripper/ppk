using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinFormForPPKParser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //OpenFileDialog open = new OpenFileDialog();
            //open.DefaultExt = "*.exe";
            //open.Filter = "EXE Files (*.exe)|*.exe";
            //open.FilterIndex = 1;
            //open.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;

            //if (open.ShowDialog() == DialogResult.OK)
            //{
            //    textBox1.Text = open.FileName;
            //}

            FolderBrowserDialog open = new FolderBrowserDialog();
            open.RootFolder = Environment.SpecialFolder.Desktop;
            open.SelectedPath = AppDomain.CurrentDomain.BaseDirectory;

            if (open.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = open.SelectedPath;
            }


        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.DefaultExt = "*.exe";
            open.Filter = "Excel Files (*.xlsx)|*.xlsx|(*.xls)|*.xls";
            open.FilterIndex = 1;
            open.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;

            if (open.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = open.FileName;
            }
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var t = Int32.Parse(richTextBox2.Text);
                richTextBox1.BackColor = Color.Empty;
                if (!(t >= 1 && t <= 1000))
                {
                    richTextBox1.BackColor = Color.DimGray;
                    richTextBox1.Text = "";
                }
            }
            catch
            {
                richTextBox1.BackColor = Color.DimGray;
                richTextBox1.Text = "";
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            try
            {
                var t = Int32.Parse(richTextBox1.Text);
                richTextBox1.BackColor = Color.Empty;
                if (!(t >= 1 && t <= 1000))
                {
                    richTextBox1.BackColor = Color.DimGray;
                    richTextBox1.Text = "";
                }
            }
            catch 
            {
                richTextBox1.BackColor = Color.DimGray;
                richTextBox1.Text = "";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                var lenghtOfFlow = Int32.Parse(richTextBox1.Text);
                var NumFlows = Int32.Parse(richTextBox2.Text);
                var driverPath = (string)textBox1.Text;
                var ExcelPath = (string)textBox2.Text;
                //ExcelApp(path, driverPath, lenghtOfRow, Flows);
                //ppk5_v2.IParser parser = new ppk5_v2.Parser(driverPath, ExcelPath, NumFlows, lenghtOfFlow);
                //parser.RunParsingOKS();
                //IParser parser = new IParser();
            }
            catch (Exception ex)
            {
                Console.WriteLine("OOOPS EXCEPTION HERE!111");
                Console.WriteLine(ex);
            }


            //pkk_5_parser.Program.ExcelApp();
        }
    }
}
