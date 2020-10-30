using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QuickDemonstrationAddIn
{
    public partial class Form1 : Form
    {
        private String myval;
        public String MyVal
        {
            get { return myval; }
            set { myval = value; }
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)  //Connect
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("Bạn chưa điền địa chỉ kết nối!");
            }
            MyVal = textBox1.Text;
        }

        private void button2_Click(object sender, EventArgs e) //Huy
        {

        }

        private void progressBar2_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
