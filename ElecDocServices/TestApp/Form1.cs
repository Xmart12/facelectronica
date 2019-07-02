using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TestApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        //REgistrar
        private void button1_Click(object sender, EventArgs e)
        {
            string res = textBox6.Text;
            string doc = textBox1.Text;
            string ser = textBox2.Text;
            string corr = textBox3.Text;
        }


        //Descargar
        private void button2_Click(object sender, EventArgs e)
        {

        }


        //Anular
        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
