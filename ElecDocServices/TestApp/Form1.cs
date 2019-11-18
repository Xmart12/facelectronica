using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ElecDocServices;

namespace TestApp
{
    public partial class Form1 : Form
    {
        //private string config = "C:\\Projects\\TFS\\Compusal\\Resources\\config.config";
        private string config = "C:\\Temp\\config.config";
        private FacturacionElectronica reg = null;

        public Form1()
        {
            InitializeComponent();

            reg = new FacturacionElectronica(config, "USUARIO");
        }


        //REgistrar
        private void button1_Click(object sender, EventArgs e)
        {
            string res = textBox6.Text;
            string doc = textBox1.Text;
            string ser = textBox2.Text;
            string corr = textBox3.Text;

            try
            {

                bool re = reg.RegistrarDocNC(res, doc, ser, corr);

                if (re)
                {
                    MessageBox.Show("NC Registrada: " + reg.Mensaje);
                    reg.ObtenerDocumento(res, doc, ser, corr);
                }
                else
                {
                    if (reg.Mensaje != null)
                    {
                        MessageBox.Show("Error, respuesta: " + reg.Mensaje);
                    }
                    else
                    {
                        MessageBox.Show("Error, revisar log");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }


        //Descargar
        private void button2_Click(object sender, EventArgs e)
        {
            string res = textBox6.Text;
            string doc = textBox1.Text;
            string ser = textBox2.Text;
            string corr = textBox3.Text;

            try
            {

                bool re = reg.ObtenerDocumento(res, doc, ser, corr);

                if (re)
                {
                    MessageBox.Show("Factura descargada");
                }
                else
                {
                    if (reg.Mensaje != null)
                    {
                        MessageBox.Show("Error, respuesta: " + reg.Mensaje);
                    }
                    else
                    {
                        MessageBox.Show("Error, revisar log");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }


        //Anular
        private void button3_Click(object sender, EventArgs e)
        {
            string res = textBox6.Text;
            string doc = textBox1.Text;
            string ser = textBox2.Text;
            string corr = textBox3.Text;

            try
            {

                bool re = reg.AnularDocNC(res, doc, ser, corr);

                if (re)
                {
                    MessageBox.Show("Factura anulada: " + reg.Mensaje);
                }
                else
                {
                    if (reg.Mensaje != null)
                    {
                        MessageBox.Show("Error, respuesta: " + reg.Mensaje);
                    }
                    else
                    {
                        MessageBox.Show("Error, revisar log");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
