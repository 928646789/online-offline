using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace online_offline_Scanning
{
    public partial class validation : Form
    {
        bool OK = false;
        public validation()
        {
            InitializeComponent();
        }

        private void validation_FormClosing(object sender, FormClosingEventArgs e)
        {
            if(!OK)
                System.Environment.Exit(0);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "PKVPE")
            {
                OK = true;
                this.Close();
            }
            else
            {
                MessageBox.Show("please import the correct password");
            }
                
        }
    }
}
