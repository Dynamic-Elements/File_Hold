using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Accounting_PL
{
    public partial class Form2 : Form
    {
        public Form2(string title)
        {
            InitializeComponent();
            textBox1.Text = title;
        }

        private void button1_Click(object sender, EventArgs e)  // Save and Send data back to main form
        {

        }
    }
}
