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

        /// <summary>
        /// Save Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)  // Save and Send data back to main form
        {
            // Will combine line items for master link



        }

        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e)
        {
            // Add columns together
            decimal totalSalary = 0;
            decimal qty = 0;
            decimal amt = 0;

            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                var value = dataGridView1.Rows[i].Cells[2].Value;
                var value1 = dataGridView1.Rows[i].Cells[3].Value;
                if (value != DBNull.Value && value1 != DBNull.Value)
                {
                    qty = Convert.ToDecimal(value);
                    amt = Convert.ToDecimal(value1);
                    totalSalary += qty * amt;
                }
            }

            textBox2.Text = totalSalary.ToString("C");
        }
    }
}
