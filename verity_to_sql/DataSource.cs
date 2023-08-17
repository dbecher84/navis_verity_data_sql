using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace verity_to_sql
{
    public partial class DataSource : Form
    {
        public string usenavisdata = "false";
        public string usecsvdata = "false";

        public string SetDataSource(RadioButton input_button1, RadioButton input_button2)
        {
            if (input_button1.Checked)
            {
                usenavisdata = "true";
                return usenavisdata;
            }
            if (input_button2.Checked)
            {
                usecsvdata = "true";
                return usecsvdata;
            }
            else
            {
                return "fail";
            }
        }

        public DataSource()
        {
            InitializeComponent();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetDataSource(radioButton1, radioButton2);
        }
    }
}
