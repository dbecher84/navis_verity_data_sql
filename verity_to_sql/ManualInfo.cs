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
    public partial class ManualInfo : Form
    {
        public static string ProjectNumForm { get; set; }

        public static string ModelNameForm { get; set; }

        public static string BuildingNameForm { get; set; }

        public ManualInfo()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProjectNumForm = textBox1.Text;
            ModelNameForm = textBox2.Text;
            BuildingNameForm = textBox3.Text;
        }

    }
}
