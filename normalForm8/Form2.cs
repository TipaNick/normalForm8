using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace normalForm8
{
    public partial class Form2 : Form
    {
        public string employee1FIO { get; set; }
        public string employee1Job { get; set; }
        public string employee2FIO { get; set; }
        public string employee2Job { get; set; }
        public string employee3FIO { get; set; }
        public string employee3Job { get; set; }
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            employee1FIO = textBox1.Text;
            employee1Job = textBox2.Text;
            employee2FIO = textBox4.Text;
            employee2Job = textBox3.Text;
            employee3FIO = textBox6.Text;
            employee3Job = textBox5.Text;
            this.Close();
        }
    }
}
