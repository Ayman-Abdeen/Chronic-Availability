using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Chronic_Availability
{
    public partial class logenForm : Form
    {
        public logenForm()
        {
            InitializeComponent();
        }

        public String userName { get; set; }
        public String password { get; set; }

        private void Login_butt_Click(object sender, EventArgs e)
        {
            userName = textBox1.Text;
            password = textBox2.Text;

            this.Hide();
            mainForm form = new mainForm(userName, password);
            form.Show();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
