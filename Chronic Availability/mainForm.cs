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
    public partial class mainForm : Form
    {

        public String user { get; set; }
        public String pas { get; set; }

        public mainForm(string userName, string passwored)
        {
            InitializeComponent();
            user = userName;
            pas = passwored;
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
        }

        private void chonicToolStripMenuItem_Click(object sender, EventArgs e)
        {
            chronic chronicForm = new chronic();
            chronicForm.MdiParent = this; 
            chronicForm.Show();
        }

        private void mainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }
    }
}
