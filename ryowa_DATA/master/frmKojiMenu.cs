using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ryowa_DATA.master
{
    public partial class frmKojiMenu : Form
    {
        public frmKojiMenu()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmKojiMenu_Load(object sender, EventArgs e)
        {

        }

        private void frmKojiMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Hide();
            master.frmMsKoji frm = new master.frmMsKoji();
            frm.ShowDialog();
            this.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmKojiIDCnv frm = new master.frmKojiIDCnv();
            frm.ShowDialog();
        }
    }
}
