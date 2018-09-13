using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ryowa_Confir
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Hide();
            data.frmRequest frm = new data.frmRequest();
            frm.ShowDialog();
            Show();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 閉じる
            Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Hide();
            data.frmAuthSend frm = new data.frmAuthSend();
            frm.ShowDialog();
            Show();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!hpStatus())
            {
                string msg = "";
                msg = "メール情報が未設定のため、使用できません。" + Environment.NewLine;
                msg += "出勤簿登録メニューの「メール情報設定」を行った後、再度処理を行ってください。";

                MessageBox.Show(msg, "システム準備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(0);
            }
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     メール設定登録確認 </summary>
        /// <returns>
        ///     true:登録済み, false:未登録</returns>
        ///--------------------------------------------------------------------------
        private bool hpStatus()
        {
            confirDataSet dts = new confirDataSet();
            confirDataSetTableAdapters.メール設定TableAdapter mAdp = new confirDataSetTableAdapters.メール設定TableAdapter();
            
            mAdp.Fill(dts.メール設定);

            // メール設定が登録済みか？
            if (dts.メール設定.Count() > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
