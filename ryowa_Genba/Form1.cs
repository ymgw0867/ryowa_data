using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ryowa_Genba.common;

namespace ryowa_Genba
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // 個人情報登録確認
            if (msStatus())
            {
                // 登録済みのときログイン
                frmLogin frm = new frmLogin();
                frm.ShowDialog();

                // ログイン未完了のときは終了します
                if (!global.loginStatus)
                {
                    Environment.Exit(0);
                }
            }
            else
            {
                MessageBox.Show("使用する方の情報を登録してください", "システム準備", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 社員マスター登録
                this.Hide();
                master.frmMsShain frm = new master.frmMsShain();
                frm.ShowDialog();
                this.Show();

                MessageBox.Show("システムを一旦終了します。再度起動後、ログインしてください", "システム起動", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Environment.Exit(0);
            }
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     使用者情報登録確認 </summary>
        /// <returns>
        ///     true:登録済み, false:未登録</returns>
        ///----------------------------------------------------------------
        private bool msStatus()
        {
            genbaDataSet dts = new genbaDataSet();
            genbaDataSetTableAdapters.M_社員TableAdapter adp = new genbaDataSetTableAdapters.M_社員TableAdapter();
            adp.Fill(dts.M_社員);

            if (dts.M_社員.Count() > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // メール情報未登録のとき、出勤簿入力を不可とする
            if (!hpStatus())
            {
                string msg = "";
                msg = "現在、必要なマスター情報が未登録のため、出勤簿入力ができません。" + Environment.NewLine;
                msg += "以下の処理を行ってください。" + Environment.NewLine + Environment.NewLine;
                msg += "メール情報設定" + Environment.NewLine;
                msg += "休日カレンダーのインポート" + Environment.NewLine;
                msg += "工事・所属部署のインポート";

                MessageBox.Show(msg, "システム準備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                linkLabel7.Enabled = false;
            }
            else
            {
                linkLabel7.Enabled = true;
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void linkLabel7_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            data.frmKintaiList frm = new data.frmKintaiList();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            master.frmMsImport frm = new master.frmMsImport();
            frm.ShowDialog();
            this.Show();

            // メール情報、休日カレンダー、工事・所属部署未登録のとき、出勤簿入力を不可とする
            if (!hpStatus())
            {
                linkLabel7.Enabled = false;
            }
            else
            {
                linkLabel7.Enabled = true;
            }
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            master.frmMsShain frm = new master.frmMsShain();
            frm.ShowDialog();
            this.Show();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Hide();
            master.frmMsMail frm = new master.frmMsMail();
            frm.ShowDialog();
            this.Show();

            // メール情報、休日カレンダー、工事・所属部署未登録のとき、出勤簿入力を不可とする
            if (!hpStatus())
            {
                linkLabel7.Enabled = false;
            }
            else
            {
                linkLabel7.Enabled = true;
            }
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     休日カレンダー、工事・所属部署マスター、メール設定登録確認 </summary>
        /// <returns>
        ///     true:登録済み, false:未登録</returns>
        ///--------------------------------------------------------------------------
        private bool hpStatus()
        {
            genbaDataSet dts = new genbaDataSet();
            genbaDataSetTableAdapters.M_休日TableAdapter adp = new genbaDataSetTableAdapters.M_休日TableAdapter();
            genbaDataSetTableAdapters.M_工事TableAdapter pAdp = new genbaDataSetTableAdapters.M_工事TableAdapter();
            genbaDataSetTableAdapters.メール設定TableAdapter mAdp = new genbaDataSetTableAdapters.メール設定TableAdapter();

            adp.Fill(dts.M_休日);
            pAdp.Fill(dts.M_工事);
            mAdp.Fill(dts.メール設定);

            // 休日カレンダーと工事・所属部署、メール設定が登録済みか？
            if (dts.M_休日.Count() > 0 && dts.M_工事.Count() > 0 && dts.メール設定.Count() > 0)
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
