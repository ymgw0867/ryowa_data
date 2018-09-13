using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ryowa_DATA.common;

namespace ryowa_DATA.master
{
    public partial class frmKojiIDCnv : Form
    {
        public frmKojiIDCnv()
        {
            InitializeComponent();
        }

        ryowaDataSet dts = new ryowaDataSet();
        ryowaDataSetTableAdapters.M_工事TableAdapter kAdp = new ryowaDataSetTableAdapters.M_工事TableAdapter();
        ryowaDataSetTableAdapters.T_勤怠TableAdapter tAdp = new ryowaDataSetTableAdapters.T_勤怠TableAdapter();

        bool frmID = false;
        bool toID = false;

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // 画面初期化
            dispInitial(); 
        }

        private void dispInitial()
        {
            txtFrmID.Text = string.Empty;
            txtToID.Text = string.Empty;

            lblFrmName.Text = string.Empty;
            lblToName.Text = string.Empty;

            btnok.Enabled = false;
        }

        private void txtFrmID_TextChanged(object sender, EventArgs e)
        {
            frmID = false;
            btnok.Enabled = false;

            if (txtFrmID.Text.Length != 6)
            {
                lblFrmName.Text = string.Empty;
                return;
            }

            if (txtFrmID.Text.Length == 6)
            {
                kAdp.FillByID(dts.M_工事, Utility.StrtoInt(txtFrmID.Text));

                if (dts.M_工事.Count != 0)
                {
                    foreach (var t in dts.M_工事)
                    {
                        lblFrmName.Text = t.名称;
                        frmID = true;
                    }
                }
            }

            if (frmID && toID)
            {
                btnok.Enabled = true;
            }
        }

        private void txtToID_TextChanged(object sender, EventArgs e)
        {
            toID = false;
            btnok.Enabled = false;

            if (txtToID.Text.Length != 6)
            {
                lblToName.Text = string.Empty;
                return;
            }

            if (txtToID.Text.Length == 6)
            {
                kAdp.FillByID(dts.M_工事, Utility.StrtoInt(txtToID.Text));

                if (dts.M_工事.Count != 0)
                {
                    MessageBox.Show("既に登録済みの工事ＩＤです", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
            }

            toID = true;

            if (frmID && toID)
            {
                btnok.Enabled = true;
            }
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmKojiIDCnv_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void btnok_Click(object sender, EventArgs e)
        {
            if (!frmID || !toID)
            {
                return;
            }
            
            if (MessageBox.Show("工事マスター「" + txtFrmID.Text + "：" + lblFrmName.Text + "」のＩＤを「" + txtFrmID.Text + "」から「" + txtToID.Text + "」に 変更します。登録済みの勤怠データの工事ＩＤも一括変更されます。" + Environment.NewLine + Environment.NewLine + "実行してよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (MessageBox.Show("本当に実行して良いですか？", "再度の確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 工事マスター更新
            if (mstUpdate(Utility.StrtoInt(txtToID.Text), Utility.StrtoInt(txtFrmID.Text)))
            {
                string msg = "工事マスターのＩＤ変更に成功しました。" + Environment.NewLine + Environment.NewLine;

                int n = 0;

                // 勤怠データ＠工事ＩＤ更新
                bool ks = kintaiIDUpdate(Utility.StrtoInt(txtToID.Text), Utility.StrtoInt(txtFrmID.Text), out n);

                if (ks)
                {
                    if (n > 0)
                    {
                        msg += "該当する " + n + "件の勤怠データの工事ＩＤを変更しました。";
                    }
                    else
                    {
                        msg += "変更元工事ＩＤに該当する勤怠データは存在しませんでした。";
                    }
                }
                else
                {
                    msg += "勤怠データの工事ＩＤの変更に失敗しました。";
                }

                MessageBox.Show(msg, "処理終了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("工事マスターのＩＤの変更に失敗しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            // 画面初期化
            dispInitial();

        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     工事マスター更新
        ///     新ＩＤレコードを新規登録、旧ＩＤレコードを削除 </summary>
        /// <param name="newID">
        ///     新ＩＤ</param>
        /// <param name="oldID">
        ///     旧ＩＤ</param>
        /// <returns>
        ///     true:更新成功, false:更新失敗 </returns>
        ///--------------------------------------------------------------------------
        private bool mstUpdate(int newID, int oldID)
        {
            try
            {
                kAdp.FillByID(dts.M_工事, oldID);

                var s = dts.M_工事.Single(a => a.ID > 0);

                // 新ＩＤレコードを新規登録
                kAdp.InsertQueryNewID(newID, s.名称, s.現場区分, s.勤務地区分, s.勤務地名, s.開始時, s.開始分,
                    s.終了時, s.終了分, global.loginUserID, DateTime.Now, global.loginUserID, DateTime.Now);

                try
                {
                    // 旧ＩＤレコードを削除する
                    kAdp.DeleteQueryOldID(oldID);
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                    // 新ＩＤレコードを削除する
                    kAdp.DeleteQueryOldID(newID);

                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     勤怠データ＠工事ＩＤ更新 </summary>
        /// <param name="newID">
        ///     新ＩＤ</param>
        /// <param name="oldID">
        ///     旧ＩＤ</param>
        /// <returns>
        ///     更新件数</returns>
        ///--------------------------------------------------------------
        private bool kintaiIDUpdate(int newID, int oldID, out int n)
        {
            try
            {
                n = tAdp.UpdateQueryKoujiID(newID, oldID);
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                n = 0;
                return false;
            }
        }
    }
}
