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

namespace ryowa_DATA.config
{
    public partial class frmDupDatadel : Form
    {
        public frmDupDatadel()
        {
            InitializeComponent();
        }

        ryowaDataSet dts = new ryowaDataSet();
        ryowaDataSetTableAdapters.T_勤怠TableAdapter adp = new ryowaDataSetTableAdapters.T_勤怠TableAdapter();

        private void button1_Click(object sender, EventArgs e)
        {
            if (eCheck())
            {
                if (MessageBox.Show(txtYear.Text + "年" + txtMonth.Text + "月の勤怠重複データをクリーニングします。よろしいですか？",
                "確認",MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                {
                    return;
                }

                // mdbバックアップファイル作成
                mdbCopy();

                // 重複データ削除
                dupDataDelete();

                // 閉じる
                Close();
            }
        }

        private void mdbCopy()
        {
            DateTime dt = DateTime.Now;
            string frMdb = Properties.Settings.Default.mdbPath + "ryowa.mdb";
            string toMdb = Properties.Settings.Default.mdbPath + "ryowa" +
                           dt.Year.ToString() + dt.Month.ToString() + dt.Day.ToString() +
                           dt.Hour.ToString() + dt.Minute.ToString() + dt.Second.ToString() +
                           ".mdb";

            System.IO.File.Copy(frMdb, toMdb);
        }


        private bool eCheck()
        {
            if (Utility.StrtoInt(txtYear.Text) == global.flgOff)
            {
                MessageBox.Show("年が正しくありません", "確認",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (Utility.StrtoInt(txtMonth.Text) == global.flgOff)
            {
                MessageBox.Show("月が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            if (Utility.StrtoInt(txtMonth.Text) < 1 || Utility.StrtoInt(txtMonth.Text) > 12)
            {
                MessageBox.Show("月が正しくありません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtYear.Focus();
                return false;
            }

            return true;
        }


        private void dupDataDelete()
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                adp.FillByYYMM(dts.T_勤怠, Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text));

                DateTime wDt = DateTime.Parse("1900/01/01");
                int wSaCode = 0;
                int dCnt = 0;

                foreach (var t in dts.T_勤怠.OrderBy(a => a.社員ID).ThenBy(a => a.日付))
                {
                    if (wSaCode == 0)
                    {
                        wSaCode = t.社員ID;
                        wDt = t.日付;
                        continue;
                    }


                    if (wSaCode == t.社員ID && wDt == t.日付)
                    {
                        // 重複データとみなす
                        t.Delete();
                        dCnt++;
                    }
                    else
                    {
                        wSaCode = t.社員ID;
                        wDt = t.日付;
                    }
                }

                adp.Update(dts.T_勤怠);

                MessageBox.Show(dCnt + "件の重複データを削除しました");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Cursor = Cursors.Default;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmDupDatadel_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void frmDupDatadel_Load(object sender, EventArgs e)
        {
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();
        }
    }
}
