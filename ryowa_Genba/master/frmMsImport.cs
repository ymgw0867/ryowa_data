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

namespace ryowa_Genba.master
{
    public partial class frmMsImport : Form
    {
        public frmMsImport()
        {
            InitializeComponent();

            // マスター読み込み
            hAdp.Fill(dts.M_休日);
            pAdp.Fill(dts.M_工事);
        }

        genbaDataSet dts = new genbaDataSet();
        genbaDataSetTableAdapters.M_休日TableAdapter hAdp = new genbaDataSetTableAdapters.M_休日TableAdapter();
        genbaDataSetTableAdapters.M_工事TableAdapter pAdp = new genbaDataSetTableAdapters.M_工事TableAdapter();

        string msg = "";

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                msg = "工事・所属部署";
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmMsImport_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // インポートデータ選択
            string fName = openCsvFile();
            if (fName == string.Empty)
            {
                return;
            }
            else
            {
                lblFname.Text = fName;
            }
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     インポートデータ選択ダイアログボックス </summary>
        /// <returns>
        ///     パスを含めたインポートデータファイル名</returns>
        ///---------------------------------------------------------------
        private string openCsvFile()
        {
            DialogResult ret;

            //ダイアログボックスの初期設定
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Title = "インポートデータの選択";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "CSVファイル(*.csv)|*.csv";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            //if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "ファイル確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            //{
            //    return string.Empty;
            //}

            return openFileDialog1.FileName;
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     休日カレンダーCSVファイルインポート </summary>
        /// <param name="fName">
        ///     CSVファイルパス</param>
        /// -------------------------------------------------------------------
        public void holMsImport(string fName)
        {
            int cnt = 0;

            try
            {
                this.Cursor = Cursors.WaitCursor;

                // 休日カレンダー全件削除
                holMsDataDelete();

                // CSVファイルインポート
                var s = System.IO.File.ReadAllLines(fName, Encoding.Default);
                foreach (var stBuffer in s)
                {
                    // 1行目はネグる
                    if (cnt == 0)
                    {
                        cnt++;
                        continue;
                    }

                    // カンマ区切りで分割して配列に格納する
                    string[] stCSV = stBuffer.Split(',');

                    // 登録
                    sAddMaster(stCSV);

                    cnt++;
                }

                // データベースへ反映
                hAdp.Update(dts.M_休日);

                // 終了メッセージ
                MessageBox.Show(cnt.ToString() + "件の休日データを登録しました。", "処理終了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "休日マスターインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     工事・所属部署CSVファイルインポート </summary>
        /// <param name="fName">
        ///     CSVファイルパス</param>
        /// -------------------------------------------------------------------
        public void koujiMsImport(string fName)
        {
            int cnt = 0;

            try
            {
                this.Cursor = Cursors.WaitCursor;

                // 工事・所属部署全件削除
                koujiMsDataDelete();

                // CSVファイルインポート
                var s = System.IO.File.ReadAllLines(fName, Encoding.Default);
                foreach (var stBuffer in s)
                {
                    // 1行目はネグる
                    if (cnt == 0)
                    {
                        cnt++;
                        continue;
                    }

                    // カンマ区切りで分割して配列に格納する
                    string[] stCSV = stBuffer.Split(',');

                    // 登録
                    kAddMaster(stCSV);

                    cnt++;
                }

                // データベースへ反映
                pAdp.Update(dts.M_工事);

                // 終了メッセージ
                MessageBox.Show(cnt.ToString() + "件の工事・所属部署データを登録しました。", "処理終了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "工事・所属部署マスターインポート処理", MessageBoxButtons.OK);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     休日カレンダーにインポートデータ追加する </summary>
        /// <param name="c">
        ///     CSVデータ配列</param>
        /// ---------------------------------------------------------------
        private void sAddMaster(string[] c)
        {
            DateTime dt;
            genbaDataSet.M_休日Row r = dts.M_休日.NewM_休日Row();

            r.ID = Utility.StrtoInt(c[9]);
            r.日付 = DateTime.Parse(c[1]);
            r.曜日 = Utility.StrtoInt(c[4]);
            r.法定休日 = Utility.StrtoInt(c[3]);
            r.摘要 = c[2];
            r.登録ユーザーID = Utility.StrtoInt(c[5]);

            if (DateTime.TryParse(c[6], out dt))
            {
                r.登録年月日 = dt;
            }
            else
            {
                r.登録年月日 = DateTime.Now;
            }
            
            r.更新ユーザーID = Utility.StrtoInt(c[7]);

            if (DateTime.TryParse(c[8], out dt))
            {
                r.更新年月日 = dt;
            }
            else
            {
                r.更新年月日 = DateTime.Now;
            }

            dts.M_休日.AddM_休日Row(r);
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     工事・所属部署マスターにインポートデータを追加する </summary>
        /// <param name="c">
        ///     CSVデータ配列</param>
        /// ---------------------------------------------------------------
        private void kAddMaster(string[] c)
        {
            DateTime dt;
            genbaDataSet.M_工事Row r = dts.M_工事.NewM_工事Row();

            r.ID = Utility.StrtoInt(c[0]);
            r.名称 = c[1];
            r.現場区分 = Utility.StrtoInt(c[11]);
            r.勤務地区分 = Utility.StrtoInt(c[12]);
            r.勤務地名 = c[4];
            r.開始時 = timeSplit(c[5], 1);
            r.開始分 = timeSplit(c[5], 2);
            r.終了時 = timeSplit(c[6], 1);
            r.終了分 = timeSplit(c[6], 2);
            r.登録ユーザーID = Utility.StrtoInt(c[13]);

            if (DateTime.TryParse(c[8], out dt))
            {
                r.登録年月日 = dt;
            }
            else
            {
                r.登録年月日 = DateTime.Now;
            }

            r.更新ユーザーID = Utility.StrtoInt(c[14]);

            if (DateTime.TryParse(c[10], out dt))
            {
                r.更新年月日 = dt;
            }
            else
            {
                r.更新年月日 = DateTime.Now;
            }

            dts.M_工事.AddM_工事Row(r);
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     時間：分表記を時分ごとに分けて返す </summary>
        /// <param name="t">
        ///     時間：分表記文字列</param>
        /// <param name="rtn">
        ///     1:時間、2:分を返す</param>
        /// <returns>
        ///     戻り値</returns>
        ///---------------------------------------------------------------
        private int timeSplit(string t, int rtn)
        {
            string[] arr = t.Split(':');

            if (arr.Length < 2)
            {
                return 0;
            }
            else
            {
                return Utility.StrtoInt(arr[rtn - 1]);
            }
        } 


        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            string msg2 = "を更新します。" +  Environment.NewLine + msg + "マスターはインポートするCSVデータの内容に全て書き換えられます" + Environment.NewLine + "よろしいですか？";

            if (!radioButton1.Checked && !radioButton2.Checked)
            {
                MessageBox.Show("インポートするマスターを選択してください", "マスター未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (lblFname.Text == string.Empty)
            {
                MessageBox.Show("インポートするCSVデータを選択してください", "CSVデータ未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button1.Focus();
                return;
            }

            if (!System.IO.File.Exists(lblFname.Text))
            {
                MessageBox.Show("指定されたCSVデータは存在しません。再度CSVデータを選択してください", "CSVデータ未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button1.Focus();
                return;
            }
            
            if (MessageBox.Show(msg + msg2,"確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (radioButton1.Checked)
            {
                // 休日カレンダー更新
                holMsImport(lblFname.Text);
            }
            else if (radioButton2.Checked)
            {
                // 工事・所属部署更新
                koujiMsImport(lblFname.Text);
            }

            // 閉じる
            Close();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                msg = "休日カレンダー";
            }
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     休日カレンダーマスター全件削除 </summary>
        ///----------------------------------------------------------------
        private void holMsDataDelete()
        {
            foreach (var h in dts.M_休日.OrderBy(a => a.ID))
            {
                h.Delete();
            }

            hAdp.Update(dts.M_休日);
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     工事・所属部署マスター全件削除 </summary>
        ///----------------------------------------------------------------
        private void koujiMsDataDelete()
        {
            foreach (var h in dts.M_工事.OrderBy(a => a.ID))
            {
                h.Delete();
            }

            pAdp.Update(dts.M_工事);
        }

        private void frmMsImport_Load(object sender, EventArgs e)
        {
            lblFname.Text = string.Empty;
        }
    }
}
