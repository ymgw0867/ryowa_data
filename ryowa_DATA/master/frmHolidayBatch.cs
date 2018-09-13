using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ryowa_DATA.common;

namespace ryowa_DATA.master
{
    public partial class frmHolidayBatch : Form
    {
        // マスター名
        string msName = "休日一括登録";

        // フォームモードインスタンス
        Utility.frmMode fMode = new Utility.frmMode();
        
        // データテーブル生成
        ryowaDataSet dts = new ryowaDataSet();

        // 休日マスターテーブルアダプター生成
        ryowaDataSetTableAdapters.M_休日TableAdapter adp = new ryowaDataSetTableAdapters.M_休日TableAdapter();

        int C_SUNDAY = 0;   // 日曜日
        int C_SATURDAY = 6; // 土曜日

        public frmHolidayBatch()
        {
            InitializeComponent();

            // データテーブルにデータを読み込む
            adp.Fill(dts.M_休日);
        }

        //カラム定義
        string cDate = "col1";
        string cMemo = "col2";
        string cDel = "col3";
        string cHoutei = "col4";
        string cWeek = "col5";
        string cAID = "col6";
        string cADate = "col7";
        string cEID = "col8";
        string cEDate = "col9";
        string cID = "col10";

        private void frm_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッドビュー定義
            GridViewSetting(dg);

            // データグリッドビュー表示
            GridViewShow(dg);

            // 画面初期化
            DispInitial();
        }

        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        private void GridViewSetting(DataGridView g)
        {
            try
            {
                g.EnableHeadersVisualStyles = false;
                g.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                // 列スタイルを変更する

                g.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                g.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 列ヘッダーフォント指定
                g.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // データフォント指定
                g.DefaultCellStyle.Font = new Font("Meiryo UI", 9, FontStyle.Regular);

                // 行の高さ
                g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                g.ColumnHeadersHeight = 20;
                g.RowTemplate.Height = 20;

                // 全体の高さ
                g.Height = 382;

                // 奇数行の色
                g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.UseColumnTextForButtonValue = true;
                btn.Text = "削除";
                g.Columns.Add(btn);
                g.Columns[0].Name = cDel;
                g.Columns[cDel].HeaderText = "";

                g.Columns.Add(cDate, "年月日");
                g.Columns.Add(cMemo, "摘要");
                g.Columns.Add(cHoutei, "法定休日");
                g.Columns.Add(cWeek, "曜日");
                g.Columns.Add(cAID, "登録者");
                g.Columns.Add(cADate, "登録年月日");
                g.Columns.Add(cEID, "更新者");
                g.Columns.Add(cEDate, "更新年月日");
                g.Columns.Add(cID, "IDコード");

                g.Columns[cDel].Width = 50;
                g.Columns[cDate].Width = 90;
                g.Columns[cMemo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                g.Columns[cHoutei].Visible = false;
                g.Columns[cWeek].Visible = false;
                g.Columns[cAID].Visible = false;
                g.Columns[cADate].Visible = false;
                g.Columns[cEID].Visible = false;
                g.Columns[cEDate].Visible = false;
                g.Columns[cID].Visible = false;

                g.Columns[cDate].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                // 行ヘッダを表示しない
                g.RowHeadersVisible = false;

                // 選択モード
                g.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                g.MultiSelect = false;

                // 編集不可とする
                g.ReadOnly = true;

                // 追加行表示しない
                g.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                g.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                g.AllowUserToOrderColumns = false;

                // 列サイズ変更可
                g.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                g.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                //TAB動作
                g.StandardTab = true;

                // 罫線
                g.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                g.CellBorderStyle = DataGridViewCellBorderStyle.None;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューにデータを表示します  </summary>
        /// <param name="tempGrid">
        ///     データグリッドビューオブジェクト名</param>
        ///---------------------------------------------------------------------
        private void GridViewShow(DataGridView g)
        {
            g.Rows.Clear();

            int iX = 0;
            global gl = new global();

            try
            {
                foreach (var t in dts.M_休日.OrderByDescending(a => a.日付))
                {
                    g.Rows.Add();

                    g[cDate, iX].Value = t.日付.ToShortDateString();
                    g[cMemo, iX].Value = t.摘要;
                    g[cHoutei, iX].Value = t.法定休日.ToString();
                    g[cID, iX].Value = t.ID.ToString();
                    g[cWeek, iX].Value = t.曜日.ToString();
                    g[cAID, iX].Value = t.登録ユーザーID.ToString();
                    g[cADate, iX].Value = t.更新年月日.ToString();
                    g[cEID, iX].Value = t.更新ユーザーID.ToString();
                    g[cEDate, iX].Value = t.更新年月日.ToString();

                    iX++;
                }

                if (g.Rows.Count > 0)
                {
                    g.CurrentCell = null;
                    linkLabel5.Enabled = true;
                }
                else
                {
                    linkLabel5.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     画面の初期化 </summary>
        ///---------------------------------------------------------------
        private void DispInitial()
        {
            radioButton1.Checked = true;

            // 一括登録
            dt1.Value = DateTime.Today;
            dt2.Value = DateTime.Today;
            chkSun.CheckState = CheckState.Unchecked;
            chkSat.CheckState = CheckState.Unchecked;
            chkShuku.CheckState = CheckState.Unchecked;
            chkSonota.CheckState = CheckState.Unchecked;
            txtMemo.Text = string.Empty;

            linkLabel4.Enabled = true;
        }

        

        private void btnUpdate_Click(object sender, EventArgs e)
        {
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     休日マスター追加 </summary>
        /// <returns>
        ///     登録件数</returns>
        /// ------------------------------------------------------------------
        private int addHolidayBatch()
        {
            int cnt = 0;

            for (DateTime dt = dt1.Value; dt <= dt2.Value; dt = dt.AddDays(1))
            {
                // 休日摘要初期化
                string teki = string.Empty;

                // 祝日の設定
                if (chkShuku.Checked)
                {
                    string mmdd = dt.Month.ToString().PadLeft(2, '0') + "/" + dt.Day.ToString().PadLeft(2, '0');

                    for (int i = 0; i < global.wHoriDay.GetLength(0); i++)
                    {
                        if (mmdd == global.wHoriDay[i, 0])
                        {
                            teki = global.wHoriDay[i, 1];   // 祝日名を取得
                            break;
                        }
                    }
                }

                // 祝日以外の日曜日または土曜日のとき
                if (teki == string.Empty)
                {
                    if ((chkSun.Checked && (int)dt.DayOfWeek == 0) || (chkSat.Checked && (int)dt.DayOfWeek == 6))
                    {
                        // 摘要の設定
                        if ((int)dt.DayOfWeek == C_SUNDAY)
                        {
                            teki = "日曜";
                        }

                        if ((int)dt.DayOfWeek == C_SATURDAY)
                        {
                            teki = "土曜";
                        }
                    }
                }

                // 任意の期間の休日
                if (teki == string.Empty)
                {
                    teki = txtMemo.Text;
                }

                // 休日で未登録の日付のとき、新規登録
                if (teki != string.Empty && !dts.M_休日.Any(a => a.日付 == dt && a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached))
                {
                    var d = dts.M_休日.NewM_休日Row();
                    d.日付 = dt;
                    d.曜日 = (int)dt.DayOfWeek;

                    if (d.曜日 == C_SUNDAY)
                    {
                        d.法定休日 = global.flgOn;
                    }
                    else
                    {
                        d.法定休日 = global.flgOff;
                    }

                    d.摘要 = teki;
                    d.登録ユーザーID = global.loginUserID;
                    d.登録年月日 = DateTime.Now;
                    d.更新ユーザーID = global.loginUserID;
                    d.更新年月日 = DateTime.Now;

                    dts.M_休日.AddM_休日Row(d);
                    cnt++;
                }
            }

            return cnt;
        }

        private int singleDataAdd()
        {
            int cnt = 0;
            DateTime dt = dateTimePicker1.Value;

            // 休日で未登録の日付のとき、新規登録
            if (txtMemo2.Text != string.Empty && !dts.M_休日.Any(a => a.日付 == dt && a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached))
            {
                var d = dts.M_休日.NewM_休日Row();
                d.日付 = dt;
                d.曜日 = (int)dt.DayOfWeek;

                if (d.曜日 == C_SUNDAY)
                {
                    d.法定休日 = global.flgOn;
                }
                else
                {
                    d.法定休日 = global.flgOff;
                }

                d.摘要 = txtMemo2.Text;
                d.登録ユーザーID = global.loginUserID;
                d.登録年月日 = DateTime.Now;
                d.更新ユーザーID = global.loginUserID;
                d.更新年月日 = DateTime.Now;

                dts.M_休日.AddM_休日Row(d);
                cnt++;
            }
            return cnt; 
        }


        // 登録データチェック
        private Boolean fDataCheck()
        {
            try
            {
                if (radioButton1.Checked)
                {
                    // 期間
                    if (dt1.Value > dt2.Value)
                    {
                        dt1.Focus();
                        throw new Exception("設定期間が正しくありません");
                    }

                    // 曜日選択
                    if (chkSun.CheckState == CheckState.Unchecked && chkSat.CheckState == CheckState.Unchecked && chkShuku.CheckState == CheckState.Unchecked)
                    {
                        chkSun.Focus();
                        throw new Exception("曜日を選択してください");
                    }

                    // その他の休日
                    if (chkSonota.Checked && txtMemo.Text == string.Empty)
                    {
                        txtMemo.Focus();
                        throw new Exception("その他休日の適用を入力してください");
                    }

                    return true;
                }
                else
                {
                    // その他の休日
                    if (txtMemo2.Text == string.Empty)
                    {
                        txtMemo2.Focus();
                        throw new Exception("休日の適用を入力してください");
                    }

                    return true;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, msName + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frm_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Dispose();
        }

        private void frmKintaiKbn_Shown(object sender, EventArgs e)
        {
            dt1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void txtCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            frmShukujitsu frm = new frmShukujitsu();
            frm.ShowDialog();
        }

        private void chkSonota_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSonota.Checked)
            {
                txtMemo.Enabled = true;
            }
            else
            {
                txtMemo.Enabled = false;
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // フォームを閉じます
            this.Close();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //エラーチェック
            if (!fDataCheck())
            {
                return;
            }

            // 確認
            string msg = string.Empty;

            if (radioButton1.Checked)
            {
                msg = "休日一括登録を行います。";
            }
            else
            {
                msg = "個別登録を行います。";
            }

            if (MessageBox.Show(msg + "よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
            {
                return;
            }

            int n = 0;

            if (radioButton1.Checked)
            {
                // 一括登録を行います
                n = addHolidayBatch();
            }
            else
            {
                // 個別登録を行います
                n = singleDataAdd();
            }

            // 確認表示
            MessageBox.Show(n.ToString() + "件登録しました", "休日一括登録", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // データセットの内容をデータベースへ反映させます。
            if (n > 0)
            {
                adp.Update(dts.M_休日);

                // グリッドビュー再表示
                GridViewShow(dg);
            }
        }

        private void dg_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (e.ColumnIndex == 0)
                {
                    // データ削除 
                    dateDelete(e.RowIndex);
                }
            }
        }

        private bool dateDelete(int r)
        {
            string iDate = dg[cDate, r].Value.ToString();
            int iid = Utility.StrtoInt(dg[cID, r].Value.ToString());

            if (MessageBox.Show(iDate + " を削除します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return false;
            }

            // 休日データ削除
            var s = dts.M_休日.Single(a => a.ID == iid);
            s.Delete();

            // データベースアップデート
            adp.Update(dts.M_休日);

            // グリッドビュー再表示
            GridViewShow(dg);

            return true;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                // 一括登録
                label4.Enabled = true;
                panel1.Enabled = true;
                label5.Enabled = false;
                panel5.Enabled = false;
            }
            else
            {
                label4.Enabled = false;
                panel1.Enabled = false;
                label5.Enabled = true;
                panel5.Enabled = true;
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime dt = dateTimePicker1.Value;

            string teki = string.Empty; 
            
            // 祝日の設定
            string mmdd = dt.Month.ToString().PadLeft(2, '0') + "/" + dt.Day.ToString().PadLeft(2, '0');

            for (int i = 0; i < global.wHoriDay.GetLength(0); i++)
            {
                if (mmdd == global.wHoriDay[i, 0])
                {
                    teki = global.wHoriDay[i, 1];   // 祝日名を取得
                    break;
                }
            }

            // 祝日以外の日曜日または土曜日のとき
            if (teki == string.Empty)
            {
                if ((int)dt.DayOfWeek == 0 || (int)dt.DayOfWeek == 6)
                {
                    // 摘要の設定
                    if ((int)dt.DayOfWeek == C_SUNDAY)
                    {
                        teki = "日曜";
                    }

                    if ((int)dt.DayOfWeek == C_SATURDAY)
                    {
                        teki = "土曜";
                    }
                }
            }

            if (teki != string.Empty)
            {
                txtMemo2.Text = teki;
            }
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MyLibrary.CsvOut.GridView(dg, "休日設定");
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
        }

    }
}
