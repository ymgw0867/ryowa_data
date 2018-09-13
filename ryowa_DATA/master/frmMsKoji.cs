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
    public partial class frmMsKoji : Form
    {
        // マスター名
        string msName = "工事";

        // フォームモードインスタンス
        Utility.frmMode fMode = new Utility.frmMode();

        // 工事マスターテーブルアダプター生成
        ryowaDataSetTableAdapters.M_工事TableAdapter adp = new ryowaDataSetTableAdapters.M_工事TableAdapter();
        ryowaDataSetTableAdapters.M_社員1TableAdapter mAdp = new ryowaDataSetTableAdapters.M_社員1TableAdapter();

        // データテーブル生成
        ryowa_DATA.ryowaDataSet dts = new ryowaDataSet();

        public frmMsKoji()
        {
            InitializeComponent();

            // データテーブルにデータを読み込む
            adp.Fill(dts.M_工事);
            mAdp.Fill(dts.M_社員1);
        }

        private void frm_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッド定義
            GridViewSetting(dg);

            // データグリッドビューにデータを表示します
            GridViewShow(dg);

            // コンボ値ロード
            cmbLoad();

            // 画面初期化
            DispInitial();
        }

        //カラム定義
        string cCode = "col1";
        string cName = "col2";
        string cGenbaKbn = "col3";
        string cKinmuchiKbn = "col4";
        string cKinmuchi = "col5";
        string cAccount = "col6";
        string cDate = "col7";
        string cEAccount = "col8";
        string cEDate = "col9";
        string cSTM = "col10";
        string cETM = "col11";
        string cGenbaKbnID = "col12";
        string cKinmuchiKbnID = "col13";
        string cAccountID = "col14";
        string cEAccountID = "col15";

        private void cmbLoad()
        {
            global g = new global();

            // 現場区分
            for (int i = 0; i < g.arrStyle.GetLength(0); i++)
            {
                cmbKbn.Items.Add(g.arrStyle[i, 1]);
            }

            // 勤務地区分
            for (int i = 0; i < g.kArrStyle.GetLength(0); i++)
            {
                cmbKinmuchi.Items.Add(g.kArrStyle[i, 1]);
            }
        }
        
        ///-------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///-------------------------------------------------------------------
        private void GridViewSetting(DataGridView g)
        {
            try
            {
                g.EnableHeadersVisualStyles = false;
                g.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

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
                g.Height = 301;

                // 奇数行の色
                g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                g.Columns.Add(cCode, "コード");
                g.Columns.Add(cName, "名称");
                g.Columns.Add(cGenbaKbn, "現場区分");
                g.Columns.Add(cKinmuchiKbn, "勤務地区分");
                g.Columns.Add(cKinmuchi, "勤務地名");
                g.Columns.Add(cSTM, "開始時刻");
                g.Columns.Add(cETM, "終了時刻");
                g.Columns.Add(cAccount, "登録者");
                g.Columns.Add(cDate, "登録年月日");
                g.Columns.Add(cEAccount, "更新者");
                g.Columns.Add(cEDate, "更新年月日");
                g.Columns.Add(cGenbaKbnID, "現場区分ID");
                g.Columns.Add(cKinmuchiKbnID, "勤務地区分ID");
                g.Columns.Add(cAccountID, "登録者ID");
                g.Columns.Add(cEAccountID, "更新者ID");

                g.Columns[cCode].Width = 70;
                g.Columns[cName].Width = 300;
                g.Columns[cGenbaKbn].Width = 80;
                g.Columns[cKinmuchiKbn].Width = 60;
                g.Columns[cKinmuchi].Width = 120;
                g.Columns[cSTM].Width = 110;
                g.Columns[cETM].Width = 110;
                g.Columns[cAccount].Width = 120;
                g.Columns[cDate].Width = 160;
                g.Columns[cEAccount].Width = 120;
                g.Columns[cEDate].Width = 160;

                g.Columns[cGenbaKbnID].Visible = false;
                g.Columns[cKinmuchiKbnID].Visible = false;
                g.Columns[cAccountID].Visible = false;
                g.Columns[cEAccountID].Visible = false;

                g.Columns[cSTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cETM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                //tempDGV.Columns[C_Memo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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
        ///     データグリッドビューにデータを表示します </summary>
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
                foreach (var t in dts.M_工事.OrderBy(a => a.ID))
                {
                    g.Rows.Add();
                    g[cCode, iX].Value = t.ID.ToString();
                    g[cName, iX].Value = t.名称.ToString();
                    g[cGenbaKbn, iX].Value = string.Empty;

                    for (int i = 0; i < gl.arrStyle.GetLength(0); i++)
                    {
                        if (t.現場区分.ToString() == gl.arrStyle[i, 0])
                        {
                            g[cGenbaKbn, iX].Value = gl.arrStyle[i, 1];
                        }
                    }

                    g[cKinmuchiKbn, iX].Value = string.Empty;

                    for (int i = 0; i < gl.kArrStyle.GetLength(0); i++)
                    {
                        if (t.勤務地区分.ToString() == gl.kArrStyle[i, 0])
                        {
                            g[cKinmuchiKbn, iX].Value = gl.kArrStyle[i, 1];
                        }
                    }

                    g[cKinmuchi, iX].Value = t.勤務地名;

                    if (t.Is開始時Null() || t.Is開始分Null())
                    {
                        g[cSTM, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[cSTM, iX].Value = t.開始時.ToString() + ":" + t.開始分.ToString().PadLeft(2, '0');
                    }

                    if (t.Is終了時Null() || t.Is終了分Null())
                    {
                        g[cETM, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[cETM, iX].Value = t.終了時.ToString() + ":" + t.終了分.ToString().PadLeft(2, '0');
                    }

                    if (t.M_社員1RowByM_工事_M_社員1 == null)
                    {
                        g[cAccount, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[cAccount, iX].Value = t.M_社員1RowByM_工事_M_社員1.氏名;
                    }

                    g[cDate, iX].Value = t.登録年月日;

                    if (t.M_社員1RowByM_工事_M_社員11 == null)
                    {
                        g[cEAccount, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[cEAccount, iX].Value = t.M_社員1RowByM_工事_M_社員11.氏名;
                    }

                    g[cEDate, iX].Value = t.更新年月日;

                    g[cGenbaKbnID, iX].Value = t.現場区分;
                    g[cKinmuchiKbnID, iX].Value = t.勤務地区分;
                    g[cAccountID, iX].Value = t.登録ユーザーID;
                    g[cEAccountID, iX].Value = t.更新ユーザーID;

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

        /// <summary>
        /// 画面の初期化
        /// </summary>
        private void DispInitial()
        {
            fMode.Mode = global.FORM_ADDMODE;
            txtCode.Text = string.Empty;    // 2018/07/19
            txtName.Text = string.Empty;
            cmbKbn.SelectedIndex = 0;
            cmbKinmuchi.SelectedIndex = 0;
            txtKinmuchi.Text = string.Empty;
            txtSh.Text = Properties.Settings.Default.startH;
            txtSm.Text = Properties.Settings.Default.startM.PadLeft(2, '0');
            txtEh.Text = Properties.Settings.Default.endH;
            txtEm.Text = Properties.Settings.Default.endM.PadLeft(2, '0');

            linkLabel4.Enabled = true;
            linkLabel2.Enabled = false;
            linkLabel3.Enabled = false;
            //txtName.Focus();

            txtCode.Enabled = true;     // 2018/07/19
            txtCode.Focus();            // 2018/07/19
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
        }

        //登録データチェック
        private Boolean fDataCheck()
        {
            try
            {
                // 新規登録時
                if (fMode.Mode == global.FORM_ADDMODE)
                {
                    if (Utility.StrtoInt(txtCode.Text) == 0)
                    {
                        txtCode.Focus();
                        throw new Exception("工事コードを入力してください");
                    }

                    if (dts.M_工事.Count(a => a.ID == Utility.StrtoInt(txtCode.Text)) > 0)
                    {
                        txtCode.Focus();
                        throw new Exception("既に登録済みの工事コードです");
                    }
                }

                //名称チェック
                if (txtName.Text.Trim().Length < 1)
                {
                    txtName.Focus();
                    throw new Exception(msName + "名を入力してください");
                }
                
                // 現場区分
                if (cmbKbn.SelectedIndex == -1)
                {
                    cmbKbn.Focus();
                    throw new Exception("現場区分を選択してください");
                }

                // 勤務地区分
                if (cmbKinmuchi.SelectedIndex == -1)
                {
                    cmbKbn.Focus();
                    throw new Exception("勤務地区分を選択してください");
                }

                // 開始時刻
                if (txtSh.Text == string.Empty)
                {
                    txtSh.Focus();
                    throw new Exception("開始時刻を入力してください");
                }

                if (Utility.StrtoInt(txtSh.Text) > 23)
                {
                    txtSh.Focus();
                    throw new Exception("開始時刻は０～２３の範囲で入力してください");
                }

                if (txtSm.Text == string.Empty)
                {
                    txtSm.Focus();
                    throw new Exception("開始時刻を入力してください");
                }

                if (Utility.StrtoInt(txtSm.Text) > 59)
                {
                    txtSm.Focus();
                    throw new Exception("開始時刻は０～５９の範囲で入力してください");
                }

                // 終了時刻
                if (txtEh.Text == string.Empty)
                {
                    txtEh.Focus();
                    throw new Exception("終了時刻を入力してください");
                }

                if (Utility.StrtoInt(txtEh.Text) > 23)
                {
                    txtEh.Focus();
                    throw new Exception("終了時刻は０～２３の範囲で入力してください");
                }

                if (txtEm.Text == string.Empty)
                {
                    txtEm.Focus();
                    throw new Exception("終了時刻を入力してください");
                }

                if (Utility.StrtoInt(txtEm.Text) > 59)
                {
                    txtEm.Focus();
                    throw new Exception("終了時刻は０～５９の範囲で入力してください");
                }

                int st = Utility.StrtoInt(txtSh.Text) * 100 + Utility.StrtoInt(txtSm.Text);
                int et = Utility.StrtoInt(txtEh.Text) * 100 + Utility.StrtoInt(txtEm.Text);

                if (st >= et)
                {
                    if (MessageBox.Show("勤務終了時刻が勤務開始時刻以前となっています。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        return false;
                    }
                }

                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, msName + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        /// <summary>
        /// グリッドビュー行選択時処理
        /// </summary>
        private void GridEnter()
        {
            string msgStr;
            fMode.rowIndex = dg.SelectedRows[0].Index;

            // 選択確認
            msgStr = "";
            msgStr += dg[1, fMode.rowIndex].Value.ToString() + "：" + dg[3, fMode.rowIndex].Value.ToString() + Environment.NewLine + Environment.NewLine;
            msgStr += "上記の" + msName + "が選択されました。よろしいですか？";

            if (MessageBox.Show(msgStr, "選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No) 
                return;

            // 対象となるデータテーブルROWを取得します
            ryowaDataSet.M_工事Row sQuery = dts.M_工事.FindByID(int.Parse(dg[0, fMode.rowIndex].Value.ToString()));

            if (sQuery != null)
            {
                // モードステータスを「編集モード」にします
                fMode.Mode = global.FORM_EDITMODE;

                // 編集画面に表示
                ShowData(sQuery);
            }
            else
            {
                MessageBox.Show(dg[0, fMode.rowIndex].Value.ToString() + "がキー不在です：データの読み込みに失敗しました", "データ取得エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        /// -------------------------------------------------------
        /// <summary>
        ///     マスターの内容を画面に表示する </summary>
        /// <param name="sTemp">
        ///     マスターインスタンス</param>
        /// -------------------------------------------------------
        private void ShowData(ryowaDataSet.M_工事Row s)
        {
            fMode.ID = s.ID.ToString();
            txtCode.Text = s.ID.ToString(); // 2018/07/19

            // 2018/07/19
            if (fMode.Mode == global.FORM_EDITMODE)
            {
                txtCode.Enabled = false;
            }

            txtName.Text = s.名称;
            cmbKbn.SelectedIndex = s.現場区分;
            cmbKinmuchi.SelectedIndex = s.勤務地区分;
            txtKinmuchi.Text = s.勤務地名;

            if (s.Is開始時Null())
            {
                txtSh.Text = string.Empty;
            }
            else
            {
                txtSh.Text = s.開始時.ToString();
            }

            if (s.Is開始分Null())
            {
                txtSm.Text = string.Empty;
            }
            else
            {
                txtSm.Text = s.開始分.ToString().PadLeft(2, '0');
            }

            if (s.Is終了時Null())
            {
                txtEh.Text = string.Empty;
            }
            else
            {
                txtEh.Text = s.終了時.ToString();
            }

            if (s.Is終了分Null())
            {
                txtEm.Text = string.Empty;
            }
            else
            {
                txtEm.Text = s.終了分.ToString().PadLeft(2, '0');
            }

            linkLabel2.Enabled = true;
            linkLabel3.Enabled = true;
        }

        private void dg_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            GridEnter();
        }

        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // データセットの内容をデータベースへ反映させます
            adp.Update(dts.M_工事);

            this.Dispose();
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
        }

        private void frmKintaiKbn_Shown(object sender, EventArgs e)
        {
            //txtName.Focus();
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

        private void txtSh_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void txtSm_Leave(object sender, EventArgs e)
        {
            txtSm.Text = txtSm.Text.PadLeft(2, '0');
        }

        private void txtEm_Leave(object sender, EventArgs e)
        {
            txtEm.Text = txtEm.Text.PadLeft(2, '0');
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //エラーチェック
            if (!fDataCheck())
            {
                return;
            }

            switch (fMode.Mode)
            {
                // 新規登録
                case global.FORM_ADDMODE:

                    // 確認
                    if (MessageBox.Show(txtName.Text + "を登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // データセットにデータを追加します
                    var s = dts.M_工事.NewM_工事Row();
                    s.ID = Utility.StrtoInt(txtCode.Text);
                    s.名称 = txtName.Text;
                    s.現場区分 = cmbKbn.SelectedIndex;
                    s.勤務地区分 = cmbKinmuchi.SelectedIndex;
                    s.勤務地名 = txtKinmuchi.Text;
                    s.開始時 = Utility.StrtoInt(txtSh.Text);
                    s.開始分 = Utility.StrtoInt(txtSm.Text);
                    s.終了時 = Utility.StrtoInt(txtEh.Text);
                    s.終了分 = Utility.StrtoInt(txtEm.Text);
                    s.登録ユーザーID = global.loginUserID;
                    s.登録年月日 = DateTime.Now;
                    s.更新ユーザーID = global.loginUserID;
                    s.更新年月日 = DateTime.Now;

                    dts.M_工事.AddM_工事Row(s);

                    break;

                // 更新処理
                case global.FORM_EDITMODE:

                    // 確認
                    if (MessageBox.Show(txtName.Text + "を更新します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // データセット更新
                    var r = dts.M_工事.Single(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                               a.ID == int.Parse(fMode.ID));

                    if (!r.HasErrors)
                    {
                        r.名称 = txtName.Text;
                        r.現場区分 = cmbKbn.SelectedIndex;
                        r.勤務地区分 = cmbKinmuchi.SelectedIndex;
                        r.勤務地名 = txtKinmuchi.Text;
                        r.開始時 = Utility.StrtoInt(txtSh.Text);
                        r.開始分 = Utility.StrtoInt(txtSm.Text);
                        r.終了時 = Utility.StrtoInt(txtEh.Text);
                        r.終了分 = Utility.StrtoInt(txtEm.Text);
                        r.更新ユーザーID = global.loginUserID;
                        r.更新年月日 = DateTime.Now;
                    }
                    else
                    {
                        MessageBox.Show(fMode.ID + "がキー不在です：データの更新に失敗しました", "更新エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }

                    break;

                default:
                    break;
            }

            // 更新をコミット
            adp.Update(dts.M_工事);

            // データテーブルにデータを読み込む
            adp.Fill(dts.M_工事);

            // 画面データ消去
            DispInitial();

            // グリッド表示
            GridViewShow(dg);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                // 確認
                if (MessageBox.Show("削除してよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                    return;

                // 削除データ取得（エラー回避のためDataRowState.Deleted と DataRowState.Detachedは除外して抽出する）
                var d = dts.M_工事.Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached && a.ID == int.Parse(fMode.ID));

                // foreach用の配列を作成する
                var list = d.ToList();

                // 削除
                foreach (var it in list)
                {
                    ryowaDataSet.M_工事Row dl = (ryowaDataSet.M_工事Row)dts.M_工事.Rows.Find(it.ID);
                    dl.Delete();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("データの削除に失敗しました" + Environment.NewLine + ex.Message);
            }
            finally
            {
                // 削除をコミット
                adp.Update(dts.M_工事);

                // データテーブルにデータを読み込む
                adp.Fill(dts.M_工事);

                // 画面データ消去
                DispInitial();

                // グリッド表示
                GridViewShow(dg);
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DispInitial();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // フォームを閉じます
            this.Close();
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MyLibrary.CsvOut.GridView(dg, "工事・所属部署マスター");
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
        }
    }
}
