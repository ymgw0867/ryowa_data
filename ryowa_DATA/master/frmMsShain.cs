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
    public partial class frmMsShain : Form
    {
        // マスター名
        string msName = "社員マスター";

        // フォームモードインスタンス
        Utility.frmMode fMode = new Utility.frmMode();

        // 社員マスターテーブルアダプター生成
        ryowaDataSetTableAdapters.M_社員TableAdapter adp = new ryowaDataSetTableAdapters.M_社員TableAdapter();
        ryowaDataSetTableAdapters.M_社員1TableAdapter mAdp = new ryowaDataSetTableAdapters.M_社員1TableAdapter();

        // データテーブル生成
        ryowa_DATA.ryowaDataSet dts = new ryowaDataSet();

        public frmMsShain()
        {
            InitializeComponent();

            // データテーブルにデータを読み込む
            adp.Fill(dts.M_社員);
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
            GridViewShow(dg, string.Empty);

            // コンボ値ロード
            cmbLoad();

            // 画面初期化
            DispInitial();
        }

        //カラム定義
        string cCode = "col1";
        string cName = "col2";
        string cFuri = "col3";
        string cSzCode = "col4";
        string cSzName = "col5";
        string cJinTanka = "col6";
        string cZan = "col7";
        string cTooshiTanka = "col8";
        string cKihon10 = "col9";
        string cPassWord = "col10";
        string cKintaiEdit = "col11";
        string cSystemKbn = "col12";
        string cBikou = "col13";
        string cAccount = "col14";
        string cDate = "col15";
        string cEAccount = "col16";
        string cEDate = "col17";
        string cGenbaKbn = "col18";
        string cKoteiZan = "col19";

        private void cmbLoad()
        {
            global g = new global();

            // 残業区分
            for (int i = 0; i < g.zArrStyle.GetLength(0); i++)
            {
                cmbZan.Items.Add(g.zArrStyle[i, 1]);
            }

            // 勤怠編集権限区分
            for (int i = 0; i < g.eArrStyle.GetLength(0); i++)
            {
                cmbKintaiEdit.Items.Add(g.eArrStyle[i, 1]);
            }

            // システムユーザー区分
            for (int i = 0; i < g.uArrStyle.GetLength(0); i++)
            {
                cmbSysKbn.Items.Add(g.uArrStyle[i, 1]);
            }

            // 現場手当有無区分：2018/09/03
            for (int i = 0; i < g.gArrStyle.GetLength(0); i++)
            {
                cmbGenbateate.Items.Add(g.gArrStyle[i, 1]);
            }
        }
        
        ///------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------------
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
                g.Height = 242;

                // 奇数行の色
                g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
                //g.AlternatingRowsDefaultCellStyle.BackColor = Color.Lavender;
                //g.AlternatingRowsDefaultCellStyle.ForeColor = Color.White;

                g.Columns.Add(cCode, "コード");
                g.Columns.Add(cName, "名称");
                g.Columns.Add(cFuri, "フリガナ");
                g.Columns.Add(cSzCode, "所属コード");
                g.Columns.Add(cSzName, "所属名");
                g.Columns.Add(cJinTanka, "人件費単価");
                g.Columns.Add(cZan, "残業有無");
                g.Columns.Add(cTooshiTanka, "通し勤務単価");
                g.Columns.Add(cKihon10, "基本給10％");
                g.Columns.Add(cPassWord, "パスワード");
                g.Columns.Add(cKintaiEdit, "勤怠編集権限");
                g.Columns.Add(cSystemKbn, "システムユーザー区分");
                g.Columns.Add(cBikou, "備考");
                g.Columns.Add(cAccount, "登録者");
                g.Columns.Add(cDate, "登録年月日");
                g.Columns.Add(cEAccount, "更新者");
                g.Columns.Add(cEDate, "更新年月日");
                g.Columns.Add(cGenbaKbn, "現場手当有無");
                g.Columns.Add(cKoteiZan, "固定残業時間");

                g.Columns[cCode].Width = 70;
                g.Columns[cName].Width = 140;
                g.Columns[cFuri].Width = 100;
                g.Columns[cSzCode].Width = 100;
                g.Columns[cSzName].Width = 140;
                g.Columns[cJinTanka].Width = 110;
                g.Columns[cZan].Width = 110;
                g.Columns[cTooshiTanka].Width = 120;
                g.Columns[cKihon10].Width = 120;
                g.Columns[cPassWord].Width = 110;
                g.Columns[cKintaiEdit].Width = 120;
                g.Columns[cSystemKbn].Width = 160;
                g.Columns[cBikou].Width = 200;
                g.Columns[cAccount].Width = 120;
                g.Columns[cDate].Width = 160;
                g.Columns[cEAccount].Width = 120;
                g.Columns[cEDate].Width = 160;
                g.Columns[cGenbaKbn].Width = 120;
                g.Columns[cKoteiZan].Width = 120;

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
        private void GridViewShow(DataGridView g, string sName)
        {
            mAdp.Fill(dts.M_社員1);

            g.Rows.Clear();

            int iX = 0;
            global gl = new global();

            try
            {
                var s = dts.M_社員.OrderBy(a => a.ID);

                if (sName != string.Empty)
                {
                    s = s.Where(a => a.氏名.Contains(sName)).OrderBy(a => a.ID);
                }
                
                foreach (var t in s)
                {
                    g.Rows.Add();
                    g[cCode, iX].Value = t.ID.ToString();
                    g[cName, iX].Value = t.氏名;
                    g[cFuri, iX].Value = t.フリガナ;
                    g[cSzCode, iX].Value = t.所属コード;
                    g[cSzName, iX].Value = t.所属名;
                    g[cJinTanka, iX].Value = t.人件費単価;

                    g[cZan, iX].Value = string.Empty;

                    for (int i = 0; i < gl.zArrStyle.GetLength(0); i++)
                    {
                        if (t.残業有無.ToString() == gl.zArrStyle[i, 0])
                        {
                            g[cZan, iX].Value = gl.zArrStyle[i, 1];
                        }
                    }

                    g[cTooshiTanka, iX].Value = t.通し勤務単価.ToString("#,##0");
                    g[cKihon10, iX].Value = t.基本給10.ToString("#,##0");
                    g[cPassWord, iX].Value = t.パスワード;
                    
                    g[cKintaiEdit, iX].Value = string.Empty;

                    for (int i = 0; i < gl.eArrStyle.GetLength(0); i++)
                    {
                        if (t.アカウント権限.ToString() == gl.eArrStyle[i, 0])
                        {
                            g[cKintaiEdit, iX].Value = gl.eArrStyle[i, 1];
                        }
                    }

                    g[cSystemKbn, iX].Value = string.Empty;

                    for (int i = 0; i < gl.uArrStyle.GetLength(0); i++)
                    {
                        if (t.システムユーザー区分.ToString() == gl.uArrStyle[i, 0])
                        {
                            g[cSystemKbn, iX].Value = gl.uArrStyle[i, 1];
                        }
                    }

                    g[cBikou, iX].Value = t.備考;

                    if (t.M_社員1RowByM_社員1_M_社員 == null)
                    {
                        g[cAccount, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[cAccount, iX].Value = t.M_社員1RowByM_社員1_M_社員.氏名;
                    }

                    g[cDate, iX].Value = t.登録年月日;

                    if (t.M_社員1RowByM_社員1_M_社員1 == null)
                    {
                        g[cEAccount, iX].Value = string.Empty;
                    }
                    else
                    {
                        g[cEAccount, iX].Value = t.M_社員1RowByM_社員1_M_社員1.氏名;
                    }

                    g[cEDate, iX].Value = t.更新年月日;

                    // 2018/09/04
                    for (int i = 0; i < gl.gArrStyle.GetLength(0); i++)
                    {
                        if (t.現場手当有無.ToString() == gl.gArrStyle[i, 0])
                        {
                            g[cGenbaKbn, iX].Value = gl.gArrStyle[i, 1];
                        }
                    }

                    // 2018/09/03
                    g[cKoteiZan, iX].Value = t.固定残業時間.ToString("#,##0");

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

        ///-------------------------------------------------------
        /// <summary>
        ///     画面の初期化 </summary>
        ///-------------------------------------------------------
        private void DispInitial()
        {
            fMode.Mode = global.FORM_ADDMODE;
            txtsName.Text = string.Empty;
            txtCode.Text = string.Empty;
            txtCode.Enabled = true;
            txtName.Text = string.Empty;
            txtFuri.Text = string.Empty;
            txtSzCode.Text = string.Empty;
            txtSzName.Text = string.Empty;
            txtJinTanka.Text = string.Empty;
            cmbZan.SelectedIndex = 0;
            txtTooshiTanka.Text = string.Empty;
            txtKihon10.Text = string.Empty;
            txtPass.Text = string.Empty;
            cmbKintaiEdit.SelectedIndex = 0;
            cmbSysKbn.SelectedIndex = 0;
            txtBikou.Text = string.Empty;
            dtKiten.Checked = false;
            txtKm.Text = string.Empty;
            dtTaishoku.Checked = false;

            cmbGenbateate.SelectedIndex = 0;    // 現場手当有無 2018/09/03
            txtKoteiZan.Text = string.Empty;    // 固定残業時間 2018/09/03

            linkLabel4.Enabled = true;
            linkLabel2.Enabled = false;
            linkLabel3.Enabled = false;
            txtCode.Focus();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
        }

        //登録データチェック
        private Boolean fDataCheck()
        {
            try
            {
                if (fMode.Mode == global.FORM_ADDMODE)
                {
                    // 個人コード
                    if (Utility.StrtoInt(txtCode.Text) == global.flgOff)
                    {
                        txtCode.Focus();
                        throw new Exception("個人コードを入力してください");
                    }

                    // 個人コード登録済みチェック
                    if (dts.M_社員.Any(a => a.ID == Utility.StrtoInt(txtCode.Text)))
                    {
                        txtCode.Focus();
                        throw new Exception("登録済みの個人コードです");
                    }
                }

                // 名称チェック
                if (txtName.Text.Trim().Length < 1)
                {
                    txtName.Focus();
                    throw new Exception("氏名を入力してください");
                }

                // 残業有無区分
                if (cmbZan.SelectedIndex == -1)
                {
                    cmbZan.Focus();
                    throw new Exception("残業有無を選択してください");
                }

                // システムユーザー区分
                if (cmbSysKbn.SelectedIndex == -1)
                {
                    cmbSysKbn.Focus();
                    throw new Exception("システムユーザー区分を選択してください");
                }

                // 勤怠編集権限
                if (cmbKintaiEdit.SelectedIndex == -1)
                {
                    cmbKintaiEdit.Focus();
                    throw new Exception("勤怠編集権限を選択してください");
                }

                // 走行起点日付
                if (Utility.StrtoInt(txtKm.Text) > 0 && !dtKiten.Checked)
                {
                    dtKiten.Focus();
                    throw new Exception("走行起点日付を入力してください");
                }
                
                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, msName + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        ///---------------------------------------------------------
        /// <summary>
        ///     グリッドビュー行選択時処理 </summary>
        ///---------------------------------------------------------
        private void GridEnter()
        {
            string msgStr;
            fMode.rowIndex = dg.SelectedRows[0].Index;

            // 選択確認
            msgStr = "";
            msgStr += dg[0, fMode.rowIndex].Value.ToString() + "：" + dg[1, fMode.rowIndex].Value.ToString() + Environment.NewLine + Environment.NewLine;
            msgStr += "上記の" + msName + "が選択されました。よろしいですか？";

            if (MessageBox.Show(msgStr, "選択", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) == DialogResult.No) 
                return;

            // 対象となるデータテーブルROWを取得します
            ryowaDataSet.M_社員Row sQuery = dts.M_社員.FindByID(int.Parse(dg[0, fMode.rowIndex].Value.ToString()));

            if (sQuery != null)
            {
                // 編集画面に表示
                ShowData(sQuery);

                // モードステータスを「編集モード」にします
                fMode.Mode = global.FORM_EDITMODE;
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
        private void ShowData(ryowaDataSet.M_社員Row s)
        {
            fMode.ID = s.ID.ToString();
            txtCode.Text = s.ID.ToString();
            txtCode.Enabled = false;
            txtName.Text = s.氏名;
            txtFuri.Text = s.フリガナ;
            txtSzCode.Text = s.所属コード.ToString();
            txtSzName.Text = s.所属名;
            txtJinTanka.Text = s.人件費単価.ToString();
            cmbZan.SelectedIndex = s.残業有無;
            txtTooshiTanka.Text = s.通し勤務単価.ToString();
            txtKihon10.Text = s.基本給10.ToString();
            txtPass.Text = s.パスワード;
            cmbSysKbn.SelectedIndex = s.システムユーザー区分;
            cmbKintaiEdit.SelectedIndex = s.アカウント権限;
            txtBikou.Text = s.備考;

            if (s.Is走行起点Null())
            {
                txtKm.Text = "0";
            }
            else
            {
                txtKm.Text = s.走行起点.ToString();
            }

            if (s.Is走行起点日付Null() || s.走行起点日付 == string.Empty)
            {
                dtKiten.Checked = false;
            }
            else
            {
                dtKiten.Value = DateTime.Parse(s.走行起点日付);
            }

            if (s.Is退職年月日Null() || s.退職年月日 == string.Empty)
            {
                dtTaishoku.Checked = false;
            }
            else
            {
                dtTaishoku.Value = DateTime.Parse(s.退職年月日);
            }
        
            cmbGenbateate.SelectedIndex = s.現場手当有無;     // 2018/09/03
            txtKoteiZan.Text = s.固定残業時間.ToString();     // 2018/09/03

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
            adp.Update(dts.M_社員);

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
            txtName.Focus();
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

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //エラーチェック
            if (!fDataCheck()) return;

            switch (fMode.Mode)
            {
                // 新規登録
                case global.FORM_ADDMODE:

                    // 確認
                    if (MessageBox.Show(txtName.Text + "を登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // データセットにデータを追加します
                    var s = dts.M_社員.NewM_社員Row();
                    s.ID = Utility.StrtoInt(txtCode.Text);
                    s.氏名 = txtName.Text;
                    s.フリガナ = txtFuri.Text;
                    s.所属コード = Utility.StrtoInt(txtSzCode.Text);
                    s.所属名 = txtSzName.Text;
                    s.人件費単価 = Utility.StrtoInt(txtJinTanka.Text);
                    s.残業有無 = cmbZan.SelectedIndex;
                    s.通し勤務単価 = Utility.StrtoInt(txtTooshiTanka.Text);
                    s.基本給10 = Utility.StrtoInt(txtKihon10.Text);
                    s.パスワード = txtPass.Text;
                    s.アカウント権限 = cmbKintaiEdit.SelectedIndex;
                    s.システムユーザー区分 = cmbSysKbn.SelectedIndex;
                    s.備考 = txtBikou.Text;

                    s.走行起点 = Utility.StrtoInt(txtKm.Text);

                    if (dtKiten.Checked)
                    {
                        s.走行起点日付 = dtKiten.Value.ToShortDateString();
                    }
                    else
                    {
                        s.走行起点日付 = string.Empty;
                    }

                    if (dtTaishoku.Checked)
                    {
                        s.退職年月日 = dtTaishoku.Value.ToShortDateString();
                    }
                    else
                    {
                        s.退職年月日 = string.Empty;
                    }
                    
                    s.登録ユーザーID = global.loginUserID;
                    s.登録年月日 = DateTime.Now;
                    s.更新ユーザーID = global.loginUserID;
                    s.更新年月日 = DateTime.Now;

                    s.現場手当有無 = cmbGenbateate.SelectedIndex;         // 2018/09/03
                    s.固定残業時間 = Utility.StrtoInt(txtKoteiZan.Text);  // 2018/09/03

                    dts.M_社員.AddM_社員Row(s);
                    break;

                // 更新処理
                case global.FORM_EDITMODE:

                    // 確認
                    if (MessageBox.Show(txtName.Text + "を更新します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // データセット更新
                    var r = dts.M_社員.Single(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached &&
                                               a.ID == int.Parse(fMode.ID));

                    if (!r.HasErrors)
                    {
                        r.氏名 = txtName.Text;
                        r.フリガナ = txtFuri.Text;
                        r.所属コード = Utility.StrtoInt(txtSzCode.Text);
                        r.所属名 = txtSzName.Text;
                        r.人件費単価 = Utility.StrtoInt(txtJinTanka.Text);
                        r.残業有無 = cmbZan.SelectedIndex;
                        r.通し勤務単価 = Utility.StrtoInt(txtTooshiTanka.Text);
                        r.基本給10 = Utility.StrtoInt(txtKihon10.Text);
                        r.パスワード = txtPass.Text;
                        r.アカウント権限 = cmbKintaiEdit.SelectedIndex;
                        r.システムユーザー区分 = cmbSysKbn.SelectedIndex;
                        r.備考 = txtBikou.Text;

                        r.走行起点 = Utility.StrtoInt(txtKm.Text);

                        if (dtKiten.Checked)
                        {
                            r.走行起点日付 = dtKiten.Value.ToShortDateString();
                        }
                        else
                        {
                            r.走行起点日付 = string.Empty;
                        }

                        if (dtTaishoku.Checked)
                        {
                            r.退職年月日 = dtTaishoku.Value.ToShortDateString();
                        }
                        else
                        {
                            r.退職年月日 = string.Empty;
                        }

                        r.更新ユーザーID = global.loginUserID;
                        r.更新年月日 = DateTime.Now;

                        r.現場手当有無 = cmbGenbateate.SelectedIndex;         // 2018/09/03
                        r.固定残業時間 = Utility.StrtoInt(txtKoteiZan.Text);  // 2018/09/03
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
            adp.Update(dts.M_社員);

            // データテーブルにデータを読み込む
            adp.Fill(dts.M_社員);

            // 画面データ消去
            DispInitial();

            // グリッド表示
            GridViewShow(dg, txtsName.Text);
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                // 確認
                if (MessageBox.Show("削除してよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                    return;

                // 削除データ取得（エラー回避のためDataRowState.Deleted と DataRowState.Detachedは除外して抽出する）
                var d = dts.M_社員.Where(a => a.RowState != DataRowState.Deleted && a.RowState != DataRowState.Detached && a.ID == int.Parse(fMode.ID));

                // foreach用の配列を作成する
                var list = d.ToList();

                // 削除
                foreach (var it in list)
                {
                    ryowaDataSet.M_社員Row dl = (ryowaDataSet.M_社員Row)dts.M_社員.Rows.Find(it.ID);
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
                adp.Update(dts.M_社員);

                // データテーブルにデータを読み込む
                adp.Fill(dts.M_社員);

                // 画面データ消去
                DispInitial();

                // グリッド表示
                GridViewShow(dg, txtsName.Text);
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
            MyLibrary.CsvOut.GridView(dg, "社員マスター");
        }

        private void linkLabel6_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // グリッド表示
            GridViewShow(dg, txtsName.Text);
        }
    }
}
