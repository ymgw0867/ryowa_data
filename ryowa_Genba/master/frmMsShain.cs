using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ryowa_Genba.common;

namespace ryowa_Genba.master
{
    public partial class frmMsShain : Form
    {
        // マスター名
        string msName = "社員マスター";

        // フォームモードインスタンス
        Utility.frmMode fMode = new Utility.frmMode();

        // 社員マスターテーブルアダプター生成
        genbaDataSetTableAdapters.M_社員TableAdapter adp = new genbaDataSetTableAdapters.M_社員TableAdapter();

        // データテーブル生成
        genbaDataSet dts = new genbaDataSet();

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

            // 画面初期化
            DispInitial();

            // マスター表示
            ShowData();
        }
                
        ///-------------------------------------------------------
        /// <summary>
        ///     画面の初期化 </summary>
        ///-------------------------------------------------------
        private void DispInitial()
        {
            fMode.Mode = global.FORM_ADDMODE;
            txtCode.Text = string.Empty;
            txtCode.Enabled = true;
            txtName.Text = string.Empty;
            txtPass.Text = string.Empty;
            dtKiten.Checked = false;
            txtKm.Text = string.Empty;

            linkLabel4.Enabled = true;
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
                
                // パスワードチェック
                if (txtPass.Text.Trim().Length < 1)
                {
                    txtPass.Focus();
                    throw new Exception("パスワードを入力してください");
                }

                if (txtPass.Text.Trim().Length < 4)
                {
                    txtPass.Focus();
                    throw new Exception("パスワードは4文字以上で登録してください");
                }

                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, msName + "保守", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }
        
        /// -------------------------------------------------------
        /// <summary>
        ///     マスターの内容を画面に表示する </summary>
        /// <param name="sTemp">
        ///     マスターインスタンス</param>
        /// -------------------------------------------------------
        private void ShowData()
        {
            DateTime dt;
            var m = dts.M_社員;

            if (m.Count > 0)
            {
                foreach (var s in m)
                {
                    fMode.ID = s.ID.ToString();
                    txtCode.Text = s.ID.ToString();
                    txtCode.Enabled = false;
                    txtName.Text = s.氏名;
                    txtPass.Text = s.パスワード;

                    if (DateTime.TryParse(s.走行起点日付, out dt))
                    {
                        dtKiten.Checked = true;
                        dtKiten.Value = dt;
                    }
                    else
                    {
                        dtKiten.Checked = false;
                    }

                    txtKm.Text = s.走行起点.ToString();
                }

                fMode.Mode = global.FORM_EDITMODE;
            }
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
            // エラーチェック
            if (!fDataCheck()) return;

            switch (fMode.Mode)
            {
                // 新規登録
                case global.FORM_ADDMODE:

                    string msg = "";
                    msg += "個人コードは正しく入力されていますか？" + Environment.NewLine;
                    msg += "間違った個人コードが登録されると「どなたの出勤簿データであるか」正しく認識されません。" + Environment.NewLine;
                    msg += "また、正しい個人コードが入れ直すには当システムの再インストールが必要となります。" + Environment.NewLine;
                    msg += "よろしいですか？";

                    if (MessageBox.Show(msg, "個人コード確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // 再確認
                    if (MessageBox.Show(txtName.Text + "を登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // データセットにデータを追加します
                    var s = dts.M_社員.NewM_社員Row();
                    s.ID = Utility.StrtoInt(txtCode.Text);
                    s.氏名 = txtName.Text;
                    s.フリガナ = string.Empty;
                    s.所属コード = 0;
                    s.所属名 = string.Empty;
                    s.人件費単価 = 0;
                    s.残業有無 = 0;
                    s.通し勤務単価 = 0;
                    s.基本給10 = 0;
                    s.パスワード = txtPass.Text;
                    s.アカウント権限 = 0;
                    s.システムユーザー区分 = 0;

                    s.走行起点 = Utility.StrtoInt(txtKm.Text);

                    if (dtKiten.Checked)
                    {
                        s.走行起点日付 = dtKiten.Value.ToShortDateString();
                    }
                    else
                    {
                        s.走行起点日付 = string.Empty;
                    }
                    
                    s.退職年月日 = string.Empty;
                    s.備考 = "";                    
                    s.登録ユーザーID = Utility.StrtoInt(txtCode.Text);
                    s.登録年月日 = DateTime.Now;
                    s.更新ユーザーID = Utility.StrtoInt(txtCode.Text);
                    s.更新年月日 = DateTime.Now;

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
                        r.パスワード = txtPass.Text;
                        r.更新ユーザーID = Utility.StrtoInt(txtCode.Text);
                        r.更新年月日 = DateTime.Now;

                        // 2018/10/22
                        r.走行起点 = Utility.StrtoInt(txtKm.Text);

                        if (dtKiten.Checked)
                        {
                            r.走行起点日付 = dtKiten.Value.ToShortDateString();
                        }
                        else
                        {
                            r.走行起点日付 = string.Empty;
                        }
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

            // 終了
            this.Close();
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
        
    }
}
