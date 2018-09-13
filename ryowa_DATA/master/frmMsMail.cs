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
    public partial class frmMsMail : Form
    {
        // マスター名
        string msName = "メール設定";

        // フォームモードインスタンス
        Utility.frmMode fMode = new Utility.frmMode();

        // メール設定テーブルアダプター生成
        ryowaDataSetTableAdapters.メール設定TableAdapter adp = new ryowaDataSetTableAdapters.メール設定TableAdapter();

        // データテーブル生成
        ryowaDataSet dts = new ryowaDataSet();

        public frmMsMail()
        {
            InitializeComponent();

            // データテーブルにデータを読み込む
            adp.Fill(dts.メール設定);
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
            txtMailAddress.Text = string.Empty;
            txtName.Text = string.Empty;
            txtAccount.Text = string.Empty;
            txtPass.Text = string.Empty;
            txtSmtpServer.Text = string.Empty;
            txtSmtpPort.Text = string.Empty;
            txtPopServer.Text = string.Empty;
            txtPopPort.Text = string.Empty;

            linkLabel4.Enabled = true;
            txtMailAddress.Focus();
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
        }

        //登録データチェック
        private Boolean fDataCheck()
        {
            try
            {
                // メールアドレスチェック
                if (txtMailAddress.Text.Trim().Length < 1)
                {
                    txtMailAddress.Focus();
                    throw new Exception("メールアドレスを入力してください");
                }

                // 名称チェック
                if (txtName.Text.Trim().Length < 1)
                {
                    txtName.Focus();
                    throw new Exception("メールアドレス名称を入力してください");
                }

                // アカウントチェック
                if (txtAccount.Text.Trim().Length < 1)
                {
                    txtAccount.Focus();
                    throw new Exception("メールアカウントを入力してください");
                }

                // パスワードチェック
                if (txtPass.Text.Trim().Length < 1)
                {
                    txtPass.Focus();
                    throw new Exception("メールパスワードを入力してください");
                }

                // SMTPサーバー
                if (txtSmtpServer.Text.Trim().Length < 1)
                {
                    txtSmtpServer.Focus();
                    throw new Exception("SMTPサーバー名を入力してください");
                }

                // SMTPポート番号
                if (txtSmtpPort.Text.Trim().Length < 1)
                {
                    txtSmtpPort.Focus();
                    throw new Exception("SMTPポート番号を入力してください");
                }

                // POPサーバー
                if (txtPopServer.Text.Trim().Length < 1)
                {
                    txtPopServer.Focus();
                    throw new Exception("POPサーバー名を入力してください");
                }

                // POPポート番号
                if (txtPopPort.Text.Trim().Length < 1)
                {
                    txtPopPort.Focus();
                    throw new Exception("POPポート番号を入力してください");
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
            if (dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                var s = dts.メール設定.Single(a => a.ID == global.mailKey);

                txtMailAddress.Text = s.メールアドレス;
                txtName.Text = s.メール名称;
                txtAccount.Text = s.ログイン名;
                txtPass.Text = s.パスワード;
                txtSmtpServer.Text = s.SMTPサーバー;
                txtSmtpPort.Text = s.SMTPポート番号.ToString();
                txtPopServer.Text = s.POPサーバー;
                txtPopPort.Text = s.POPポート番号.ToString();

                fMode.Mode = global.FORM_EDITMODE;
            }
        }
        
        private void btnRtn_Click(object sender, EventArgs e)
        {
        }

        private void frm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // データセットの内容をデータベースへ反映させます
            adp.Update(dts.メール設定);

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
                    if (MessageBox.Show("登録します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // データセットにデータを追加します
                    var s = dts.メール設定.Newメール設定Row();

                    s.ID = global.mailKey;
                    s.SMTPサーバー = txtSmtpServer.Text;
                    s.SMTPポート番号 = Utility.StrtoInt(txtSmtpPort.Text);
                    s.ログイン名 = txtAccount.Text;
                    s.パスワード = txtPass.Text;
                    s.メールアドレス = txtMailAddress.Text;
                    s.メール名称 = txtName.Text;
                    s.送信先アドレス = string.Empty;
                    s.POPサーバー = txtPopServer.Text;
                    s.POPポート番号 = Utility.StrtoInt(txtPopPort.Text);

                    dts.メール設定.Addメール設定Row(s);
                    break;

                // 更新処理
                case global.FORM_EDITMODE:

                    // 確認
                    if (MessageBox.Show("更新します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No)
                        return;

                    // データセット更新
                    var r = dts.メール設定.Single(a => a.ID == global.mailKey);

                    if (!r.HasErrors)
                    {
                        r.SMTPサーバー = txtSmtpServer.Text;
                        r.SMTPポート番号 = Utility.StrtoInt(txtSmtpPort.Text);
                        r.ログイン名 = txtAccount.Text;
                        r.パスワード = txtPass.Text;
                        r.メールアドレス = txtMailAddress.Text;
                        r.メール名称 = txtName.Text;
                        r.POPサーバー = txtPopServer.Text;
                        r.POPポート番号 = Utility.StrtoInt(txtPopPort.Text);
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
            adp.Update(dts.メール設定);

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

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }
    }
}
