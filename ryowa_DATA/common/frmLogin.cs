using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ryowa_DATA.common
{
    public partial class frmLogin : Form
    {
        public frmLogin()
        {
            InitializeComponent();

            // ログインステータス
            global.loginStatus = false;

            // ログインユーザーデータ読み込み
            adp.Fill(dts.M_社員1);
        }

        ryowaDataSet dts = new ryowaDataSet();
        ryowaDataSetTableAdapters.M_社員1TableAdapter adp = new ryowaDataSetTableAdapters.M_社員1TableAdapter();

        // ログイントライ回数
        int lTry = 0;

        // 個人コード＆パスワード記憶ファイル：2018/09/05
        string ID_PASS = "idpass.csv"; 

        private void button2_Click(object sender, EventArgs e)
        {
        }

        private void frmLogin_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 個人コードとパスワードを記憶する
            idPassPut(checkBox1.Checked);

            // 後片付け
            this.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        /// ---------------------------------------------------------
        /// <summary>
        ///     ユーザー認証 </summary>
        /// <returns>
        ///     認証成功：true、認証失敗：false</returns>
        /// ---------------------------------------------------------
        private bool SerachloginUser()
        {
            // ログインユーザー
            if (txtName.Text == string.Empty)
            {
                MessageBox.Show("ログインユーザー名を入力してください", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                lTry++;
                txtName.Focus();
                return false;
            }

            // パスワード
            if (txtPassword.Text == string.Empty)
            {
                MessageBox.Show("パスワードを入力してください", "エラー", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                lTry++;
                txtPassword.Focus();
                return false;
            }

            // admin認証
            if (txtName.Text == Properties.Settings.Default.login && 
                txtPassword.Text == Properties.Settings.Default.password)
            {
                global.loginUserID = 0;
                global.loginSysUser = 1;
                global.loginType = 1;
                return true;
            }

            // ユーザー認証
            if (!dts.M_社員1.Any(a => a.ID.ToString() == txtName.Text && a.パスワード == txtPassword.Text))
            {
                lTry++;
                MessageBox.Show("ユーザー認証に失敗しました。" + Environment.NewLine + "ログインユーザー名とパスワードを確認してください。", "認証エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtName.Focus();
                return false;
            }
            else
            {
                var s = dts.M_社員1.Single(a => a.ID.ToString() == txtName.Text && a.パスワード == txtPassword.Text);
                global.loginUserID = s.ID;
                global.loginType = s.アカウント権限;
                global.loginSysUser = s.システムユーザー区分;
            }

            // 認証成功
            return true;
        }

        private void txtName_Enter(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;
            txtObj.SelectAll();
            //txtObj.BackColor = SystemColors.ControlLight;
        }

        private void txtPassword_Leave(object sender, EventArgs e)
        {
            TextBox txtObj = (TextBox)sender;
            //txtObj.BackColor = Color.White;
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 認証結果検証
            if (!SerachloginUser())
            {
                // 3回ログインエラーでシステムは終了します
                if (lTry > 2)
                {
                    MessageBox.Show("3回続けてログインに失敗しました。システムを終了します。", "中止", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Environment.Exit(0);
                }
            }
            else
            {
                global.loginStatus = true;
                this.Close();
            }
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     個人コードとパスワードを記憶する </summary>
        /// <param name="isPut">
        ///     true:記憶する、false:記憶しない</param>
        ///-----------------------------------------------------------
        private void idPassPut(bool isPut)
        {
            // ファイルを削除
            if (System.IO.File.Exists(@"c:\ryowa_data\xls\" + ID_PASS))
            {
                System.IO.File.Delete(@"c:\ryowa_data\xls\" + ID_PASS);
            }

            // CSVファイル出力
            if (isPut)
            {
                System.IO.File.WriteAllText(@"c:\ryowa_data\xls\" + ID_PASS, txtName.Text + "," + txtPassword.Text, System.Text.Encoding.GetEncoding("utf-8"));

                // ID_PASSの属性を取得する
                System.IO.FileAttributes attr = System.IO.File.GetAttributes(@"c:\ryowa_data\xls\" + ID_PASS);
                
                // 隠しファイル属性を追加する
                System.IO.File.SetAttributes(@"c:\ryowa_data\xls\" + ID_PASS,
                    attr | System.IO.FileAttributes.Hidden);
            }
        }


        private void frmLogin_Load(object sender, EventArgs e)
        {
        }

        private void frmLogin_Shown(object sender, EventArgs e)
        {
            txtName.Text = string.Empty;
            txtPassword.Text = string.Empty;
            checkBox1.Checked = false;

            if (System.IO.File.Exists(@"c:\ryowa_data\xls\" + ID_PASS))
            {
                foreach (var item in System.IO.File.ReadLines(@"c:\ryowa_data\xls\" + ID_PASS))
                {
                    string[] p = item.Split(',');
                    if (p.Length == 2)
                    {
                        txtName.Text = p[0];
                        txtPassword.Text = p[1];
                        checkBox1.Checked = true;
                        linkLabel4.Focus();
                    }

                    break;
                }
            }
        }
    }
}
