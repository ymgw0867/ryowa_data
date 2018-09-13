using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ryowa_Confir.common;

namespace ryowa_Confir.data
{
    public partial class frmRequest : Form
    {
        public frmRequest()
        {
            InitializeComponent();

            // メール設定読み込み
            mAdp.Fill(dts.メール設定);
        }

        confirDataSet dts = new confirDataSet();
        confirDataSetTableAdapters.メール設定TableAdapter mAdp = new confirDataSetTableAdapters.メール設定TableAdapter();

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        //カラム定義
        string cDel = "col1";
        string cYear = "col2";
        string cMonth = "col3";
        string cCode = "col4";

        ///-----------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///-----------------------------------------------------------------
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
                g.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // データフォント指定
                g.DefaultCellStyle.Font = new Font("Meiryo UI", 10, FontStyle.Regular);

                // 行の高さ
                g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                g.ColumnHeadersHeight = 20;
                g.RowTemplate.Height = 22;

                // 全体の高さ
                g.Height = 155;

                // 奇数行の色
                g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.UseColumnTextForButtonValue = true;
                btn.Text = "削除";
                g.Columns.Add(btn);
                g.Columns[0].Name = cDel;
                g.Columns[cDel].HeaderText = "";

                g.Columns.Add(cYear, "年");
                g.Columns.Add(cMonth, "月");
                g.Columns.Add(cCode, "個人コード");

                g.Columns[cDel].Width = 50;
                g.Columns[cYear].Width = 60;
                g.Columns[cMonth].Width = 60;
                g.Columns[cCode].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                g.Columns[cYear].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cMonth].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cCode].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Close();
        }

        private void frmRequest_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (fDataCheck())
            {
                gridShow();
                txtNum.Text = string.Empty;
                txtNum.Focus();
            }
        }

        private void gridShow()
        {
            dg.Rows.Add();
            dg[cYear, dg.RowCount - 1].Value = txtYear.Text;
            dg[cMonth, dg.RowCount - 1].Value = txtMonth.Text;
            dg[cCode, dg.RowCount - 1].Value = txtNum.Text;

            dg.CurrentCell = null;
            linkLabel4.Enabled = true;
        }
        
        // 登録データチェック
        private Boolean fDataCheck()
        {
            try
            {
                if (Utility.StrtoInt(txtYear.Text) == global.flgOff)
                {
                    txtYear.Focus();
                    throw new Exception("年を西暦で入力してください");
                }

                if (Utility.StrtoInt(txtYear.Text) < 2016)
                {
                    txtYear.Focus();
                    throw new Exception("年は2016年以降で指定してください");
                }

                if (Utility.StrtoInt(txtMonth.Text) < 1 || Utility.StrtoInt(txtMonth.Text) > 12)
                {
                    txtMonth.Focus();
                    throw new Exception("月が正しくありません");
                }
                
                if (Utility.StrtoInt(txtNum.Text) == global.flgOff)
                {
                    txtNum.Focus();
                    throw new Exception("個人コードを入力してください");
                }

                for (int i = 0; i < dg.RowCount; i++)
                {
                    if (dg[cYear, i].Value.ToString() == txtYear.Text &&
                        dg[cMonth, i].Value.ToString() == txtMonth.Text &&
                        dg[cCode, i].Value.ToString() == txtNum.Text)
                    {
                        txtNum.Focus();
                        throw new Exception("既に追加済みです");
                    }
                }              

                return true;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "出勤簿リクエスト", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }
        }

        private void frmRequest_Load(object sender, EventArgs e)
        {
            // フォーム最大サイズ
            Utility.WindowsMaxSize(this, this.Width, this.Height);

            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // グリッドビュー定義
            GridViewSetting(dg);

            // 画面初期化
            DispInitial();
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     画面の初期化 </summary>
        ///-------------------------------------------------------
        private void DispInitial()
        {
            txtYear.Text = DateTime.Today.Year.ToString();
            txtMonth.Text = DateTime.Today.Month.ToString();
            txtNum.Text = string.Empty;

            linkLabel4.Enabled = false;
            txtYear.Focus();
        }

        private void dg_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                if (e.ColumnIndex == 0)
                {
                    // 行削除 
                    dg.Rows.RemoveAt(e.RowIndex);

                    // 全行削除？
                    if (dg.RowCount == 0)
                    {
                        linkLabel4.Enabled = false;
                    }
                    else
                    {
                        dg.CurrentCell = null;
                    }
                }
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     出勤簿リクエスト用CSVデータ作成 </summary>
        /// <param name="g">
        ///     DataGridViewオブジェクト</param>
        /// <param name="mlAdd">
        ///     送信元メールアドレス</param>
        /// <returns>
        ///     パスを含むCSVファイル名</returns>
        ///---------------------------------------------------------------------
        private string putKintaiCsv(DataGridView g, string mlAdd)
        {
            StringBuilder sb = new StringBuilder();

            string[] st = new string[g.RowCount];

            for (int i = 0; i < g.RowCount; i++)
            {
                sb.Clear();
                sb.Append(g[cYear, i].Value.ToString()).Append(",");
                sb.Append(g[cMonth, i].Value.ToString()).Append(",");
                sb.Append(g[cCode, i].Value.ToString()).Append(",");
                sb.Append(mlAdd);
                st[i] = sb.ToString();
            }
            
            // 添付ファイルパス名
            string outFileName = Properties.Settings.Default.attachPath + "出勤簿要求.csv";

            // 添付ファイルフォルダー内のファイルをすべて削除する
            foreach (var file in System.IO.Directory.GetFiles(Properties.Settings.Default.attachPath))
            {
                System.IO.File.Delete(file);
            }

            // CSVファイル出力
            System.IO.File.WriteAllLines(outFileName, st, System.Text.Encoding.GetEncoding("utf-8"));

            // パスを含むCSVファイル名を返す
            return outFileName;
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 送信確認メッセージ
            if (dg.RowCount < 1)
            {
                MessageBox.Show("要求する出勤簿が追加されていません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 送信確認メッセージ
            if (MessageBox.Show("出勤簿要求メールを送信します。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 自身のメールアドレスを取得
            string mlAdd = string.Empty;
            if (dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                var ml = dts.メール設定.Single(a => a.ID == global.mailKey);
                mlAdd = ml.メールアドレス;
            }

            string outCsvFileName = putKintaiCsv(dg, mlAdd);

            // 要求メール送信
            if (System.IO.File.Exists(outCsvFileName))
            {
                // メール件名
                string sbj = "<" + DateTime.Today.ToShortDateString() + "> " + "TimeCards Request";

                // メール本文
                string sBody = "出勤簿・車両走行報告書要求メール";

                // 送信
                Utility.sendKintaiMail(outCsvFileName, sbj, sBody, "出勤簿・車両走行報告書要求メール", 0);
            }

            // 閉じる
            Close();
        }
    }
}
