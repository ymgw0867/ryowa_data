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
using Excel = Microsoft.Office.Interop.Excel;

namespace ryowa_Confir.data
{
    public partial class frmAuthSend : Form
    {
        public frmAuthSend()
        {
            InitializeComponent();
        }

        //カラム定義
        string cDel = "col1";
        string cFileName = "col2";

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
                g.Height = 198;

                // 奇数行の色
                g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.UseColumnTextForButtonValue = true;
                btn.Text = "削除";
                g.Columns.Add(btn);
                g.Columns[0].Name = cDel;
                g.Columns[cDel].HeaderText = "";

                g.Columns.Add(cFileName, "ファイル名");

                g.Columns[cDel].Width = 50;
                g.Columns[cFileName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

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

        private void gridShow(string fName)
        {
            dg.Rows.Add();
            dg[cFileName, dg.RowCount - 1].Value = fName;

            dg.CurrentCell = null;
            linkLabel4.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // インポートデータ選択
            string fName = openCsvFile();
            if (fName == string.Empty)
            {
                return;
            }

            // 追加済みチェック
            //lblFname.Text = fName;
                
            for (int i = 0; i < dg.RowCount; i++)
            {
                if (dg[cFileName, i].Value.ToString() == fName)
                {
                    button1.Focus();
                    MessageBox.Show("既に追加済みです","",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                    button1.Focus();
                    return;
                }
            }

            // グリッドに追加
            gridShow(fName);
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

            openFileDialog1.Title = "確認送信対象のエクセル出勤簿の選択";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.FileName = "";
            openFileDialog1.Filter = "Excelファイル(*.xls;*.xlsx)|*.xls;*.xlsx";

            //ダイアログボックスの表示
            ret = openFileDialog1.ShowDialog();
            if (ret == System.Windows.Forms.DialogResult.Cancel)
            {
                return string.Empty;
            }

            if (MessageBox.Show(openFileDialog1.FileName + Environment.NewLine + " が選択されました。よろしいですか?", "ファイル確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return string.Empty;
            }

            return openFileDialog1.FileName;
        }

        private void frmAuthSend_Load(object sender, EventArgs e)
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
            //lblFname.Text = string.Empty;
            linkLabel4.Enabled = false;
            button1.Focus();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //gridShow(fName);
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

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 閉じる
            Close();
        }

        private void frmAuthSend_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        ///------------------------------------------------------------
        /// <summary>
        ///     出勤簿確認送信メール作成 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        ///------------------------------------------------------------
        private void putCsvFile(DataGridView g)
        {
            for (int i = 0; i < g.RowCount; i++)
            {
                string outCsvFileName = exlToCsv(g[1, i].Value.ToString());

                if (outCsvFileName != string.Empty)
                {
                    // 確認送信
                    if (System.IO.File.Exists(outCsvFileName))
                    {
                        // メール件名
                        string sbj = "<" + DateTime.Today.ToShortDateString() + "> " + "TimeCards Checked";

                        // メール本文
                        string sBody = "出勤簿確認送信メール";

                        // 送信
                        Utility.sendKintaiMail(outCsvFileName, sbj, sBody, "出勤簿確認送信メール", 0);
                    }
                }
            }
        }


        ///------------------------------------------------------------------
        /// <summary>
        ///     出勤簿・車両走行報告書作成 </summary>
        ///------------------------------------------------------------------
        private string exlToCsv(string exlName)
        {
            string xlsName = string.Empty;

            // 該当ファイルがないときは終わる
            if (!System.IO.File.Exists(exlName))
            {
                return string.Empty;
            }

            //エクセルファイル日付明細開始行
            const int S_GYO = 5;
            const int E_GYO = 35;

            int eRow = 0;

            try
            {
                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                // 勤務報告書テンプレートシート
                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(exlName,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                Excel.Worksheet oxlsPrintSheet = null;    // 印刷用ワークシート

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                string Category = string.Empty;

                string[] xlsArray = null;

                StringBuilder sb = new StringBuilder();

                try
                {
                    // カレントのシートを設定
                    oxlsPrintSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                    // 年
                    rng[0] = (Excel.Range)oxlsPrintSheet.Cells[2, 6];
                    //int sYear = Utility.StrtoInt(rng[0].Text) + Properties.Settings.Default.rekiHosei;
                    int sYear = Utility.StrtoInt(rng[0].Text);  // 和暦から西暦へ 2018/07/13

                    // 月
                    rng[0] = (Excel.Range)oxlsPrintSheet.Cells[2, 8];
                    string sMonth = rng[0].Text;

                    // 個人コード
                    rng[0] = (Excel.Range)oxlsPrintSheet.Cells[3, 4];
                    string sNum = rng[0].Text;

                    int iX = 0;

                    for (int i = S_GYO; i <= E_GYO; i++)
                    {
                        // 工事
                        rng[0] = (Excel.Range)oxlsPrintSheet.Cells[i, 2];
                        string sKouji = rng[0].Text;

                        // 出勤
                        rng[0] = (Excel.Range)oxlsPrintSheet.Cells[i, 3];
                        string sWork = rng[0].Text;

                        // 出勤簿入力されていない日はネグる
                        if (sKouji == string.Empty || sWork == string.Empty)
                        {
                            continue;
                        }
                        
                        // 日
                        rng[0] = (Excel.Range)oxlsPrintSheet.Cells[i, 1];
                        string sDay = rng[0].Text;

                        string chk = "0";
                        rng[0] = (Excel.Range)oxlsPrintSheet.Cells[i, 17];

                        if (rng[0].Text != null && rng[0].Text != string.Empty)
                        {
                            //chk = "1";

                            // チェック欄に管理者の個人コードを入力してもらい取得する
                            if (Utility.StrtoInt(rng[0].Text) != 0)
                            {
                                chk = (Utility.StrtoInt(rng[0].Text)).ToString();
                            }
                        }

                        sb.Clear();
                        sb.Append(sNum).Append(",");
                        sb.Append(sYear + "/" + sMonth + "/" + sDay).Append(",");
                        sb.Append(chk);
                    
                        Array.Resize(ref xlsArray, iX + 1);
                        xlsArray[iX] = sb.ToString();
                        iX++;
                    }

                    // 確認のためのウィンドウを表示する
                    //oXls.Visible = true;

                    //印刷
                    //oXlsBook.PrintPreview(true);
                    //oXlsBook.PrintOut();

                    //保存処理
                    oXls.DisplayAlerts = false;

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();

                    // 添付ファイルパス名
                    string outFileName = Properties.Settings.Default.attachPath + sNum + " 確認送信.csv";
                    
                    // 添付ファイルフォルダー内のファイルをすべて削除する
                    foreach (var file in System.IO.Directory.GetFiles(Properties.Settings.Default.attachPath))
                    {
                        System.IO.File.Delete(file);
                    }

                    // CSVファイル出力
                    System.IO.File.WriteAllLines(outFileName, xlsArray, System.Text.Encoding.GetEncoding("utf-8"));

                    // パスを含むCSVファイル名を返す
                    return outFileName;
                }

                catch (Exception e)
                {
                    //MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                    // ウィンドウを非表示にする
                    oXls.Visible = false;

                    //保存処理
                    oXls.DisplayAlerts = false;

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
                }

                finally
                {
                    // COM オブジェクトの参照カウントを解放する
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsPrintSheet);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            return xlsName;
        }





        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 送信確認メッセージ
            if (dg.RowCount < 1)
            {
                MessageBox.Show("確認送信する出勤簿エクセルファイルが追加されていません", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // 送信確認メッセージ
            if (MessageBox.Show("出勤簿の確認送信を行います。よろしいですか？", "登録確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 確認送信
            putCsvFile(dg);

            // 閉じる
            Close();
        }
    }
}
