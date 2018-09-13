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

namespace ryowa_DATA.data
{
    public partial class frmKintaiCheckList : Form
    {
        public frmKintaiCheckList()
        {
            InitializeComponent();

        }

        #region カラム定義
        string cEdit = "col0";
        string cID = "col1";
        string cName = "col2";
        string cZanTM = "col3";
        string cSiTM = "col4";
        string cHolWorkTime = "col5";
        string cHouteiTm = "col6";
        string cHolDTime = "col7";
        string cHouteiDTm = "col8";
        string cKinmuchiT = "col9";
        string cSonotaT = "col10";
        string cGenbaNum = "col11";
        string cInStay = "col12";
        string cOutStay = "col13";
        string cEnkakuchi = "col14";
        string cJyo = "col15";
        string cTokushu = "col16";
        string cTooshi = "col17";
        string cYakan = "col18";
        string cShokumu = "col19";
        #endregion

        // 勤怠チェックリストデータクラス
        clsKintaiData kd;

        private void frmKintaiCheckList_Load(object sender, EventArgs e)
        {
            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッド定義
            GridViewSetting(dg, global.loginType);

            // 画面初期化
            dispInitial();
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///-------------------------------------------------------------------
        private void GridViewSetting(DataGridView g, int lgType)
        {
            try
            {
                g.EnableHeadersVisualStyles = false;
                g.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
                g.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;

                g.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                g.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.TopCenter;

                // 列ヘッダーフォント指定
                g.ColumnHeadersDefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);

                // データフォント指定
                g.DefaultCellStyle.Font = new Font("Meiryo UI", (float)9.5, FontStyle.Regular);

                // 行の高さ
                g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                g.ColumnHeadersHeight = 38;
                g.RowTemplate.Height = 22;

                // 全体の高さ
                g.Height = 582;

                // 奇数行の色
                g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                //DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                //btn.UseColumnTextForButtonValue = false;
                //btn.Text = "登録";
                //g.Columns.Add(btn);
                //g.Columns[0].Name = cEdit;
                //g.Columns[0].HeaderText = "";

                g.Columns.Add(cID, "コード");
                g.Columns.Add(cName, "氏名");
                g.Columns.Add(cZanTM, "普通残業");
                g.Columns.Add(cSiTM, "深夜残業");
                g.Columns.Add(cHolWorkTime, "休日勤務時間");
                g.Columns.Add(cHouteiTm, "法定休日勤務時間");
                g.Columns.Add(cHolDTime, "休日代休時間");
                g.Columns.Add(cHouteiDTm, "法定休日代休時間");
                g.Columns.Add(cKinmuchiT, "勤務地手当");
                g.Columns.Add(cSonotaT, "その他支給２");
                g.Columns.Add(cGenbaNum, "現場回数");
                g.Columns.Add(cInStay, "県内宿泊回数");
                g.Columns.Add(cOutStay, "県外宿泊回数");
                g.Columns.Add(cEnkakuchi, "遠隔地回数");
                g.Columns.Add(cJyo, "除雪当番");
                g.Columns.Add(cTokushu, "特殊勤務");
                g.Columns.Add(cTooshi, "通し勤務");
                g.Columns.Add(cYakan, "夜間手当");
                g.Columns.Add(cShokumu, "職務手当");

                g.Columns[cID].Width = 70;
                g.Columns[cName].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                g.Columns[cZanTM].Width = 60;
                g.Columns[cSiTM].Width = 60;
                g.Columns[cHolWorkTime].Width = 60;
                g.Columns[cHouteiTm].Width = 60;
                g.Columns[cHolDTime].Width = 60;
                g.Columns[cHouteiDTm].Width = 60;
                g.Columns[cKinmuchiT].Width = 70;
                g.Columns[cSonotaT].Width = 70;
                g.Columns[cGenbaNum].Width = 50;
                g.Columns[cInStay].Width = 50;
                g.Columns[cOutStay].Width = 50;
                g.Columns[cEnkakuchi].Width = 50;
                g.Columns[cJyo].Width = 50;
                g.Columns[cTokushu].Width = 50;
                g.Columns[cTooshi].Width = 50;
                g.Columns[cYakan].Width = 50;
                g.Columns[cShokumu].Width = 50;

                g.Columns[cID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cZanTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cSiTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cHolWorkTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cHouteiTm].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cHolDTime].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cHouteiDTm].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cKinmuchiT].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cSonotaT].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cGenbaNum].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cInStay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cOutStay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cEnkakuchi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                g.Columns[cJyo].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cTokushu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cTooshi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cYakan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cShokumu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                
                //g.Columns[cName].Frozen = true;

                // 行ヘッダを表示しない
                g.RowHeadersVisible = false;

                // 選択モード
                g.SelectionMode = DataGridViewSelectionMode.CellSelect;
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
                //g.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     画面初期化  </summary>
        ///---------------------------------------------------------------------
        private void dispInitial()
        {
            linkLabel1.Enabled = false;
            linkLabel2.Enabled = false;

            //txtYear.Text = (DateTime.Today.Year - Properties.Settings.Default.rekiHosei).ToString();
            txtYear.Text = DateTime.Today.Year.ToString();    // 和暦から西暦へ 2018/07/13
            txtMonth.Text = DateTime.Today.Month.ToString();

            //label6.Text = Properties.Settings.Default.gengou;  // 和暦から西暦へ コメント化 2018/07/18
        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            // 年
            if ((Utility.StrtoInt(txtYear.Text) < 20))
            {
                dg.Rows.Clear();
                return;
            }

            // 月
            if ((Utility.StrtoInt(txtMonth.Text) < 1) || (Utility.StrtoInt(txtMonth.Text) > 12))
            {
                dg.Rows.Clear();
                return;
            }

            //int sYY = Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei;
            int sYY = Utility.StrtoInt(txtYear.Text);   // 和暦から西暦へ
            int sMM = Utility.StrtoInt(txtMonth.Text);

            // 勤怠データ表示
            gridDataShow(sYY, sMM, dg);
        }

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     チェックリスト表示 </summary>
        /// <param name="sYY">
        ///     対象年</param>
        /// <param name="sMM">
        ///     対象月</param>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        ///-----------------------------------------------------------------------------
        private void gridDataShow(int sYY, int sMM, DataGridView g)
        {
            g.Rows.Clear();

            // 勤怠データ集計クラス
            kd = new clsKintaiData(sYY, sMM);

            if (kd.wd == null)
            {
                linkLabel1.Enabled = false;
                linkLabel2.Enabled = false;
                return;
            }
            
            for (int i = 0; i < kd.wd.Length; i++)
            {
                g.Rows.Add();

                g[cID, i].Value = kd.wd[i].code;
                g[cName, i].Value = kd.wd[i].name;
                g[cZanTM, i].Value = Utility.intToHhMM10(kd.wd[i].zanTM, Properties.Settings.Default.marume);
                g[cSiTM, i].Value = Utility.intToHhMM10(kd.wd[i].siTM, Properties.Settings.Default.marume);
                g[cHolWorkTime, i].Value = Utility.intToHhMM(kd.wd[i].holTM);
                g[cHouteiTm, i].Value = Utility.intToHhMM(kd.wd[i].houteiTM);
                g[cHolDTime, i].Value = Utility.intToHhMM(kd.wd[i].holDaikyuTM);
                g[cHouteiDTm, i].Value = Utility.intToHhMM(kd.wd[i].houteiDaikyuTM);
                g[cKinmuchiT, i].Value = kd.wd[i].kinmuchiTe.ToString("#,##0");
                g[cSonotaT, i].Value = kd.wd[i].sonota2.ToString("#,##0");
                g[cGenbaNum, i].Value = kd.wd[i].genbaNum;
                g[cInStay, i].Value = kd.wd[i].inStayNum;
                g[cOutStay, i].Value = kd.wd[i].outStayNum;
                g[cEnkakuchi, i].Value = kd.wd[i].enkakuchiNum;
                g[cJyo, i].Value = kd.wd[i].jyosekiNum;
                g[cTokushu, i].Value = kd.wd[i].tokushuNum;
                g[cTooshi, i].Value = kd.wd[i].tooshiNum;
                g[cYakan, i].Value = kd.wd[i].yakanNum;
                g[cShokumu, i].Value = kd.wd[i].shokumuNum;
            }

            if (g.Rows.Count > 0)
            {
                g.CurrentCell = null;

                linkLabel1.Enabled = true;
                linkLabel2.Enabled = true;
            }
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MyLibrary.CsvOut.GridView(dg, "勤怠チェックリスト");
        }

        private void frmKintaiCheckList_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("給与大臣用ＣＳＶデータを作成しますか","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // ＣＳＶデータ作成
            kd.csvOutput();
        }
    }
}
