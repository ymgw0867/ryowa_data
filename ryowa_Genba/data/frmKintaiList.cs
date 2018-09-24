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
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace ryowa_Genba.data
{
    public partial class frmKintaiList : Form
    {
        public frmKintaiList()
        {
            InitializeComponent();

            // データ読み込み
            aAdp.Fill(dts.M_社員);
            hAdp.Fill(dts.M_休日);
            kAdp.Fill(dts.M_工事);
            mAdp.Fill(dts.メール設定);

            // T_勤怠テーブル 早出残業フィールド追加 2018/07/12
            dbCreateAlter();

            //// 早出残業null → 0 : 2018/07/12
            bool upStatus = false;
            adp.FillByHayadeNull(dts.T_勤怠);
            foreach (var item in dts.T_勤怠)
            {
                item.早出残業 = 0;
                upStatus = true;
            }

            if (upStatus)
            {
                adp.Update(dts.T_勤怠);
            }
        }

        genbaDataSet dts = new genbaDataSet();
        genbaDataSetTableAdapters.T_勤怠TableAdapter adp = new genbaDataSetTableAdapters.T_勤怠TableAdapter();
        genbaDataSetTableAdapters.M_社員TableAdapter aAdp = new genbaDataSetTableAdapters.M_社員TableAdapter();
        genbaDataSetTableAdapters.M_休日TableAdapter hAdp = new genbaDataSetTableAdapters.M_休日TableAdapter();
        genbaDataSetTableAdapters.M_工事TableAdapter kAdp = new genbaDataSetTableAdapters.M_工事TableAdapter();
        genbaDataSetTableAdapters.メール設定TableAdapter mAdp = new genbaDataSetTableAdapters.メール設定TableAdapter();

        bool changeStatus = false;
        DateTime taDt;  // 退職年月日

        double fixZan = 0;  // 社員固定残業時間

        private void frmKintaiList_Load(object sender, EventArgs e)
        {
            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッド定義
            GridViewSetting(dg, global.loginType);
            GridViewSetting2(dg2, global.loginType);

            // 画面初期化
            dispInitial();
        }
        
        #region カラム定義
        string cEdit = "col0";
        string cDay = "col1";
        string cDayMemo = "col27";
        string cName = "col2";
        string cKinmuKbn = "col3";
        string cInTM = "col4";
        string cSTM = "col5";
        string cETM = "col6";
        string cOutTM = "col7";
        string cRestTM = "col8";
        string cZanTM = "col9";
        string cSiTM = "col10";
        string cStay = "col11";
        string cMemo = "col12";
        string cAllKm = "col13";
        string cKmTuukin = "col14";
        string cKmShiyou = "col15";
        string cJyosetsu = "col16";
        string cTokushu = "col17";
        string cTooshi = "col18";
        string cYakan = "col19";
        string cShokumu = "col20";
        string cKakunin = "col21";
        string cAccount = "col22";
        string cDate = "col23";
        string cEAccount = "col24";
        string cEDate = "col25";
        string cCode = "col26";
        string cHayade = "col44";       // 2018/07/12 早出

        string cName2 = "col42";
        string cChiiki2 = "col27";
        string cShukkin2 = "col28";
        string cDaikyu2 = "col29";
        string cHolWorkTime2 = "col30";
        string cHouteiTm2 = "col31";
        string cZan2 = "col32";
        string cShinya2 = "col33";
        string cStay2 = "col34";
        string cKmTuukin2 = "col35";
        string cKmShiyou2 = "col36";
        string cJyosetsu2 = "col37";
        string cTokushu2 = "col38";
        string cTooshi2 = "col39";
        string cYakan2 = "col40";
        string cShokumu2 = "col41";
        string cCode2 = "col43";
        string cHayade2 = "col45";       // 2018/07/12 早出
        #endregion

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
                g.Height = 498;

                // 奇数行の色
                //g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                btn.UseColumnTextForButtonValue = false;
                btn.Text = "登録";
                g.Columns.Add(btn);
                g.Columns[0].Name = cEdit;
                g.Columns[0].HeaderText = "";
                
                g.Columns.Add(cDay, "日");
                g.Columns.Add(cName, "工事／所属部署");

                DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                chk.FalseValue = global.chkBoxFalse;
                chk.TrueValue = global.chkBoxTrue;
                g.Columns.Add(chk);
                g.Columns[3].Name = cKakunin;
                g.Columns[3].HeaderText = "確認印";

                g.Columns.Add(cKinmuKbn, "出勤区分");
                g.Columns.Add(cInTM, "出社時刻");
                g.Columns.Add(cSTM, "勤務開始");
                g.Columns.Add(cETM, "勤務終了");
                g.Columns.Add(cOutTM, "退社時刻");
                g.Columns.Add(cRestTM, "休憩");
                g.Columns.Add(cHayade, "早出残業");
                g.Columns.Add(cZanTM, "普通残業");
                g.Columns.Add(cSiTM, "深夜残業");
                g.Columns.Add(cStay, "宿泊");
                g.Columns.Add(cMemo, "所定時間内で処理できない業務内容、他特記事項");
                g.Columns.Add(cAllKm, "全走行");
                g.Columns.Add(cKmTuukin, "通勤＋業務");
                g.Columns.Add(cKmShiyou, "私用");

                //g.Columns.Add(cJyosetsu, "除雪手当");
                DataGridViewCheckBoxColumn chk2 = new DataGridViewCheckBoxColumn();
                chk2.FalseValue = global.chkBoxFalse;
                chk2.TrueValue = global.chkBoxTrue;
                g.Columns.Add(chk2);
                g.Columns[18].Name = cJyosetsu;
                g.Columns[18].HeaderText = "除雪手当";

                //g.Columns.Add(cTokushu, "特殊出勤");
                DataGridViewCheckBoxColumn chk3 = new DataGridViewCheckBoxColumn();
                chk3.FalseValue = global.chkBoxFalse;
                chk3.TrueValue = global.chkBoxTrue;
                g.Columns.Add(chk3);
                g.Columns[19].Name = cTokushu;
                g.Columns[19].HeaderText = "特殊出勤";

                //g.Columns.Add(cTooshi, "通し勤務");
                DataGridViewCheckBoxColumn chk4 = new DataGridViewCheckBoxColumn();
                chk4.FalseValue = global.chkBoxFalse;
                chk4.TrueValue = global.chkBoxTrue;
                g.Columns.Add(chk4);
                g.Columns[20].Name = cTooshi;
                g.Columns[20].HeaderText = "通し勤務";

                //g.Columns.Add(cYakan, "夜間手当");
                DataGridViewCheckBoxColumn chk5 = new DataGridViewCheckBoxColumn();
                chk5.FalseValue = global.chkBoxFalse;
                chk5.TrueValue = global.chkBoxTrue;
                g.Columns.Add(chk5);
                g.Columns[21].Name = cYakan;
                g.Columns[21].HeaderText = "夜間手当";

                //g.Columns.Add(cShokumu, "職務手当");
                DataGridViewCheckBoxColumn chk6 = new DataGridViewCheckBoxColumn();
                chk6.FalseValue = global.chkBoxFalse;
                chk6.TrueValue = global.chkBoxTrue;
                g.Columns.Add(chk6);
                g.Columns[22].Name = cShokumu;
                g.Columns[22].HeaderText = "職務手当";

                g.Columns.Add(cDayMemo, "摘要");
                g.Columns.Add(cAccount, "登録者");
                g.Columns.Add(cDate, "登録年月日");
                g.Columns.Add(cEAccount, "更新者");
                g.Columns.Add(cEDate, "更新年月日");
                g.Columns.Add(cCode, "");
                g.Columns[cCode].Visible = false;

                g.Columns[cEdit].Width = 50;
                g.Columns[cDay].Width = 66;
                g.Columns[cDayMemo].Width = 100;
                g.Columns[cName].Width = 300;
                g.Columns[cKinmuKbn].Width = 80;
                g.Columns[cInTM].Width = 50;
                g.Columns[cSTM].Width = 50;
                g.Columns[cETM].Width = 50;
                g.Columns[cOutTM].Width = 50;
                g.Columns[cRestTM].Width = 50;
                g.Columns[cHayade].Width = 50;
                g.Columns[cZanTM].Width = 50;
                g.Columns[cSiTM].Width = 50;
                g.Columns[cStay].Width = 50;
                g.Columns[cMemo].Width = 160;
                g.Columns[cAllKm].Width = 60;
                g.Columns[cKmTuukin].Width = 60;
                g.Columns[cKmShiyou].Width = 60;
                g.Columns[cJyosetsu].Width = 50;
                g.Columns[cTokushu].Width = 50;
                g.Columns[cTooshi].Width = 50;
                g.Columns[cYakan].Width = 50;
                g.Columns[cShokumu].Width = 50;
                g.Columns[cKakunin].Width = 50;
                g.Columns[cAccount].Width = 80;
                g.Columns[cDate].Width = 166;
                g.Columns[cEAccount].Width = 80;
                g.Columns[cEDate].Width = 166;

                g.Columns[cDay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cInTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cSTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cETM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cOutTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cRestTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cHayade].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cZanTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cSiTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cStay].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cAllKm].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cKmTuukin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cKmShiyou].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cJyosetsu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cTokushu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cTooshi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cYakan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cShokumu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cKakunin].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cAccount].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cEAccount].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                //tempDGV.Columns[C_Memo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                g.Columns[cName].Frozen = true;
                g.Columns[cDayMemo].DefaultCellStyle.ForeColor = Color.Red;

                // 行ヘッダを表示しない
                g.RowHeadersVisible = false;

                // 選択モード
                g.SelectionMode = DataGridViewSelectionMode.CellSelect;
                g.MultiSelect = false;

                // 編集不可とする
                g.ReadOnly = false;

                // 自分以外の編集権限があるとき
                if (lgType == global.flgOn)
                {
                    foreach (DataGridViewColumn c in g.Columns)
                    {
                        if (c.Name == cKakunin || c.Name == cJyosetsu || c.Name == cTokushu ||
                            c.Name == cTooshi || c.Name == cYakan || c.Name == cShokumu)
                        {
                            c.ReadOnly = false;
                        }
                        else
                        {
                            c.ReadOnly = true;
                        }
                    }
                }
                else
                {
                    // 自分のみ編集権限があるとき
                    g.ReadOnly = true;　// 編集不可

                    // 確認欄、特殊勤務チェック欄は非表示とする
                    foreach (DataGridViewColumn c in g.Columns)
                    {
                        if (c.Name == cKakunin || c.Name == cJyosetsu || c.Name == cTokushu ||
                            c.Name == cTooshi || c.Name == cYakan || c.Name == cShokumu)
                        {
                            c.Visible = false;
                        }
                        else
                        {
                            c.ReadOnly = true;
                        }
                    }
                }

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
                //g.CellBorderStyle = DataGridViewCellBorderStyle.None;
                g.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///-------------------------------------------------------------------
        private void GridViewSetting2(DataGridView g, int lgType)
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
                g.ColumnHeadersHeight = 22;
                g.RowTemplate.Height = 22;

                // 全体の高さ
                g.Height = 90;

                // 奇数行の色
                g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                g.Columns.Add(cName2, "工事名");
                g.Columns.Add(cChiiki2, "勤務地市町村名");
                g.Columns.Add(cShukkin2, "出勤数");
                g.Columns.Add(cDaikyu2, "代休日数");
                g.Columns.Add(cHolWorkTime2, "休日勤務時間");
                g.Columns.Add(cHouteiTm2, "法定休日勤務時間");
                g.Columns.Add(cHayade2, "早出残業");
                g.Columns.Add(cZan2, "残業時間");
                g.Columns.Add(cShinya2, "深夜残業");
                g.Columns.Add(cStay2, "宿泊");
                g.Columns.Add(cKmTuukin2, "通勤＋業務");
                g.Columns.Add(cKmShiyou2, "私用");
                g.Columns.Add(cJyosetsu2, "除雪手当");
                g.Columns.Add(cTokushu2, "特殊出勤");
                g.Columns.Add(cTooshi2, "通し勤務");
                g.Columns.Add(cYakan2, "夜間手当");
                g.Columns.Add(cShokumu2, "職務手当");
                g.Columns.Add(cCode2, "");
                g.Columns[cCode2].Visible = false;

                g.Columns[cName2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                g.Columns[cChiiki2].Width = 100;
                g.Columns[cShukkin2].Width = 50;
                g.Columns[cDaikyu2].Width = 50;
                g.Columns[cHolWorkTime2].Width = 50;
                g.Columns[cHouteiTm2].Width = 50;
                g.Columns[cHayade2].Width = 50;
                g.Columns[cZan2].Width = 50;
                g.Columns[cShinya2].Width = 50;
                g.Columns[cStay2].Width = 50;
                g.Columns[cKmTuukin2].Width = 60;
                g.Columns[cKmShiyou2].Width = 60;
                g.Columns[cJyosetsu2].Width = 50;
                g.Columns[cTokushu2].Width = 50;
                g.Columns[cTooshi2].Width = 50;
                g.Columns[cYakan2].Width = 50;
                g.Columns[cShokumu2].Width = 50;

                g.Columns[cShukkin2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cDaikyu2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cHolWorkTime2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cHouteiTm2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cHayade2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cZan2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cShinya2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cStay2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cKmTuukin2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cKmShiyou2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cJyosetsu2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cTokushu2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cTooshi2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cYakan2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cShokumu2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;

                //tempDGV.Columns[C_Memo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                //g.Columns[cName2].Frozen = true;

                // 行ヘッダを表示しない
                g.RowHeadersVisible = false;

                // 選択モード
                g.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                g.MultiSelect = false;

                // 編集不可とする
                g.ReadOnly = true;

                // 自分のみ編集権限のとき
                if (lgType != global.flgOn)
                {
                    g.ReadOnly = true;　// 編集不可

                    // 特殊勤務チェック欄は非表示とする
                    foreach (DataGridViewColumn c in g.Columns)
                    {
                        if (c.Name == cJyosetsu2 || c.Name == cTokushu2 ||
                            c.Name == cTooshi2 || c.Name == cYakan2 || c.Name == cShokumu2)
                        {
                            c.Visible = false;
                        }
                    }
                }

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
        ///     データグリッドビューにデータを表示します </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト名</param>
        /// <param name="syy">
        ///     対象年</param>
        /// <param name="smm">
        ///     対象月</param>
        /// <param name="sNum">
        ///     個人コード</param>
        ///---------------------------------------------------------------------
        private void GridViewShow(DataGridView g, int syy, int smm, int sNum)
        {
            bool dataStatus = false;
            g.Rows.Clear();

            // 当月日数分の行を確保
            DateTime dt = DateTime.Parse(syy.ToString() + "/" + smm.ToString() + "/01").AddMonths(1).AddDays(-1);
            g.Rows.Add(dt.Day);
            
            global gl = new global();

            try
            {
                for (int iX = 0; iX < dt.Day; iX++)
                {
                    DateTime sDt = DateTime.Parse(syy.ToString() + "/" + smm.ToString() + "/" + (iX + 1).ToString());
                    
                    g[cDay, iX].Value = sDt.Day.ToString() + "(" + sDt.ToString("ddd") + ")";

                    // 休日か？
                    if (dts.M_休日.Any(a => a.日付 == sDt))
                    {
                        g[cDay, iX].Style.ForeColor = Color.Red;
                        var ss = dts.M_休日.Single(a => a.日付 == sDt);
                        g[cDayMemo, iX].Value = ss.摘要;

                        g.Rows[iX].DefaultCellStyle.BackColor = Color.LightPink;
                    }
                    else
                    {
                        g[cDay, iX].Style.ForeColor = Color.Black;
                        g[cDayMemo, iX].Value = string.Empty;
                    }

                    // 勤怠データを表示
                    if (dts.T_勤怠.Any(a => a.日付 == sDt && a.社員ID == sNum))
                    {
                        dataStatus = true;

                        var t = dts.T_勤怠.Single(a => a.日付 == sDt && a.社員ID == sNum);

                        g[cCode, iX].Value = t.ID.ToString();

                        g[cEdit, iX].Value = "編集";

                        if (t.M_工事Row != null)
                        {
                            g[cName, iX].Value = t.M_工事Row.名称;
                        }
                        else
                        {
                            g[cName, iX].Value = string.Empty;
                        }

                        if (t.確認印 != global.flgOff)
                        {
                            g[cKakunin, iX].Value = global.chkBoxTrue;
                        }
                        else
                        {
                            g[cKakunin, iX].Value = global.chkBoxFalse;
                        }

                        if (t.出勤印 == global.flgOn)
                        {
                            g[cKinmuKbn, iX].Value = "出勤";
                        }
                        else if (t.休日出勤 == global.flgOn)
                        {
                            g[cKinmuKbn, iX].Value = "休日出勤";
                        }
                        else if (t.代休 == global.flgOn)
                        {
                            g[cKinmuKbn, iX].Value = "代休";
                        }
                        else if (t.休日出勤 == global.flgOn)
                        {
                            g[cKinmuKbn, iX].Value = "休日出勤";
                        }
                        else if (!t.Is休日Null() && t.休日 == global.flgOn)
                        {
                            g[cKinmuKbn, iX].Value = "休日";
                        }
                        else if (!t.Is欠勤Null() && t.欠勤 == global.flgOn)
                        {
                            g[cKinmuKbn, iX].Value = "勤務日休み";
                        }

                        if (t.出勤印 == global.flgOn || t.休日出勤 == global.flgOn)
                        {
                            g[cInTM, iX].Value = t.出社時刻時 + ":" + t.出社時刻分.PadLeft(2, '0');
                            g[cSTM, iX].Value = t.開始時刻時 + ":" + t.開始時刻分.PadLeft(2, '0');
                            g[cETM, iX].Value = t.終了時刻時 + ":" + t.終了時刻分.PadLeft(2, '0');
                            g[cOutTM, iX].Value = t.退出時刻時 + ":" + t.退出時刻分.PadLeft(2, '0');
                            g[cOutTM, iX].Value = t.退出時刻時 + ":" + t.退出時刻分.PadLeft(2, '0');

                            if (t.休憩 > 0)
                            {
                                g[cRestTM, iX].Value = Utility.intToHhMM(t.休憩);
                            }
                            else
                            {
                                g[cRestTM, iX].Value = string.Empty;
                            }

                            // 2018/07/12
                            if (t.Is早出残業Null() || t.早出残業 == 0)
                            {
                                g[cHayade, iX].Value = string.Empty;
                            }
                            else
                            {
                                g[cHayade, iX].Value = Utility.intToHhMM(t.早出残業);
                            }

                            if (t.普通残業 > 0)
                            {
                                g[cZanTM, iX].Value = Utility.intToHhMM(t.普通残業);
                            }
                            else
                            {
                                g[cZanTM, iX].Value = string.Empty;
                            }

                            if (t.深夜残業 > 0)
                            {
                                g[cSiTM, iX].Value = Utility.intToHhMM(t.深夜残業);
                            }
                            else
                            {
                                g[cSiTM, iX].Value = string.Empty;
                            }
                        }
                        else
                        {
                            // 出勤日以外は特殊勤務は入力不可とする
                            g[cJyosetsu, iX].ReadOnly = true;
                            g[cTokushu, iX].ReadOnly = true;
                            g[cTooshi, iX].ReadOnly = true;
                            g[cYakan, iX].ReadOnly = true;
                            g[cShokumu, iX].ReadOnly = true;
                        }

                        if (t.宿泊 == global.flgOn)
                        {
                            g[cStay, iX].Value = "○";
                        }
                        else
                        {
                            g[cStay, iX].Value = "";
                        }

                        if (!t.Is備考Null())
                        {
                            g[cMemo, iX].Value = t.備考;
                        }
                        else
                        {
                            g[cMemo, iX].Value = string.Empty;
                        }

                        //g[cAllKm, iX].Value = "0";
                        g[cAllKm, iX].Value = getKmAll(sNum, sDt);
                        g[cKmTuukin, iX].Value = t.通勤業務走行.ToString();
                        g[cKmShiyou, iX].Value = t.私用走行.ToString();

                        if (t.除雪当番 == global.flgOn)
                        {
                            g[cJyosetsu, iX].Value = global.chkBoxTrue;
                        }
                        else
                        {
                            g[cJyosetsu, iX].Value = global.chkBoxFalse;
                        }

                        if (t.特殊出勤 == global.flgOn)
                        {
                            g[cTokushu, iX].Value = global.chkBoxTrue;
                        }
                        else
                        {
                            g[cTokushu, iX].Value = global.chkBoxFalse;
                        }

                        if (t.通し勤務 == global.flgOn)
                        {
                            g[cTooshi, iX].Value = global.chkBoxTrue;
                        }
                        else
                        {
                            g[cTooshi, iX].Value = global.chkBoxFalse;
                        }

                        if (t.夜間手当 == global.flgOn)
                        {
                            g[cYakan, iX].Value = global.chkBoxTrue;
                        }
                        else
                        {
                            g[cYakan, iX].Value = global.chkBoxFalse;
                        }

                        if (t.職務手当 == global.flgOn)
                        {
                            g[cShokumu, iX].Value = global.chkBoxTrue;
                        }
                        else
                        {
                            g[cShokumu, iX].Value = global.chkBoxFalse;
                        }

                        if (t.M_社員RowByM_社員_T_勤怠 == null)
                        {
                            g[cAccount, iX].Value = string.Empty;
                        }
                        else
                        {
                            g[cAccount, iX].Value = t.M_社員RowByM_社員_T_勤怠.氏名;
                        }

                        g[cDate, iX].Value = t.登録年月日;

                        if (t.M_社員RowByM_社員_T_勤怠2 == null)
                        {
                            g[cEAccount, iX].Value = string.Empty;
                        }
                        else
                        {
                            g[cEAccount, iX].Value = t.M_社員RowByM_社員_T_勤怠2.氏名;
                        }

                        g[cEDate, iX].Value = t.更新年月日;
                        g[cCode, iX].Value = t.ID.ToString();
                    }
                    else
                    {
                        // 退職日以降、もしくは翌日以降
                        if (sDt > taDt || sDt > DateTime.Today)
                        {
                            g[cEdit, iX].Value = "---";
                        }
                        else
                        {
                            g[cEdit, iX].Value = "登録";
                        }

                        // 入力不可
                        g[cCode, iX].Value = global.flgOff;
                        g[cKakunin, iX].ReadOnly = true;
                        g[cJyosetsu, iX].ReadOnly = true;
                        g[cTokushu, iX].ReadOnly = true;
                        g[cTooshi, iX].ReadOnly = true;
                        g[cYakan, iX].ReadOnly = true;
                        g[cShokumu, iX].ReadOnly = true;
                    }
                }

                if (g.Rows.Count > 0)
                {
                    g.CurrentCell = null;
                    changeStatus = true;
                }

                // 当月勤怠データの有無
                if (dataStatus)
                {
                    // 当月勤怠データあり
                    linkLabel2.Enabled = true;
                    linkLabel3.Enabled = true;
                }
                else
                {
                    // 当月勤怠データなし
                    linkLabel2.Enabled = false;
                    linkLabel3.Enabled = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 出勤簿表示
            //showKintaiData();
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     勤怠データ表示 </summary>
        ///-------------------------------------------------------------------------
        private void showKintaiData(double fZan)
        {
            changeStatus = false;

            //int sYY = Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei;
            int sYY = Utility.StrtoInt(txtYear.Text);       // 和暦から西暦へ 2018/07/13
            int sMM = Utility.StrtoInt(txtMonth.Text);

            adp.FillByYYMM(dts.T_勤怠, sYY, sMM);

            DateTime dt;

            if (DateTime.TryParse(sYY.ToString() + "/" + sMM.ToString() + "/01", out dt))
            {
                // データグリッドビューにデータを表示します
                GridViewShow(dg, sYY, sMM, Utility.StrtoInt(txtNum.Text));
                showSumData(dg2, sYY, sMM, Utility.StrtoInt(txtNum.Text), fZan);
            }
            else
            {
                MessageBox.Show("対象年月が正しくありません", "対象年月エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtMonth.Focus();
                return;
            }
        }

        ///----------------------------------------------------------------------------------
        /// <summary>
        ///     工事部署別集計欄表示 </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="sYY">
        ///     対象年月</param>
        /// <param name="sMM">
        ///     対象月</param>
        /// <param name="sNum">
        ///     個人コード</param>
        /// <param name="fZan">
        ///     固定残業時間</param>
        ///----------------------------------------------------------------------------------
        private void showSumData(DataGridView g, int sYY, int sMM, int sNum, double fZan)
        {
            var ss = dts.T_勤怠
                .Where(a => a.社員ID == sNum && a.日付.Year == sYY && a.日付.Month == sMM)
                .GroupBy(a => a.工事ID)
                .Select(gr => new
                {
                    pID = gr.Key,
                    pHayade = gr.Sum(a => a.早出残業),  // 2018/07/12
                    pZan = gr.Sum(a => a.普通残業),
                    pSinya = gr.Sum(a => a.深夜残業),
                    pkm = gr.Sum(a => a.通勤業務走行),
                    pkmShiyou = gr.Sum(a => a.私用走行)
                });

            int iX = 0;

            double tlZan = 0;   // 固定残業時間と比較する累積残業時間 2018/09/24
            fZan *= 60;         // 固定残業時間を分単位に変換 2018/09/24

            g.Rows.Clear();

            foreach (var t in ss)
            {
                g.Rows.Add();
                g[cCode2, iX].Value = t.pID;

                var m = dts.M_工事.Single(a => a.ID == t.pID);
                g[cName2, iX].Value = m.名称;
                g[cChiiki2, iX].Value = m.勤務地名;

                // 出勤日数
                int s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == sYY && 
                            a.日付.Month == sMM && a.社員ID == sNum && 
                            (a.出勤印 == global.flgOn || a.休日出勤 == global.flgOn));

                g[cShukkin2, iX].Value = s.ToString();

                // 代休日数
                s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == sYY &&
                            a.日付.Month == sMM && a.社員ID == sNum &&
                            a.代休 == global.flgOn);

                g[cDaikyu2, iX].Value = s.ToString();

                // 工事部署ごとの休日勤務時間・法定休日勤務時間を求める
                int hol = 0;
                int hotei = 0;
                Utility.getHolTime(dts.T_勤怠, out hol, out hotei, t.pID, sYY, sMM, sNum);

                // 工事部署ごとの休日代休時間・法定休日時間取得を求める
                int holD = 0;
                int hoteiD = 0;
                Utility.getdaikyuTime(dts, out holD, out hoteiD, sYY, sMM, sNum, t.pID);

                // 代休取得した時間を差し引く
                hol -= holD;        
                hotei -= hoteiD;

                /* ------------------------------------------------------------------
                 * 以下、固定残業時間を超過した時間を残業時間を表示する  
                 * 2018/09/17
                 * ------------------------------------------------------------------
                 */
                fix2Zan(g, iX, cHolWorkTime2, hol, ref tlZan, fZan);    // 休日勤務時間
                fix2Zan(g, iX, cHouteiTm2, hotei, ref tlZan, fZan);     // 法定休日勤務時間
                fix2Zan(g, iX, cHayade2, t.pHayade, ref tlZan, fZan);   // 早出
                fix2Zan(g, iX, cZan2, t.pZan, ref tlZan, fZan);         // 普通残業
                fix2Zan(g, iX, cShinya2, t.pSinya, ref tlZan, fZan);    // 深夜残業

                //--------------------------------------------------------------------

                // 2018/09/24 コメント化
                //g[cHolWorkTime2, iX].Value = Utility.intToHhMM(hol);
                //g[cHouteiTm2, iX].Value = Utility.intToHhMM(hotei);

                // 宿泊日数
                s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == sYY &&
                            a.日付.Month == sMM && a.社員ID == sNum &&
                            a.宿泊 == global.flgOn);

                g[cStay2, iX].Value = s.ToString();

                // 2018/09/24 コメント化
                //g[cHayade2, iX].Value = Utility.intToHhMM(t.pHayade);   // 2018/07/12
                //g[cZan2, iX].Value = Utility.intToHhMM(t.pZan);
                //g[cShinya2, iX].Value = Utility.intToHhMM(t.pSinya);

                g[cKmTuukin2, iX].Value = t.pkm.ToString();
                g[cKmShiyou2, iX].Value = t.pkmShiyou.ToString();

                // 除雪手当
                s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == sYY &&
                            a.日付.Month == sMM && a.社員ID == sNum &&
                            a.除雪当番 == global.flgOn);

                g[cJyosetsu2, iX].Value = s.ToString();

                // 特殊勤務
                s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == sYY &&
                            a.日付.Month == sMM && a.社員ID == sNum &&
                            a.特殊出勤 == global.flgOn);

                g[cTokushu2, iX].Value = s.ToString();

                // 通し勤務
                s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == sYY &&
                            a.日付.Month == sMM && a.社員ID == sNum &&
                            a.通し勤務 == global.flgOn);

                g[cTooshi2, iX].Value = s.ToString();

                // 夜間手当
                s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == sYY &&
                            a.日付.Month == sMM && a.社員ID == sNum &&
                            a.夜間手当 == global.flgOn);

                g[cYakan2, iX].Value = s.ToString();

                // 職務手当
                s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == sYY &&
                            a.日付.Month == sMM && a.社員ID == sNum &&
                            a.職務手当 == global.flgOn);

                g[cShokumu2, iX].Value = s.ToString();

                iX++;
            }

            if (iX > 0)
            {
                g.CurrentCell = null;
            }
        }


        ///------------------------------------------------------------------
        /// <summary>
        ///     固定残業時間を差し引いた時間外勤務を表示 </summary>
        /// <param name="g">
        ///     グリッドビューオブジェクト</param>
        /// <param name="iX">
        ///     rowインデックス</param>
        /// <param name="gCol">
        ///     グリッドビューカラム名</param>
        /// <param name="sZan">
        ///     該当項目残業時間</param>
        /// <param name="tlZan">
        ///     累積残業時間</param>
        /// <param name="fZan">
        ///     該当社員の固定残業時間</param>
        ///------------------------------------------------------------------
        private void fix2Zan(DataGridView g, int iX, string gCol, int sZan, ref double tlZan, double fZan)
        {
            // 該当項目残業時間がゼロなら表示して戻る
            if (sZan == global.flgOff)
            {
                g[gCol, iX].Value = Utility.intToHhMM(global.flgOff);
                return;
            }

            if (tlZan > fZan)
            {
                // 既に累積残業時間が固定残業時間を超過している場合、残業時間を表示して戻る
                g[gCol, iX].Value = Utility.intToHhMM(sZan);
            }
            else
            {
                // 累積残業時間加算
                tlZan += sZan;

                // 固定残業時間と比較する
                if (tlZan > fZan)
                {
                    // 超過した場合、差し引いた時間を表示
                    g[gCol, iX].Value = Utility.intToHhMM((int)(tlZan - fZan));
                }
                else
                {
                    // 超過していない場合、ゼロ表示
                    g[gCol, iX].Value = Utility.intToHhMM(global.flgOff);
                }
            }
        }


        ///-------------------------------------------------------------------------
        /// <summary>
        ///     走行距離取得 </summary>
        /// <param name="sNum">
        ///     個人コード</param>
        /// <param name="dt">
        ///     日付</param>
        /// <returns>
        ///     当日までの走行距離</returns>
        ///-------------------------------------------------------------------------
        private int getKmAll(int sNum, DateTime dt)
        {
            int sKm = 0;
            DateTime kDt = DateTime.Parse("1900/01/01");

            if (dts.M_社員.Any(a => a.ID == sNum))
            {
                var s = dts.M_社員.Single(a => a.ID == sNum);
                if (!s.Is走行起点Null())
                {
                    sKm = s.走行起点;

                    if (s.走行起点日付 != null && s.走行起点日付 != string.Empty)
                    {
                        kDt = DateTime.Parse(s.走行起点日付);
                    }
                }
                else
                {
                    sKm = 0;
                }
            }
            
            if (dt < kDt)
            {
                // 走行起点日付以前のとき
                var sss = dts.T_勤怠.Where(a => a.社員ID == sNum && a.日付 <= dt)
                    .GroupBy(a => a.社員ID)
                    .Select(g => new
                    {
                        pid = g.Key,
                        pKm1 = g.Sum(a => a.通勤業務走行),
                        pKm2 = g.Sum(a => a.私用走行)
                    });

                foreach (var t in sss)
                {
                    if (t.pid == sNum)
                    {
                        sKm = t.pKm1 + t.pKm2;
                    }
                }
            }
            else
            {
                // 走行起点日の翌日以降のとき走行起点Kmに加算する
                var ss = dts.T_勤怠.Where(a => a.社員ID == sNum && a.日付 <= dt && a.日付 > kDt)
                    .GroupBy(a => a.社員ID)
                    .Select(g => new
                    {
                        pid = g.Key,
                        pKm1 = g.Sum(a => a.通勤業務走行),
                        pKm2 = g.Sum(a => a.私用走行)
                    });

                foreach (var t in ss)
                {
                    if (t.pid == sNum)
                    {
                        sKm += t.pKm1;
                        sKm += t.pKm2;
                    }
                }
            }

            return sKm;
        }

        private void txtNum_TextChanged(object sender, EventArgs e)
        {

        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        ///-------------------------------------------------------------
        private void dispInitial()
        {
            //txtYear.Text = (DateTime.Today.Year - Properties.Settings.Default.rekiHosei).ToString();
            txtYear.Text = DateTime.Today.Year.ToString();      // 和暦から西暦へ 2018/07/13
            txtMonth.Text = DateTime.Today.Month.ToString();
            txtNum.Text = global.loginUserID.ToString();

            if (global.loginType == global.flgOff)
            {
                txtNum.ReadOnly = true;
                txtNum.BackColor = Color.White;
                linkLabel1.Enabled = true;
            }
            else
            {
                txtNum.ReadOnly = false;
                linkLabel1.Enabled = false;
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

        private void dg_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //int sYY = Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei;
            int sYY = Utility.StrtoInt(txtYear.Text);       // 和暦から西暦へ 2018/07/13
            int sMM = Utility.StrtoInt(txtMonth.Text);

            DateTime dt;

            if (e.ColumnIndex == 0)
            {
                if (DateTime.TryParse(sYY.ToString() + "/" + sMM.ToString() + "/" + (e.RowIndex + 1).ToString(), out dt))
                {
                    // 退職年月日以降
                    if (dt > taDt)
                    {
                        MessageBox.Show("退職年月日以降の勤怠は入力できません", "入力不可", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                    if (dt > DateTime.Today)
                    {
                        MessageBox.Show("翌日以降の勤怠は入力できません", "入力不可", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    else
                    {
                        // 勤怠データID取得
                        int pID = Utility.StrtoInt(dg[cCode, e.RowIndex].Value.ToString());

                        // 勤怠データ登録画面表示  
                        this.Hide();
                        frmKintai frm = new frmKintai(Utility.StrtoInt(txtNum.Text), dt, pID);
                        frm.ShowDialog();
                        this.Show();

                        // 出勤簿データ再表示
                        //adp.Fill(dts.T_勤怠);
                        showKintaiData(fixZan);
                    }
                }
            }
        }

        private void dg_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //int sYY = Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei;
            int sYY = Utility.StrtoInt(txtYear.Text);   // 和暦から西暦へ 2018/07/13
            int sMM = Utility.StrtoInt(txtMonth.Text);

            DateTime dt;

            if (e.ColumnIndex == 3)
            {
                if (DateTime.TryParse(sYY.ToString() + "/" + sMM.ToString() + "/" + (e.RowIndex + 1).ToString(), out dt))
                {
                    if (dt > DateTime.Today)
                    {
                        if (e.FormattedValue.ToString() == "True")
                        {
                            MessageBox.Show("翌日以降の勤怠は入力できません。チェックを外してください", "入力不可", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            //dg.Rows[e.RowIndex].ErrorText = "翌日以降の勤怠は入力できません。チェックを外してください";
                            dg.CancelEdit();
                            e.Cancel = true;
                            return;
                        }
                    }
                }
            }
        }

        private void dg_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if ((dg.CurrentCellAddress.X == 3 || dg.CurrentCellAddress.X == 17 ||
                 dg.CurrentCellAddress.X == 18 || dg.CurrentCellAddress.X == 19 ||
                 dg.CurrentCellAddress.X == 20 || dg.CurrentCellAddress.X == 21)
                 && dg.IsCurrentCellDirty)
            {
                // コミットする
                dg.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dg_CellValidated(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dg_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (changeStatus)
            {
                string n = dg.Columns[e.ColumnIndex].Name;

                // チェックボックス更新～データベース更新
                chkUpdate(n, e.RowIndex);

                // 集計欄再表示
                //showSumData(dg2, Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei, Utility.StrtoInt(txtMonth.Text), Utility.StrtoInt(txtNum.Text));

                /* 和暦から西暦へ 2018/07/13
                 * 固定残業時間引数を追加 2018/09/24 */
                showSumData(dg2, Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text), Utility.StrtoInt(txtNum.Text), fixZan);
            }
        }

        ///--------------------------------------------------------------------------------
        /// <summary>
        ///     チェックボックス更新～データベース更新 </summary>
        /// <param name="cName">
        ///     データグリッド列名</param>
        /// <param name="r">
        ///     カレント行インデックス</param>
        ///--------------------------------------------------------------------------------
        private void chkUpdate(string cName, int r)
        {
            if (changeStatus)
            {
                if (cName == cKakunin || cName == cJyosetsu || cName == cTokushu ||
                    cName == cTooshi || cName == cYakan || cName == cShokumu)
                {
                    int pID = Utility.StrtoInt(dg[cCode, r].Value.ToString());
                    var s = dts.T_勤怠.Single(a => a.ID == pID);

                    // 確認印
                    if (dg[cKakunin, r].Value.ToString() == global.chkBoxTrue)
                    {
                        s.確認印 = global.flgOn;
                    }
                    else
                    {
                        s.確認印 = global.flgOff;
                    }

                    // 除雪手当
                    if (dg[cJyosetsu, r].Value.ToString() == global.chkBoxTrue)
                    {
                        s.除雪当番 = global.flgOn;
                    }
                    else
                    {
                        s.除雪当番 = global.flgOff;
                    }

                    // 特殊出勤
                    if (dg[cTokushu, r].Value.ToString() == global.chkBoxTrue)
                    {
                        s.特殊出勤 = global.flgOn;
                    }
                    else
                    {
                        s.特殊出勤 = global.flgOff;
                    }

                    // 通し勤務
                    if (dg[cTooshi, r].Value.ToString() == global.chkBoxTrue)
                    {
                        s.通し勤務 = global.flgOn;
                    }
                    else
                    {
                        s.通し勤務 = global.flgOff;
                    }

                    // 夜間手当
                    if (dg[cYakan, r].Value.ToString() == global.chkBoxTrue)
                    {
                        s.夜間手当 = global.flgOn;
                    }
                    else
                    {
                        s.夜間手当 = global.flgOff;
                    }

                    // 職務手当
                    if (dg[cShokumu, r].Value.ToString() == global.chkBoxTrue)
                    {
                        s.職務手当 = global.flgOn;
                    }
                    else
                    {
                        s.職務手当 = global.flgOff;
                    }

                    // データベース更新
                    adp.Update(dts.T_勤怠);
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 実行確認
            if (MessageBox.Show("出勤簿・車両走行報告書を印刷しますか","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 出勤簿・車両走行報告書印刷
            sReport();
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     出勤簿・車両走行報告書印刷 </summary>
        ///------------------------------------------------------------------
        private void sReport()
        {
            //// 勤怠テーブル読み込み : 2018/07/20 当月分のみ対象とする
            ////adp.Fill(dts.T_勤怠);
            //int pYY = Utility.StrtoInt(txtYear.Text);
            //int pMM = Utility.StrtoInt(txtMonth.Text);
            //adp.FillByYYMM(dts.T_勤怠, pYY, pMM);

            //エクセルファイル日付明細開始行
            const int S_GYO = 5;
            const int S_GYO2 = 38;

            int eRow = 0;

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                // 勤務報告書テンプレートシート
                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.sxlsPath,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));
                
                Excel.Worksheet oxlsPrintSheet = null;    // 印刷用ワークシート

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                string Category = string.Empty;

                try
                {
                    // カレントのシートを設定
                    oxlsPrintSheet = (Excel.Worksheet)oXlsBook.Sheets[1];

                    // 年月
                    //oxlsPrintSheet.Cells[2, 5] = Properties.Settings.Default.gengou; // 和暦から西暦へ 2018/07/13
                    oxlsPrintSheet.Cells[2, 6] = txtYear.Text;
                    oxlsPrintSheet.Cells[2, 7] = "年";
                    oxlsPrintSheet.Cells[2, 8] = txtMonth.Text;
                    oxlsPrintSheet.Cells[2, 9] = "月分";

                    // 氏名
                    oxlsPrintSheet.Cells[2, 2] = txtName.Text;

                    //グリッドを順番に読む
                    for (int i = 0; i < dg.RowCount; i++)
                    {
                        // 勤怠登録されている日付を対象とする
                        if (dg[cCode, i].Value.ToString() != global.FLGOFF)
                        {   
                            // 印字行
                            eRow = S_GYO + i;

                            // 当日勤怠データを取得
                            var k = dts.T_勤怠.Single(a => a.ID == Utility.StrtoInt(dg[cCode, i].Value.ToString()));

                            if (dg[cCode, i].Value.ToString() != global.FLGOFF)
                            {
                                oxlsPrintSheet.Cells[eRow, 1] = (i + 1).ToString();     // 日付

                                // 休日か？
                                if (k.M_休日Row != null)
                                {
                                    // 網掛け 
                                    rng[0] = (Excel.Range)oxlsPrintSheet.Cells[eRow, 1];
                                    rng[1] = (Excel.Range)oxlsPrintSheet.Cells[eRow, 1];
                                    oxlsPrintSheet.get_Range(rng[0], rng[1]).Interior.Color = Color.LightGray;
                                }
                            }

                            // 工事名
                            oxlsPrintSheet.Cells[eRow, 2] = dg[cName, i].Value.ToString();

                            if (dg[cKinmuKbn, i].Value.ToString() == "休日" ||
                                dg[cKinmuKbn, i].Value.ToString() == "代休")
                            {
                                oxlsPrintSheet.Cells[eRow, 3] = dg[cKinmuKbn, i].Value.ToString();
                            }
                            else if (dg[cKinmuKbn, i].Value.ToString() == "勤務日休み")
                            {
                                oxlsPrintSheet.Cells[eRow, 3] = "休み";
                            }

                            oxlsPrintSheet.Cells[eRow, 4] = Utility.nulltoStr(dg[cInTM, i].Value);  // 出勤時刻
                            oxlsPrintSheet.Cells[eRow, 5] = Utility.nulltoStr(dg[cSTM, i].Value);   // 勤務開始時刻
                            oxlsPrintSheet.Cells[eRow, 6] = Utility.nulltoStr(dg[cETM, i].Value);   // 勤務終了時刻
                            oxlsPrintSheet.Cells[eRow, 7] = Utility.nulltoStr(dg[cOutTM, i].Value);  // 退出時刻
                            oxlsPrintSheet.Cells[eRow, 8] = Utility.nulltoStr(dg[cRestTM, i].Value);    // 休憩
                            
                            // 早出残業＋普通残業 2018/07/12
                            int zanTotal = hayadeZanTotal(Utility.nulltoStr(dg[cHayade, i].Value), Utility.nulltoStr(dg[cZanTM, i].Value));

                            if (zanTotal > 0)
                            {
                                oxlsPrintSheet.Cells[eRow, 9] = Utility.intToHhMM(zanTotal);
                            }
                            else
                            {
                                oxlsPrintSheet.Cells[eRow, 9] = string.Empty;
                            }
                            //oxlsPrintSheet.Cells[eRow, 9] = Utility.nulltoStr(dg[cZanTM, i].Value);     // 普通残業

                            oxlsPrintSheet.Cells[eRow, 10] = Utility.nulltoStr(dg[cSiTM, i].Value);     // 深夜時間

                            // 宿泊
                            if (dg[cStay, i].Value.ToString() != string.Empty)
                            {
                                oxlsPrintSheet.Cells[eRow, 11] = "◯";
                            }
                            else
                            {
                                oxlsPrintSheet.Cells[eRow, 11] = "";
                            }

                            oxlsPrintSheet.Cells[eRow, 12] = Utility.nulltoStr(dg[cMemo, i].Value); // 所定時間内で処理できない業務内容・他特記事項
                            oxlsPrintSheet.Cells[eRow, 14] = Utility.nulltoStr(dg[cAllKm, i].Value);    // 全走行
                            oxlsPrintSheet.Cells[eRow, 15] = Utility.nulltoStr(dg[cKmTuukin, i].Value); // 通勤＋業務
                            oxlsPrintSheet.Cells[eRow, 16] = Utility.nulltoStr(dg[cKmShiyou, i].Value); // 私用

                            // 確認印
                            if (dg[cKakunin, i].Value.ToString() == global.chkBoxTrue)
                            {
                                //oxlsPrintSheet.Cells[eRow, 17] = "✔";

                                // 確認欄に社員名を表示 : 2016/04/08 
                                if (dts.M_社員.Any(a => a.ID == k.確認印))
                                {
                                    var s = dts.M_社員.Single(a => a.ID == k.確認印);
                                    oxlsPrintSheet.Cells[eRow, 17] = s.氏名;
                                }
                                else
                                {
                                    oxlsPrintSheet.Cells[eRow, 17] = "";
                                }

                            }
                            else
                            {
                                oxlsPrintSheet.Cells[eRow, 17] = "";
                            }
                        }
                    }

                    // 集計欄
                    for (int i = 0; i < dg2.RowCount; i++)
                    {
                        // 印字は５行まで
                        if (i > 4)
                        {
                            break;
                        }

                        if (Utility.nulltoStr(dg2[cName2, i].Value) != string.Empty)
                        {
                            oxlsPrintSheet.Cells[S_GYO2 + i, 1] = Utility.nulltoStr(dg2[cName2, i].Value);  // 工事名
                            oxlsPrintSheet.Cells[S_GYO2 + i, 4] = Utility.nulltoStr(dg2[cChiiki2, i].Value);    // 地域名
                            oxlsPrintSheet.Cells[S_GYO2 + i, 5] = Utility.nulltoStr(dg2[cShukkin2, i].Value);   // 出勤数
                            oxlsPrintSheet.Cells[S_GYO2 + i, 6] = Utility.nulltoStr(dg2[cDaikyu2, i].Value);   // 代休日数
                            oxlsPrintSheet.Cells[S_GYO2 + i, 7] = Utility.nulltoStr(dg2[cHolWorkTime2, i].Value);   // 休日勤務時間
                            oxlsPrintSheet.Cells[S_GYO2 + i, 8] = Utility.nulltoStr(dg2[cHouteiTm2, i].Value);   // 法定休日勤務時間

                            // 早出残業＋普通残業 2018/07/12
                            int zanTotal = hayadeZanTotal(Utility.nulltoStr(dg2[cHayade2, i].Value), Utility.nulltoStr(dg2[cZan2, i].Value));

                            if (zanTotal > 0)
                            {
                                oxlsPrintSheet.Cells[S_GYO2 + i, 9] = Utility.intToHhMM(zanTotal);
                            }
                            else
                            {
                                oxlsPrintSheet.Cells[S_GYO2 + i, 9] = string.Empty;
                            }
                            //oxlsPrintSheet.Cells[S_GYO2 + i, 9] = Utility.nulltoStr(dg2[cZan2, i].Value);   // 残業時間

                            oxlsPrintSheet.Cells[S_GYO2 + i, 10] = Utility.nulltoStr(dg2[cShinya2, i].Value);   // 深夜時間
                            oxlsPrintSheet.Cells[S_GYO2 + i, 11] = Utility.nulltoStr(dg2[cStay2, i].Value);   // 宿直
                        }
                    }
                    
                    // 今月末走行距離
                    //int sYY = Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei;
                    int sYY = Utility.StrtoInt(txtYear.Text);   // 和暦から西暦へ 2018/07/13
                    DateTime dt = DateTime.Parse(sYY.ToString() + "/" + txtMonth.Text + "/01");
                    dt = dt.AddMonths(1).AddDays(-1);
                    oxlsPrintSheet.Cells[39, 15] = getKmAll(Utility.StrtoInt(txtNum.Text), dt);

                    // 前月末走行距離
                    dt = dt.AddMonths(-1);
                    oxlsPrintSheet.Cells[40, 15] = getKmAll(Utility.StrtoInt(txtNum.Text), dt);
                    
                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 確認のためのウィンドウを表示する
                    //oXls.Visible = true;

                    //印刷
                    //oXlsBook.PrintPreview(true);
                    oXlsBook.PrintOut();

                    //保存処理
                    oXls.DisplayAlerts = false;

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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

                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            //マウスポインタを元に戻す
            this.Cursor = Cursors.Default;
        }
        
        private int hayadeZanTotal(string gHayade, string gZan)
        {
            // 早出残業＋普通残業 2018/07/12
            string[] aa = gHayade.Split(':');
            string[] bb = gZan.Split(':');
            int hayade = 0;
            int zan = 0;

            if (aa.Length > 1)
            {
                hayade = Utility.StrtoInt(aa[0]) * 60 + Utility.StrtoInt(aa[1]);
            }

            if (bb.Length > 1)
            {
                zan = Utility.StrtoInt(bb[0]) * 60 + Utility.StrtoInt(bb[1]);
            }

            return hayade + zan;
        }

        private void txtMonth_TextChanged(object sender, EventArgs e)
        {
            linkLabel1.Enabled = false;
            linkLabel2.Enabled = false;

            // 年
            if ((Utility.StrtoInt(txtYear.Text) < 20))
            {
                dg.Rows.Clear();
                dg2.Rows.Clear();
                return;
            }

            // 月
            if ((Utility.StrtoInt(txtMonth.Text) < 1) || 
                (Utility.StrtoInt(txtMonth.Text) > 12))
            {
                dg.Rows.Clear();
                dg2.Rows.Clear();
                return;
            }

            // 個人コード
            txtName.Text = "未登録";
            txtName.ForeColor = Color.Red;

            if (Utility.StrtoInt(txtNum.Text) == global.flgOff)
            {
                dg.Rows.Clear();
                dg2.Rows.Clear();
                return;
            }
            
            if (!dts.M_社員.Any(a => a.ID == Utility.StrtoInt(txtNum.Text)))
            {
                dg.Rows.Clear();
                dg2.Rows.Clear();
                return;
            }

            // 社員情報取得
            var t = dts.M_社員.Single(a => a.ID == Utility.StrtoInt(txtNum.Text));

            txtName.Text = t.氏名;
            txtName.ForeColor = Color.Black;
            fixZan = t.固定残業時間;     // 固定残業時間取得 2018/09/24

            if (t.Is退職年月日Null() || !DateTime.TryParse(t.退職年月日, out taDt))
            {
                taDt = DateTime.Parse("2900/01/01");
                label5.Visible = false;
                txtTaiDate.Visible = false;
            }
            else
            {
                label5.Visible = true;
                txtTaiDate.Visible = true;
                txtTaiDate.Text = taDt.ToShortDateString();
            }

            //linkLabel1.Enabled = true;

            // 勤怠データ表示
            showKintaiData(fixZan);
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("表示中の勤怠データを一括して本社へ送信します。よろしいですか？","メール送信確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 当月メールをまとめて送信
            batchMail();
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     当月勤怠一括送信 </summary>
        ///---------------------------------------------------------------
        private void batchMail()
        {
            //int sYY = Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei;
            int sYY = Utility.StrtoInt(txtYear.Text);  // 和暦から西暦へ 2018/07/13
            int sMM = Utility.StrtoInt(txtMonth.Text);
            int sNum = Utility.StrtoInt(txtNum.Text);

            DateTime dt;

            if (!DateTime.TryParse(sYY.ToString() + "/" + sMM.ToString() + "/01", out dt))
            {
                return;
            }

            string[] st = null;
            int dCnt = 0;

            // 自身のメールアドレスを取得
            string mlAdd = string.Empty;
            if (dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                var ml = dts.メール設定.Single(a => a.ID == global.mailKey);
                mlAdd = ml.メールアドレス;
            }

            // 該当月勤怠データを取得
            foreach (var t in dts.T_勤怠.Where(a => a.日付.Year == dt.Year && a.日付.Month == dt.Month && a.社員ID == sNum))
            {
                Array.Resize(ref st, dCnt + 1);
                st[dCnt] = putKintaiCsv(t, mlAdd);
                dCnt++;
            }

            // 受け渡しデータパスファイル名
            string outFileName = Properties.Settings.Default.attachPath + sNum.ToString() + "_" + sYY.ToString() + sMM.ToString().PadLeft(2, '0') + ".csv";

            // 添付ファイルフォルダー内のファイルをすべて削除する
            foreach (var file in System.IO.Directory.GetFiles(Properties.Settings.Default.attachPath))
            {
                System.IO.File.Delete(file);
            }

            // CSVファイル出力
            System.IO.File.WriteAllLines(outFileName, st, System.Text.Encoding.GetEncoding("utf-8"));

            // メール件名
            string sbj = "<" + sNum.ToString() + " " + txtName.Text + "> " + sYY.ToString() + "/" + sMM.ToString().PadLeft(2, '0');

            // メール本文
            string sBody = sNum.ToString() + " " + txtName.Text + Environment.NewLine + sYY.ToString() + "年" + sMM.ToString() + "月 出勤簿データ";

            // メール送信
            Utility.sendKintaiMail(outFileName, sbj, sBody, 0);
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     送信用出勤簿CSVデータ作成 </summary>
        /// <param name="s">
        ///     genbaDataSet.T_勤怠Row</param>
        /// <returns>
        ///     パスを含むCSVファイル名</returns>
        ///---------------------------------------------------------------------
        private string putKintaiCsv(genbaDataSet.T_勤怠Row s, string mlAdd)
        {
            StringBuilder sb = new StringBuilder();

            sb.Append(s.日付.ToShortDateString()).Append(",");
            sb.Append(s.社員ID.ToString()).Append(",");
            sb.Append(s.工事ID.ToString()).Append(",");
            sb.Append(s.出勤印.ToString()).Append(",");
            sb.Append(s.出社時刻時).Append(",");
            sb.Append(s.出社時刻分).Append(",");
            sb.Append(s.開始時刻時).Append(",");
            sb.Append(s.開始時刻分).Append(",");
            sb.Append(s.終了時刻時).Append(",");
            sb.Append(s.終了時刻分).Append(",");
            sb.Append(s.退出時刻時).Append(",");
            sb.Append(s.退出時刻分).Append(",");
            sb.Append(s.休憩.ToString()).Append(",");
            sb.Append(s.普通残業.ToString()).Append(",");
            sb.Append(s.深夜残業.ToString()).Append(",");
            sb.Append(s.休日出勤.ToString()).Append(",");
            sb.Append(s.代休.ToString()).Append(",");
            sb.Append(s.休日.ToString()).Append(",");
            sb.Append(s.欠勤.ToString()).Append(",");
            sb.Append(s.宿泊.ToString()).Append(",");
            sb.Append(s.備考.Replace(",", "")).Append(",");
            sb.Append(s.除雪当番.ToString()).Append(",");
            sb.Append(s.特殊出勤.ToString()).Append(",");
            sb.Append(s.通し勤務.ToString()).Append(",");
            sb.Append(s.夜間手当.ToString()).Append(",");
            sb.Append(s.職務手当.ToString()).Append(",");
            sb.Append(global.flgOff).Append(",");
            sb.Append(s.通勤業務走行.ToString()).Append(",");
            sb.Append(s.私用走行.ToString()).Append(",");
            sb.Append(s.代休対象日.ToString()).Append(",");
            sb.Append(s.確認印.ToString()).Append(",");
            sb.Append(s.登録年月日.ToString()).Append(",");
            sb.Append(s.登録ユーザーID.ToString()).Append(",");
            sb.Append(s.更新年月日.ToString()).Append(",");
            sb.Append(s.更新ユーザーID.ToString()).Append(",");
            sb.Append(mlAdd).Append(",");
            sb.Append(s.早出残業.ToString());

            return sb.ToString();
        }

        ///--------------------------------------------------------
        /// <summary>
        ///     T_勤怠テーブル 早出残業フィールドの追加 </summary>
        ///                               2016/09/13
        ///--------------------------------------------------------
        private void dbCreateAlter()
        {
            OleDbCommand sCom = new OleDbCommand();
            mdbControl mdb = new mdbControl();
            mdb.dbConnect(sCom);

            StringBuilder sb = new StringBuilder();

            try
            {
                // T_勤怠 早出残業フィールド追加;
                sb.Clear();
                sb.Append("ALTER TABLE T_勤怠 ");
                sb.Append("ADD COLUMN 早出残業 int");

                sCom.CommandText = sb.ToString();
                sCom.ExecuteNonQuery();

                // データベース接続解除
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
            }
            finally
            {
                if (sCom.Connection.State == ConnectionState.Open)
                {
                    sCom.Connection.Close();
                }
            }
        }
    }
}
