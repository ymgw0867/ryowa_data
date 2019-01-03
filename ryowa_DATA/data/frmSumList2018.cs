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
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using LINQtoCSV;

namespace ryowa_DATA.data
{
    public partial class frmSumList2018 : Form
    {
        public frmSumList2018()
        {
            InitializeComponent();

            //adp.Fill(dts.T_勤怠);
            sAdp.Fill(dts.M_社員);
            kAdp.Fill(dts.M_工事);
            hAdp.Fill(dts.M_休日);
        }

        ryowaDataSet dts = new ryowaDataSet();
        ryowaDataSetTableAdapters.T_勤怠TableAdapter adp = new ryowaDataSetTableAdapters.T_勤怠TableAdapter();
        ryowaDataSetTableAdapters.M_社員TableAdapter sAdp = new ryowaDataSetTableAdapters.M_社員TableAdapter();
        ryowaDataSetTableAdapters.M_工事TableAdapter kAdp = new ryowaDataSetTableAdapters.M_工事TableAdapter();
        ryowaDataSetTableAdapters.M_休日TableAdapter hAdp = new ryowaDataSetTableAdapters.M_休日TableAdapter();

        // 配置日数配列クラス
        clsMounthDays md;

        // 工事毎人件費集計表データファイル
        const string K_DRAFT_DATA = @"C:\RYOWA_DATA\xlsOut\koujiShainDraft.csv";

        const string linkLabelKouji = "工事別社員別集計表示";
        const string linkLabelShain = "社員別工事別集計表示";

        private void frmSumList_Load(object sender, EventArgs e)
        {
            // フォーム最小サイズ
            Utility.WindowsMinSize(this, this.Width, this.Height);

            // データグリッド定義
            GridViewSetting(dg);

            // 画面初期化
            dispInitial();
        }

        //カラム定義
        string cNum = "col0";
        string cName = "col1";
        string cKID = "col2";
        string cKName = "col3";
        string cJinkenhi = "col4";
        string cHaichiDays = "col5";
        string cGenbaDays = "col6";
        string cKinmuchiDays = "col7";
        string cStayDays = "col8";
        string cHolTM = "col9";
        string cHouteiTM = "col10";
        string cZanTM = "col11";
        string cSiTM = "col12";
        string cJyosetsu = "col13";
        string cTokushu = "col14";
        string cTooshi = "col15";
        string cYakan = "col16";
        string cShokumu = "col17";

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
                g.Height = 630;

                // 奇数行の色
                //g.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

                g.Columns.Add(cNum, "コード");
                g.Columns.Add(cName, "氏名");
                g.Columns.Add(cKID, "コード");
                g.Columns.Add(cKName, "工事名");
                g.Columns.Add(cJinkenhi, "人件費計");
                g.Columns.Add(cHaichiDays, "配置日数");
                g.Columns.Add(cGenbaDays, "現場");
                g.Columns.Add(cKinmuchiDays, "勤務地");
                g.Columns.Add(cStayDays, "宿泊");
                g.Columns.Add(cHolTM, "休日勤務");
                g.Columns.Add(cHouteiTM, "法休勤務");
                g.Columns.Add(cZanTM, "普通残業");
                g.Columns.Add(cSiTM, "深夜残業");
                g.Columns.Add(cJyosetsu, "除雪手当");
                g.Columns.Add(cTokushu, "特殊出勤");
                g.Columns.Add(cTooshi, "通し勤務");
                g.Columns.Add(cYakan, "夜間手当");
                g.Columns.Add(cShokumu, "職務手当");
                
                g.Columns[cNum].Width = 60;
                g.Columns[cName].Width = 120;
                g.Columns[cKID].Width = 70;
                g.Columns[cKName].Width = 300;
                g.Columns[cJinkenhi].Width = 90;
                g.Columns[cHaichiDays].Width = 50;
                g.Columns[cGenbaDays].Width = 50;
                g.Columns[cKinmuchiDays].Width = 50;
                g.Columns[cStayDays].Width = 50;
                g.Columns[cHolTM].Width = 70;
                g.Columns[cHouteiTM].Width = 70;
                g.Columns[cZanTM].Width = 70;
                g.Columns[cSiTM].Width = 60;
                g.Columns[cJyosetsu].Width = 50;
                g.Columns[cTokushu].Width = 50;
                g.Columns[cTooshi].Width = 50;
                g.Columns[cYakan].Width = 50;
                g.Columns[cShokumu].Width = 50;

                g.Columns[cNum].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cKID].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomCenter;
                g.Columns[cJinkenhi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cHaichiDays].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cGenbaDays].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cKinmuchiDays].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cStayDays].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cHolTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cHouteiTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cZanTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cSiTM].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cJyosetsu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cTokushu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cTooshi].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cYakan].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;
                g.Columns[cShokumu].DefaultCellStyle.Alignment = DataGridViewContentAlignment.BottomRight;

                //tempDGV.Columns[C_Memo].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                g.Columns[cName].Frozen = true;

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
                //g.AdvancedColumnHeadersBorderStyle.All = DataGridViewAdvancedCellBorderStyle.None;
                ////g.CellBorderStyle = DataGridViewCellBorderStyle.None;
                //g.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        ///-------------------------------------------------------------
        private void dispInitial()
        {
            //txtYear.Text = (DateTime.Today.Year - Properties.Settings.Default.rekiHosei).ToString();
            txtYear.Text = DateTime.Today.Year.ToString(); // 和暦から西暦へ 2018/07/13
            txtMonth.Text = DateTime.Today.Month.ToString();
            radioButton1.Checked = true;

            //label6.Text = Properties.Settings.Default.gengou; // 和暦から西暦へ 2018/07/13

            linkLabel1.Text = linkLabelKouji;   // ラベルを工事にする 2018/09/26
            linkLabel1.Visible = false;         // 工事別社員別表示ラベル 2018/09/26
        }

        ///--------------------------------------------------------------       
        /// <summary>
        ///     人件費集計表作成（個人毎） </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="sYY">
        ///     対象年</param>
        /// <param name="sMM">
        ///     対象月</param>
        ///--------------------------------------------------------------       
        private void getSumData(DataGridView g, int sYY, int sMM)
        {
            this.Cursor = Cursors.WaitCursor;

            // 集計値クラス初期化
            clsSumData cs = new clsSumData();

            // 社員ID、工事ID毎の出勤日集計
            var s = dts.T_勤怠.Where(a => (a.日付.Year == sYY && a.日付.Month == sMM) &&
                                          (a.出勤印 == global.flgOn || a.休日出勤 == global.flgOn))
                .GroupBy(a => a.社員ID)
                .Select(n => new
                {
                    sID = n.Key,
                    pID = n.GroupBy(a => a.工事ID)
                    .Select(k => new
                    {
                        kID = k.Key,
                        cnt = k.Count(),
                        stay = k.Sum(a => a.宿泊),
                        hayadeTM = k.Sum(a => a.早出残業),
                        zanTM = k.Sum(a => a.普通残業 + a.早出残業),    // 2018/07/18
                        siTM = k.Sum(a => a.深夜残業),
                        shukkin = k.Sum(a => a.出勤印),
                        jyosetsu = k.Sum(a => a.除雪当番),
                        tokushu = k.Sum(a => a.特殊出勤),
                        tooshi = k.Sum(a => a.通し勤務),
                        yakan = k.Sum(a => a.夜間手当),
                        shokumu = k.Sum(a => a.職務手当)
                    })
                    //.OrderByDescending(a => a.cnt) 2018/09/25 コメント化
                    .OrderBy(a => a.kID)    // 工事IDで並び替え 2018/09/25
                })
                .OrderBy(a => a.sID);

            g.Rows.Clear();

            int iX = 0;
            int jkh = 0;
            int svSID = 0;
            bool colorStatus = true;
            int genbaKbn = 0;   // 2018/09/11 社員マスター現場手当有無
            double fixZan = 0;  // 2018/09/25 固定残業時間

            // 集計データ読込
            foreach (var t in s)
            {
                if (svSID != 0)
                {
                    // 社員計行追加
                    dg.Rows.Add();
                    cs.sArray[1].sID = svSID.ToString();
                    cs.sArray[1].sName = "集　　計";

                    // 端数処理：人件費単価と工事別人件費計の差異を加算 2018/09/26
                    if (cs.sArray[1].sJinkanhi != jkh)
                    {
                        int sagaku = jkh - cs.sArray[1].sJinkanhi;
                        cs.sArray[0].sJinkanhi += sagaku;
                        cs.sArray[1].sJinkanhi += sagaku;
                        cs.sArray[2].sJinkanhi += sagaku;
                        
                        // グリッドへ再表示
                        gridRowAddData(cs.sArray, dg, iX - 1, 0, colorStatus);
                    }

                    gridRowAddData(cs.sArray, dg, iX, 1, colorStatus);
                    iX++;

                    // 集計値初期化
                    cs.sumDataCrear(cs.sArray);

                    // グリッドカラーステータス反転
                    if (colorStatus)
                    {
                        colorStatus = false;
                    }
                    else
                    {
                        colorStatus = true;
                    }
                }

                cs.sArray[0].sID = t.sID.ToString();   // 個人コード

                // 氏名
                if (dts.M_社員.Any(a => a.ID == t.sID))
                {
                    var m = dts.M_社員.Single(a => a.ID == t.sID);
                    cs.sArray[0].sName = m.氏名;
                    jkh = m.人件費単価;
                    genbaKbn = m.現場手当有無;    // 2018/09/11
                    fixZan = m.固定残業時間;      // 2018/09/25
                }
                else
                {
                    cs.sArray[0].sName = "";
                    jkh = 0;
                    genbaKbn = 0;   // 2018/09/11
                    fixZan = 0;     // 2018/09/25
                }

                double tlZan = 0;   // 固定残業時間と比較する累積残業時間 2018/09/25
                fixZan *= 60;       // 固定残業時間を分単位に変換 2018/09/25

                svSID = t.sID;

                // 社員別工事別集計
                foreach (var j in t.pID)
                {
                    g.Rows.Add();

                    cs.sArray[0].pID = j.kID.ToString().PadLeft(6, '0');    // 工事ID

                    int gKbn = 0;
                    int kKbn = 0;

                    // 工事部署名
                    if (dts.M_工事.Any(a => a.ID == j.kID))
                    {
                        var m = dts.M_工事.Single(a => a.ID == j.kID);
                        cs.sArray[0].pName = m.名称;
                        gKbn = m.現場区分;
                        kKbn = m.勤務地区分;
                    }
                    else
                    {
                        cs.sArray[0].pName = "";
                    }

                    // 配置日数取得
                    cs.sArray[0].sHaichiDays = md.getHaichidays(t.sID, j.kID);

                    // 人件費
                    //cs.sArray[0].sJinkanhi = (int)(jkh / Properties.Settings.Default.tempdays * cs.sArray[0].sHaichiDays);
                    // 2018/09/26 小数点１位四捨五入
                    cs.sArray[0].sJinkanhi = (int)(jkh / Properties.Settings.Default.tempdays * cs.sArray[0].sHaichiDays + (decimal)0.5);

                    // 現場日数
                    /* 工事コードが「現場」に社員マスターの「現場手当有無」が
                     * 「有り」を条件に追加：2018/09/11 */
                    if (gKbn == global.flgOff && genbaKbn == global.flgOn)
                    {
                        cs.sArray[0].sGanbaDays = j.cnt;
                    }
                    else
                    {
                        cs.sArray[0].sGanbaDays = global.flgOff;
                    }

                    // 勤務地日数
                    if (kKbn != global.flgOff)
                    {
                        cs.sArray[0].sKinmuchiDays = j.cnt;
                    }
                    else
                    {
                        cs.sArray[0].sKinmuchiDays = global.flgOff;
                    }

                    cs.sArray[0].sStayDays = j.stay; // 宿泊

                    // 休日勤務時間・法定休日勤務時間
                    int hol = 0;
                    int hotei = 0;
                    Utility.getHolTime(dts.T_勤怠, out hol, out hotei, j.kID, sYY, sMM, t.sID);

                    // 工事部署ごとの休日代休時間・法定休日時間取得を求める
                    int holD = 0;
                    int hoteiD = 0;
                    Utility.getdaikyuTime(dts, out holD, out hoteiD, sYY, sMM, t.sID, j.kID);

                    // 代休取得した時間を差し引く
                    hol -= holD;
                    hotei -= hoteiD;



                    /* ------------------------------------------------------------------
                     * 
                     * 以下、固定残業時間を超過した時間を残業時間を表示する 2018/09/25
                     * 
                     * ------------------------------------------------------------------
                     */
                    cs.sArray[0].sHolTM = (int)fix2Zan(hol, ref tlZan, fixZan);         // 休日勤務時間
                    cs.sArray[0].sHouteiTM = (int)fix2Zan(hotei, ref tlZan, fixZan);    // 法定休日勤務時間
                    cs.sArray[0].sZanTM = (int)fix2Zan(j.zanTM, ref tlZan, fixZan);     // 普通残業
                    cs.sArray[0].sSiTM = (int)fix2Zan(j.siTM, ref tlZan, fixZan);       // 深夜残業

                    
                    // 明細行値格納
                    //cs.sArray[0].sHolTM = hol;              // 休日勤務時間 2018/09/25 コメント化
                    //cs.sArray[0].sHouteiTM = hotei;         // 法定休日勤務時間 2018/09/25 コメント化
                    //cs.sArray[0].sZanTM = j.zanTM;          // 普通残業時間 2018/09/25 コメント化
                    //cs.sArray[0].sSiTM = j.siTM;            // 深夜残業時間 2018/09/25 コメント化

                    cs.sArray[0].sJyosetsu = j.jyosetsu;    // 除雪手当
                    cs.sArray[0].sTokushu = j.tokushu;      // 特殊勤務
                    cs.sArray[0].sTooshi = j.tooshi;        // 通し勤務
                    cs.sArray[0].sYakan = j.yakan;          // 夜間手当
                    cs.sArray[0].sShokumu = j.shokumu;      // 職務手当 

                    // 集計行の値加算
                    for (int i = 1; i < cs.sArray.Length; i++)
                    {
                        cs.sArray[i].sJinkanhi += cs.sArray[0].sJinkanhi;
                        cs.sArray[i].sHaichiDays += cs.sArray[0].sHaichiDays;
                        cs.sArray[i].sGanbaDays += cs.sArray[0].sGanbaDays;
                        cs.sArray[i].sKinmuchiDays += cs.sArray[0].sKinmuchiDays;
                        cs.sArray[i].sStayDays += cs.sArray[0].sStayDays;
                        cs.sArray[i].sHolTM += cs.sArray[0].sHolTM;
                        cs.sArray[i].sHouteiTM += cs.sArray[0].sHouteiTM;
                        cs.sArray[i].sZanTM += cs.sArray[0].sZanTM;
                        cs.sArray[i].sSiTM += cs.sArray[0].sSiTM;
                        cs.sArray[i].sJyosetsu += cs.sArray[0].sJyosetsu;
                        cs.sArray[i].sTokushu += cs.sArray[0].sTokushu;
                        cs.sArray[i].sTooshi += cs.sArray[0].sTooshi;
                        cs.sArray[i].sYakan += cs.sArray[0].sYakan;
                        cs.sArray[i].sShokumu += cs.sArray[0].sShokumu;
                    }

                    // グリッドへ行追加
                    gridRowAddData(cs.sArray, dg, iX, 0, colorStatus);

                    iX++;
                }
            }

            // 社員計行追加
            dg.Rows.Add();
            cs.sArray[1].sID = svSID.ToString();
            cs.sArray[1].sName = "集　　計";

            // 端数処理：人件費単価と工事別人件費計の差異を加算 2018/09/26
            if (cs.sArray[1].sJinkanhi != jkh)
            {
                int sagaku = jkh - cs.sArray[1].sJinkanhi;
                cs.sArray[0].sJinkanhi += sagaku;
                cs.sArray[1].sJinkanhi += sagaku;
                cs.sArray[2].sJinkanhi += sagaku;

                // グリッドへ再表示
                gridRowAddData(cs.sArray, dg, iX - 1, 0, colorStatus);
            }

            gridRowAddData(cs.sArray, dg, iX, 1, colorStatus);
            iX++;

            // グリッドカラーステータス反転
            if (colorStatus)
            {
                colorStatus = false;
            }
            else
            {
                colorStatus = true;
            }

            // 総計行追加
            dg.Rows.Add();
            cs.sArray[2].sName = "総　　計";
            gridRowAddData(cs.sArray, dg, iX, 2, colorStatus);

            if (dg.Rows.Count > 0)
            {
                dg.CurrentCell = null;
            }
            
            this.Cursor = Cursors.Default;
        }
        

        ///------------------------------------------------------------------
        /// <summary>
        ///     固定残業時間を差し引いた時間外勤務を計算 </summary>
        /// <param name="sZan">
        ///     該当項目残業時間</param>
        /// <param name="tlZan">
        ///     累積残業時間</param>
        /// <param name="fZan">
        ///     該当社員の固定残業時間</param>
        ///------------------------------------------------------------------
        private double fix2Zan(int sZan, ref double tlZan, double fZan)
        {
            // 該当項目残業時間がゼロならゼロを返す
            if (sZan == global.flgOff)
            {
                return (double)global.flgOff;
            }

            if (tlZan > fZan)
            {
                // 既に累積残業時間が固定残業時間を超過している場合、残業時間を返す
                return sZan;
            }
            else
            {
                // 累積残業時間加算
                tlZan += sZan;

                // 固定残業時間と比較する
                if (tlZan > fZan)
                {
                    // 超過した場合、差し引いた時間を返す
                    return (tlZan - fZan);
                }
                else
                {
                    // 超過していない場合、ゼロを返す
                    return (double)global.flgOff;
                }
            }
        }

        
        ///--------------------------------------------------------------       
        /// <summary>
        ///     人件費集計表作成（工事毎）</summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="sYY">
        ///     対象年</param>
        /// <param name="sMM">
        ///     対象月</param>
        ///--------------------------------------------------------------       
        private void getSumDataKouji(DataGridView g, int sYY, int sMM)
        {
            this.Cursor = Cursors.WaitCursor;

            // 集計値クラスインスタンス
            clsSumData cs = new clsSumData();

            // 工事ID、社員ID毎の出勤日集計
            var s = dts.T_勤怠.Where(a => (a.日付.Year == sYY && a.日付.Month == sMM) &&
                                          (a.出勤印 == global.flgOn || a.休日出勤 == global.flgOn))
                .GroupBy(a => a.工事ID)
                .Select(n => new
                {
                    kID = n.Key,
                    pID = n.GroupBy(a => a.社員ID)
                    .Select(k => new
                    {
                        sID = k.Key,
                        cnt = k.Count(),
                        stay = k.Sum(a => a.宿泊),
                        zanTM = k.Sum(a => a.普通残業 + a.早出残業),    // 2018/07/19
                        siTM = k.Sum(a => a.深夜残業),
                        shukkin = k.Sum(a => a.出勤印),
                        jyosetsu = k.Sum(a => a.除雪当番),
                        tokushu = k.Sum(a => a.特殊出勤),
                        tooshi = k.Sum(a => a.通し勤務),
                        yakan = k.Sum(a => a.夜間手当),
                        shokumu = k.Sum(a => a.職務手当)
                    })
                    .OrderByDescending(a => a.cnt)
                })
                .OrderBy(a => a.kID).ToList();

            g.Rows.Clear();
            int iX = 0;
            int jkh = 0;
            int svSID = 0;
            bool colorStatus = true;
            int genbaKbn = 0;           // 2018/09/13 社員マスター現場手当有無

            // 集計データ読込
            foreach (var t in s)
            {
                if (svSID != 0)
                {
                    // 工事計行追加
                    dg.Rows.Add();
                    cs.sArray[1].pID = svSID.ToString();
                    cs.sArray[1].pName = "集　　計";
                    gridRowAddData(cs.sArray, dg, iX, 1, colorStatus);
                    iX++;

                    // 集計値初期化
                    cs.sumDataCrear(cs.sArray);

                    // グリッドカラーステータス反転
                    if (colorStatus)
                    {
                        colorStatus = false;
                    }
                    else
                    {
                        colorStatus = true;
                    }
                }
                
                cs.sArray[0].pID = t.kID.ToString();   // 工事ID

                int gKbn = 0;
                int kKbn = 0;

                // 工事部署名
                if (dts.M_工事.Any(a => a.ID == t.kID))
                {
                    var m = dts.M_工事.Single(a => a.ID == t.kID);
                    cs.sArray[0].pName = m.名称;
                    gKbn = m.現場区分;
                    kKbn = m.勤務地区分;
                }
                else
                {
                    cs.sArray[0].pName = "";
                }
                
                svSID = t.kID;

                // 社員別工事別集計
                foreach (var j in t.pID)
                {
                    g.Rows.Add();

                    cs.sArray[0].sID = j.sID.ToString();   // 個人コード

                    // 氏名
                    if (dts.M_社員.Any(a => a.ID == j.sID))
                    {
                        var m = dts.M_社員.Single(a => a.ID == j.sID);
                        cs.sArray[0].sName = m.氏名;
                        jkh = m.人件費単価;
                        genbaKbn = m.現場手当有無;    // 2018/09/13
                    }
                    else
                    {
                        cs.sArray[0].sName = "";
                        jkh = 0;
                        genbaKbn = 0;    // 2018/09/13
                    }

                    cs.sArray[0].sHaichiDays = md.getHaichidays(j.sID, t.kID); // 配置日数取得
                    cs.sArray[0].sJinkanhi = (int)(jkh / Properties.Settings.Default.tempdays * cs.sArray[0].sHaichiDays);   // 人件費  

                    // 現場日数
                    /* 工事コードが「現場」に社員マスターの「現場手当有無」が
                     * 「有り」を条件に追加：2018/09/13 */
                    if (gKbn == global.flgOff && genbaKbn == global.flgOn)
                    {
                        cs.sArray[0].sGanbaDays = j.cnt;
                    }
                    else
                    {
                        cs.sArray[0].sGanbaDays = global.flgOff;
                    }

                    // 勤務地日数
                    if (kKbn != global.flgOff)
                    {
                        cs.sArray[0].sKinmuchiDays = j.cnt;
                    }
                    else
                    {
                        cs.sArray[0].sKinmuchiDays = global.flgOff;
                    }

                    cs.sArray[0].sStayDays = j.stay; // 宿泊

                    // 休日勤務時間・法定休日勤務時間
                    int hol = 0;
                    int hotei = 0;
                    Utility.getHolTime(dts.T_勤怠, out hol, out hotei, t.kID, sYY, sMM, j.sID);

                    // 工事部署ごとの休日代休時間・法定休日時間取得を求める
                    int holD = 0;
                    int hoteiD = 0;
                    Utility.getdaikyuTime(dts, out holD, out hoteiD, sYY, sMM, j.sID, t.kID);

                    // 代休取得した時間を差し引く
                    hol -= holD;
                    hotei -= hoteiD;

                    // 明細行値格納
                    cs.sArray[0].sHolTM = hol; // 休日勤務時間
                    cs.sArray[0].sHouteiTM = hotei;    // 法定休日勤務時間
                    cs.sArray[0].sZanTM = j.zanTM;     // 普通残業時間
                    cs.sArray[0].sSiTM = j.siTM;       // 深夜残業時間
                    cs.sArray[0].sJyosetsu = j.jyosetsu;   // 除雪手当
                    cs.sArray[0].sTokushu = j.tokushu;     // 特殊勤務
                    cs.sArray[0].sTooshi = j.tooshi;       // 通し勤務
                    cs.sArray[0].sYakan = j.yakan;         // 夜間手当
                    cs.sArray[0].sShokumu = j.shokumu;     // 職務手当 

                    // 集計行の値加算
                    for (int i = 1; i < cs.sArray.Length; i++)
                    {
                        cs.sArray[i].sJinkanhi += cs.sArray[0].sJinkanhi;
                        cs.sArray[i].sHaichiDays += cs.sArray[0].sHaichiDays;
                        cs.sArray[i].sGanbaDays += cs.sArray[0].sGanbaDays;
                        cs.sArray[i].sKinmuchiDays += cs.sArray[0].sKinmuchiDays;
                        cs.sArray[i].sStayDays += cs.sArray[0].sStayDays;
                        cs.sArray[i].sHolTM += cs.sArray[0].sHolTM;
                        cs.sArray[i].sHouteiTM += cs.sArray[0].sHouteiTM;
                        cs.sArray[i].sZanTM += cs.sArray[0].sZanTM;
                        cs.sArray[i].sSiTM += cs.sArray[0].sSiTM;
                        cs.sArray[i].sJyosetsu += cs.sArray[0].sJyosetsu;
                        cs.sArray[i].sTokushu += cs.sArray[0].sTokushu;
                        cs.sArray[i].sTooshi += cs.sArray[0].sTooshi;
                        cs.sArray[i].sYakan += cs.sArray[0].sYakan;
                        cs.sArray[i].sShokumu += cs.sArray[0].sShokumu;
                    }

                    // グリッドへ行追加
                    gridRowAddData(cs.sArray, dg, iX, 0, colorStatus);

                    iX++;
                }
            }

            // 社員計行追加
            dg.Rows.Add();
            cs.sArray[1].pID = svSID.ToString();
            cs.sArray[1].pName = "集　　計";
            gridRowAddData(cs.sArray, dg, iX, 1, colorStatus);
            iX++;

            // グリッドカラーステータス反転
            if (colorStatus)
            {
                colorStatus = false;
            }
            else
            {
                colorStatus = true;
            }

            // 総計行追加
            dg.Rows.Add();
            cs.sArray[2].pName = "総　　計";
            gridRowAddData(cs.sArray, dg, iX, 2, colorStatus);

            if (dg.Rows.Count > 0)
            {
                dg.CurrentCell = null;
            }

            this.Cursor = Cursors.Default;
        }


        ///--------------------------------------------------------------       
        /// <summary>
        ///     人件費集計表作成（工事毎）</summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="sYY">
        ///     対象年</param>
        /// <param name="sMM">
        ///     対象月</param>
        ///--------------------------------------------------------------       
        private void getSumDataKouji_201809(DataGridView g, int sYY, int sMM)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;

                // 集計値クラスインスタンス
                clsSumData cs = new clsSumData();

                var context = new CsvContext();

                // CSVの情報を示すオブジェクトを構築
                var description = new CsvFileDescription
                {
                    SeparatorChar = ',',
                    FirstLineHasColumnNames = false,
                    EnforceCsvColumnAttribute = true,
                    TextEncoding = Encoding.GetEncoding(932)
                };

                var s = context.Read<common.clsLinqCsv>(K_DRAFT_DATA, description)
                    //.Select(n => new
                    //{
                    //    pID = n.pID,
                    //    pName = n.pName,
                    //    sID = n.sID,
                    //    sName = n.sName,
                    //    sHaichiDays = n.sHaichiDays,
                    //    sJinkanhi = n.sJinkanhi
                    //})
                    .OrderBy(a => a.pID).ThenBy(a => a.sID);


                g.Rows.Clear();
                int iX = 0;
                int svSID = 0;
                int svPID = 0;
                bool colorStatus = true;

                // 集計データ読込
                foreach (var t in s)
                {
                    if (svPID != 0 && svPID != t.pID)
                    {
                        // 工事計行追加
                        g.Rows.Add();
                        cs.sArray[1].pID = svPID.ToString();
                        cs.sArray[1].pName = "集　　計";
                        gridRowAddData(cs.sArray, dg, iX, 1, colorStatus);
                        iX++;

                        // 集計値初期化
                        cs.sumDataCrear(cs.sArray);

                        // グリッドカラーステータス反転
                        if (colorStatus)
                        {
                            colorStatus = false;
                        }
                        else
                        {
                            colorStatus = true;
                        }
                    }

                    svSID = t.sID;

                    cs.sArray[0].pID = t.pID.ToString();        // 工事ID
                    cs.sArray[0].pName = t.pName;               // 工事名称
                    cs.sArray[0].sID = t.sID.ToString();        // 個人コード
                    cs.sArray[0].sName = t.sName;               // 個人名
                    cs.sArray[0].sHaichiDays = t.sHaichiDays;   // 配置日数取得
                    cs.sArray[0].sJinkanhi = t.sJinkanhi;       // 人件費
                    cs.sArray[0].sGanbaDays = Utility.StrtoInt(t.sGanbaDays);           // 現場日数
                    cs.sArray[0].sKinmuchiDays = Utility.StrtoInt(t.sKinmuchiDays);     // 勤務地日数
                    cs.sArray[0].sStayDays = Utility.StrtoInt(t.sStayDays);     // 宿泊日数
                    cs.sArray[0].sHolTM = (int)(t.sHolTM * 60);                 // 休日勤務時間
                    cs.sArray[0].sHouteiTM = (int)(t.sHouteiTM * 60);           // 法定休日勤務時間
                    cs.sArray[0].sZanTM = (int)(t.sZanTM * 60);                 // 普通残業時間
                    cs.sArray[0].sSiTM = (int)(t.sSiTM *60);                    // 深夜残業時間
                    cs.sArray[0].sJyosetsu = Utility.StrtoInt(t.sJyosetsu);     // 除雪手当
                    cs.sArray[0].sTokushu = Utility.StrtoInt(t.sTokushu);       // 特殊勤務
                    cs.sArray[0].sTooshi = Utility.StrtoInt(t.sTooshi);         // 通し勤務
                    cs.sArray[0].sYakan = Utility.StrtoInt(t.sYakan);           // 夜間手当
                    cs.sArray[0].sShokumu = Utility.StrtoInt(t.sShokumu);       // 職務手当 


                    // 集計行の値加算
                    for (int i = 1; i < cs.sArray.Length; i++)
                    {
                        cs.sArray[i].sJinkanhi += cs.sArray[0].sJinkanhi;
                        cs.sArray[i].sHaichiDays += cs.sArray[0].sHaichiDays;
                        cs.sArray[i].sGanbaDays += cs.sArray[0].sGanbaDays;
                        cs.sArray[i].sKinmuchiDays += cs.sArray[0].sKinmuchiDays;
                        cs.sArray[i].sStayDays += cs.sArray[0].sStayDays;
                        cs.sArray[i].sHolTM += cs.sArray[0].sHolTM;
                        cs.sArray[i].sHouteiTM += cs.sArray[0].sHouteiTM;
                        cs.sArray[i].sZanTM += cs.sArray[0].sZanTM;
                        cs.sArray[i].sSiTM += cs.sArray[0].sSiTM;
                        cs.sArray[i].sJyosetsu += cs.sArray[0].sJyosetsu;
                        cs.sArray[i].sTokushu += cs.sArray[0].sTokushu;
                        cs.sArray[i].sTooshi += cs.sArray[0].sTooshi;
                        cs.sArray[i].sYakan += cs.sArray[0].sYakan;
                        cs.sArray[i].sShokumu += cs.sArray[0].sShokumu;
                    }

                    // グリッドへ行追加
                    g.Rows.Add();
                    gridRowAddData(cs.sArray, dg, iX, 0, colorStatus);

                    iX++;

                    svPID = t.pID;
                }

                // 社員計行追加
                dg.Rows.Add();
                cs.sArray[1].pID = svPID.ToString();
                cs.sArray[1].pName = "集　　計";
                gridRowAddData(cs.sArray, dg, iX, 1, colorStatus);
                iX++;

                // グリッドカラーステータス反転
                if (colorStatus)
                {
                    colorStatus = false;
                }
                else
                {
                    colorStatus = true;
                }

                // 総計行追加
                dg.Rows.Add();
                cs.sArray[2].pName = "総　　計";
                gridRowAddData(cs.sArray, dg, iX, 2, colorStatus);

                if (dg.Rows.Count > 0)
                {
                    dg.CurrentCell = null;
                }

                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        ///--------------------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューデータ行追加 </summary>
        /// <param name="sd">
        ///     データ集計配列</param>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="r">
        ///     データグリッドビューカレント行インデックス</param>
        /// <param name="i">
        ///     データ集計配列インデックス</param>
        ///--------------------------------------------------------------------------------
        private void gridRowAddData(clsSumData.sumData[] sd, DataGridView g, int r, int i, bool colorStatus)
        {
            g[cNum, r].Value = sd[i].sID.ToString();                // 個人コード
            g[cName, r].Value = sd[i].sName.ToString();             // 氏名
            g[cKID, r].Value = sd[i].pID.ToString();                // 工事コード
            g[cKName, r].Value = sd[i].pName.ToString();            // 工事名
            g[cJinkenhi, r].Value = sd[i].sJinkanhi.ToString("#,##0");  // 人件費
            g[cHaichiDays, r].Value = sd[i].sHaichiDays.ToString("#,##0.0");     // 配置日数
            g[cGenbaDays, r].Value = sd[i].sGanbaDays.ToString("#,###");       // 現場日数
            g[cKinmuchiDays, r].Value = sd[i].sKinmuchiDays.ToString("#,###"); // 勤務地
            g[cStayDays, r].Value = sd[i].sStayDays.ToString("#,###");     // 宿泊

            // 休日勤務時間
            if (sd[i].sHolTM > 0)
            {
                g[cHolTM, r].Value = Utility.intToHhMM10(sd[i].sHolTM, Properties.Settings.Default.marume); // 休日勤務時間
            }
            else
            {
                g[cHolTM, r].Value = global.FLGOFF; 
            }

            // 法定休日勤務時間      
            if (sd[i].sHouteiTM > 0)
            {
                g[cHouteiTM, r].Value = Utility.intToHhMM10(sd[i].sHouteiTM, Properties.Settings.Default.marume); 
            }
            else
            {
                g[cHouteiTM, r].Value = global.FLGOFF;
            }
            
            // 普通残業時間          
            if (sd[i].sZanTM > 0)
            {
                g[cZanTM, r].Value = Utility.intToHhMM10(sd[i].sZanTM, Properties.Settings.Default.marume); // 休日勤務時間
            }
            else
            {
                g[cZanTM, r].Value = global.FLGOFF;
            }

            // 深夜残業時間        
            if (sd[i].sSiTM > 0)
            {
                g[cSiTM, r].Value = Utility.intToHhMM10(sd[i].sSiTM, Properties.Settings.Default.marume); // 休日勤務時間
            }
            else
            {
                g[cSiTM, r].Value = global.FLGOFF;
            }

            g[cJyosetsu, r].Value = sd[i].sJyosetsu.ToString("#,###");     // 除雪手当
            g[cTokushu, r].Value = sd[i].sTokushu.ToString("#,###");       // 特殊勤務
            g[cTooshi, r].Value = sd[i].sTooshi.ToString("#,###");         // 通し勤務
            g[cYakan, r].Value = sd[i].sYakan.ToString("#,###");           // 夜間手当
            g[cShokumu, r].Value = sd[i].sShokumu.ToString("#,###");       // 職務手当

            // 集計行のとき
            if (i > 0)
            {
                g.Rows[r].DefaultCellStyle.BackColor = Color.Lavender;
                g[cName, r].Style.Alignment = DataGridViewContentAlignment.BottomCenter;
                g[cKName, r].Style.Alignment = DataGridViewContentAlignment.BottomCenter;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //getSumData(dg);
        }

        private void txtYear_TextChanged(object sender, EventArgs e)
        {
            sumShainByKouji();
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     社員別工事別社員別人件費集計表表示 : 2018/09/26 </summary>
        ///---------------------------------------------------------------
        private void sumShainByKouji()
        {
            //linkLabel1.Enabled = false;
            //linkLabel2.Enabled = false;

            // 年
            if ((Utility.StrtoInt(txtYear.Text) < 2016))
            {
                dg.Rows.Clear();
                return;
            }

            // 月
            if ((Utility.StrtoInt(txtMonth.Text) < 1) ||
                (Utility.StrtoInt(txtMonth.Text) > 12))
            {
                dg.Rows.Clear();
                return;
            }

            //int sYY = Utility.StrtoInt(txtYear.Text) + Properties.Settings.Default.rekiHosei;
            int sYY = Utility.StrtoInt(txtYear.Text); // 和暦から西暦へ 2018/07/13
            int sMM = Utility.StrtoInt(txtMonth.Text);

            // 配置日数取得
            md = new clsMounthDays(sYY, sMM);

            // 当月データ取得：2018/07/20
            adp.FillByYYMM(dts.T_勤怠, sYY, sMM);

            // 勤怠データ表示
            //if (radioButton1.Checked)
            //{
            //    getSumData(dg, sYY, sMM);
            //    grid2Csv();     // 工事毎人件費集計表データ出力
            //}
            //else if (radioButton2.Checked)
            //{
            //    //getSumDataKouji(dg, sYY, sMM);        // コメント化 2018/09/26
            //    getSumDataKouji_201809(dg, sYY, sMM);   // 2018/09/26
            //}

            // 勤怠データ表示
            getSumData(dg, sYY, sMM);
            grid2Csv();     // 工事毎人件費集計表データ出力
        }


        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 実行確認
            if (MessageBox.Show("人件費集計一覧表を印刷しますか", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            // 人件費集計一覧表印刷
            sReport2018();
        }


        ///------------------------------------------------------------------
        /// <summary>
        ///     人件費集計一覧表印刷 </summary>
        ///     
        /// 　closedXML版　2018/07/26
        ///------------------------------------------------------------------
        private void sReport2018()
        {
            //エクセルファイル日付明細開始行
            const int S_GYO = 5;
            //const int max_GYO = 22; デバッグ用
            const int max_GYO = 53;
            int item_Gyo = max_GYO - S_GYO + 1;

            int pCnt = 1;
            int eRow = 0;
            int colorStatus = 1;
            string clrCode = string.Empty;

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                Excel.Application oXls = new Excel.Application();
                Excel.Workbook oXlsBook = null;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;                
                string Category = string.Empty;

                try
                {
                    using (var tXls = new XLWorkbook(Properties.Settings.Default.sxlsJinPath, XLEventTracking.Disabled))
                    {
                        // カレントのシートを設定
                        var tmpSheet = tXls.Worksheet(1);

                        // ヘッダ
                        string tittle1 = txtYear.Text + "年 " + txtMonth.Text + "月分";
                        tmpSheet.Cell(1, 2).Value = tittle1;

                        //string tittle2 = "土木部員配置＆人件費集計一覧表";
                        string tittle2 = "部員配置＆人件費集計一覧表";   // 2018/12/15「土木」を撤廃

                        //if (radioButton1.Checked)
                        //{
                        //    tittle2 += "（個人毎）";
                        //}
                        //else
                        //{
                        //    tittle2 += "（工事毎）";
                        //}

                        // 2018/09/26
                        if (linkLabel1.Text == linkLabelKouji)
                        {
                            tittle2 += "（個人毎）";
                        }
                        else
                        {
                            tittle2 += "（工事毎）";
                        }

                        tmpSheet.Cell(2, 2).Value = tittle2;

                        //グリッドを順番に読む
                        for (int i = 0; i < dg.RowCount; i++)
                        {
                            // ページカウント
                            int p = 0;

                            if (((i + 1) % item_Gyo) > 0)
                            {
                                p = (i + 1) / item_Gyo + 1;
                            }
                            else
                            {
                                p = (i + 1) / item_Gyo;
                            }

                            // 新しいページ？
                            if (pCnt != p)
                            {
                                // ページカウント
                                pCnt = p;

                                // 3ページ以降シートを追加する 
                                if (p > 2)
                                {
                                    var copySheet = tXls.Worksheet(2);
                                    copySheet.CopyTo(tXls, "p" + p, p);
                                }

                                // カレントシート
                                tmpSheet = tXls.Worksheet(p);
                                tmpSheet.Cell("R1").Value = "P." + p.ToString("D3");

                                // カレントシートの明細欄を初期化
                                IXLRange rng = tmpSheet.Range("A5", "R53");
                                rng.Value = string.Empty;
                                rng.Style.Fill.BackgroundColor = XLColor.White;
                            }

                            // 印字行
                            eRow = S_GYO + (i - (item_Gyo * (pCnt - 1)));
                            
                            // セルに値をセット
                            tmpSheet.Cell(eRow, 1).Value = dg[cNum, i].Value.ToString();            // 個人コード
                            tmpSheet.Cell(eRow, 2).Value = dg[cName, i].Value.ToString();           // 個人名
                            tmpSheet.Cell(eRow, 3).Value = dg[cKID, i].Value.ToString();            // 工事コード
                            tmpSheet.Cell(eRow, 4).Value = dg[cKName, i].Value.ToString();          // 工事名
                            tmpSheet.Cell(eRow, 5).Value = dg[cJinkenhi, i].Value.ToString();       // 人件費
                            tmpSheet.Cell(eRow, 6).Value = dg[cHaichiDays, i].Value.ToString();     // 配置日数
                            tmpSheet.Cell(eRow, 7).Value = dg[cGenbaDays, i].Value.ToString();      // 現場
                            tmpSheet.Cell(eRow, 8).Value = dg[cKinmuchiDays, i].Value.ToString();   // 勤務地
                            tmpSheet.Cell(eRow, 9).Value = dg[cStayDays, i].Value.ToString();  // 宿泊
                            tmpSheet.Cell(eRow, 10).Value = dg[cHolTM, i].Value.ToString();    // 休日出勤時間
                            tmpSheet.Cell(eRow, 11).Value = dg[cHouteiTM, i].Value.ToString(); // 法定休日勤務時間
                            tmpSheet.Cell(eRow, 12).Value = dg[cZanTM, i].Value.ToString();    // 残業時間
                            tmpSheet.Cell(eRow, 13).Value = dg[cSiTM, i].Value.ToString();     // 深夜残業
                            tmpSheet.Cell(eRow, 14).Value = dg[cJyosetsu, i].Value.ToString(); // 除雪手当
                            tmpSheet.Cell(eRow, 15).Value = dg[cTokushu, i].Value.ToString();  // 特殊勤務
                            tmpSheet.Cell(eRow, 16).Value = dg[cTooshi, i].Value.ToString();   // 通し勤務
                            tmpSheet.Cell(eRow, 17).Value = dg[cYakan, i].Value.ToString();    // 夜間手当
                            tmpSheet.Cell(eRow, 18).Value = dg[cShokumu, i].Value.ToString();  // 職務手当

                            if (radioButton1.Checked)
                            {
                                if (clrCode != dg[cNum, i].Value.ToString())
                                {
                                    clrCode = dg[cNum, i].Value.ToString();
                                    colorStatus *= -1;
                                }
                            }
                            else
                            {
                                if (clrCode != dg[cKID, i].Value.ToString())
                                {
                                    clrCode = dg[cKID, i].Value.ToString();
                                    colorStatus *= -1;
                                }
                            }

                            // 網掛け
                            if (colorStatus == -1)
                            {
                                for (int iX = 1; iX <= 9; iX++)
                                {
                                    var cell1 = tmpSheet.Cell(eRow, iX);
                                    cell1.Style.Fill.BackgroundColor = XLColor.LightBlue;
                                }
                            }
                        }

                        // 総ページ数が1ページのとき2ページ目のシートを削除
                        if (pCnt < 2)
                        {
                            // ページ削除
                            tXls.Worksheet(2).Delete();
                        }

                        //マウスポインタを元に戻す
                        this.Cursor = Cursors.Default;

                        //保存処理
                        tXls.SaveAs(sAppPath + @"\" + global.JINKENHI_XLSX);
                    }

                    // 印刷用ブックとして開く
                    oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sAppPath + @"\" + global.JINKENHI_XLSX,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                Type.Missing, Type.Missing));


                    // 確認のためのウィンドウを表示する
                    oXls.Visible = true;
                    this.Hide();

                    //印刷
                    oXlsBook.PrintOutEx(Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //oXlsBook.PrintPreview(true);
                    //oXlsBook.PrintOut();

                    //保存処理
                    oXls.DisplayAlerts = false;

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    // 以下、エクセルファイル保存ロジック 2018/07/27 とりあえず現段階ではコメント化
                    //DialogResult ret;

                    ////ダイアログボックスの初期設定
                    //SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    //saveFileDialog1.Title = "人件費集計一覧表";
                    //saveFileDialog1.OverwritePrompt = true;
                    //saveFileDialog1.RestoreDirectory = true;
                    //DateTime dt = DateTime.Now;
                    //saveFileDialog1.FileName = txtYear.Text + "年" + txtMonth.Text.PadLeft(2, ' ') + "月 人件費集計一覧表";
                    //saveFileDialog1.Filter = "Microsoft Office Excelファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                    ////ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
                    //string fileName;
                    //ret = saveFileDialog1.ShowDialog();

                    //if (ret == System.Windows.Forms.DialogResult.OK)
                    //{
                    //    using (var uXls = new XLWorkbook(sAppPath + @"\" + global.JINKENHI_XLSX, XLEventTracking.Disabled))
                    //    {
                    //        uXls.SaveAs(saveFileDialog1.FileName);

                    //        //fileName = saveFileDialog1.FileName;
                    //        //oXlsBook.SaveAs(fileName, Type.Missing, Type.Missing,
                    //        //                Type.Missing, Type.Missing, Type.Missing,
                    //        //                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                    //        //                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //    }
                    //}

                    this.Show();

                    MessageBox.Show("処理終了しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message, "人件費集計表作成処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                }
                finally
                {
                    //Excelを終了
                    oXls.Quit();

                    // COM オブジェクトの参照カウントを解放する
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);

                    oXls = null;
                    oXlsBook = null;

                    GC.Collect();
                    
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

        ///------------------------------------------------------------------
        /// <summary>
        ///     人件費集計一覧表印刷 </summary>
        ///------------------------------------------------------------------
        private void sReport()
        {
            //エクセルファイル日付明細開始行
            const int S_GYO = 5;
            const int max_GYO = 53;
            int item_Gyo = max_GYO - S_GYO + 1;

            int pCnt = 1;
            int eRow = 0;
            int colorStatus = 1;
            string clrCode = string.Empty;

            try
            {
                //マウスポインタを待機にする
                this.Cursor = Cursors.WaitCursor;

                string sAppPath = System.AppDomain.CurrentDomain.BaseDirectory;

                Excel.Application oXls = new Excel.Application();

                // 印刷用ブック
                Excel.Workbook oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.sxlsPrnPath,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));

                // テンプレートブック
                Excel.Workbook oXlsTempBook = (Excel.Workbook)(oXls.Workbooks.Open(Properties.Settings.Default.sxlsJinPath,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                   Type.Missing, Type.Missing));


                Excel.Worksheet oxls1 = (Excel.Worksheet)oXlsTempBook.Sheets[1];    // 人件費集計一覧表１ページ目
                Excel.Worksheet oxls2 = (Excel.Worksheet)oXlsTempBook.Sheets[2];    // 人件費集計一覧表２ページ以降
                Excel.Worksheet oxlsPrintSheet = null;    // 印刷用ワークシート

                Excel.Range[] rng = new Microsoft.Office.Interop.Excel.Range[2];

                string Category = string.Empty;

                try
                {
                    // 印刷用BOOKへシートを追加する 
                    oxls1.Copy(Type.Missing, oXlsBook.Sheets[pCnt]);

                    // カレントのシートを設定
                    oxlsPrintSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt + 1];

                    // ヘッダ
                    //string tittle1 = Properties.Settings.Default.gengou + txtYear.Text + "年 " + txtMonth.Text + "月分";
                    string tittle1 = txtYear.Text + "年 " + txtMonth.Text + "月分";
                    oxlsPrintSheet.Cells[1, 2] = tittle1;

                    string tittle2 = "土木部員配置＆人件費集計一覧表";

                    if (radioButton1.Checked)
                    {
                        tittle2 += "（個人毎）";
                    }
                    else
                    {
                        tittle2 += "（工事毎）";
                    }

                    oxlsPrintSheet.Cells[2, 2] = tittle2;
                    

                    //グリッドを順番に読む
                    for (int i = 0; i < dg.RowCount; i++)
                    {
                        // ページカウント
                        int p = 0;

                        if (((i + 1) % item_Gyo) > 0)
                        {
                            p = (i + 1) / item_Gyo + 1;
                        }
                        else
                        {
                            p = (i + 1) / item_Gyo;
                        }

                        // 新しいページ？
                        if (pCnt != p)
                        {
                            // ページカウント
                            pCnt = p;

                            // 2ページ以降シートを追加する 
                            oxls2.Copy(Type.Missing, oXlsBook.Sheets[pCnt]);
                            oxlsPrintSheet = (Excel.Worksheet)oXlsBook.Sheets[pCnt + 1];
                        }
                        
                        // 印字行
                        eRow = S_GYO + (i - (item_Gyo * (pCnt - 1)));

                        // セルに値をセット
                        oxlsPrintSheet.Cells[eRow, 1] = dg[cNum, i].Value.ToString();       // 個人コード
                        oxlsPrintSheet.Cells[eRow, 2] = dg[cName, i].Value.ToString();      // 個人名
                        oxlsPrintSheet.Cells[eRow, 3] = dg[cKID, i].Value.ToString();       // 工事コード
                        oxlsPrintSheet.Cells[eRow, 4] = dg[cKName, i].Value.ToString();     // 工事名
                        oxlsPrintSheet.Cells[eRow, 5] = dg[cJinkenhi, i].Value.ToString();  // 人件費
                        oxlsPrintSheet.Cells[eRow, 6] = dg[cHaichiDays, i].Value.ToString();  // 配置日数
                        oxlsPrintSheet.Cells[eRow, 7] = dg[cGenbaDays, i].Value.ToString();  // 現場
                        oxlsPrintSheet.Cells[eRow, 8] = dg[cKinmuchiDays, i].Value.ToString();  // 勤務地
                        oxlsPrintSheet.Cells[eRow, 9] = dg[cStayDays, i].Value.ToString();  // 宿泊
                        oxlsPrintSheet.Cells[eRow, 10] = dg[cHolTM, i].Value.ToString();    // 休日出勤時間
                        oxlsPrintSheet.Cells[eRow, 11] = dg[cHouteiTM, i].Value.ToString(); // 法定休日勤務時間
                        oxlsPrintSheet.Cells[eRow, 12] = dg[cZanTM, i].Value.ToString();    // 残業時間
                        oxlsPrintSheet.Cells[eRow, 13] = dg[cSiTM, i].Value.ToString();     // 深夜残業
                        oxlsPrintSheet.Cells[eRow, 14] = dg[cJyosetsu, i].Value.ToString(); // 除雪手当
                        oxlsPrintSheet.Cells[eRow, 15] = dg[cTokushu, i].Value.ToString();  // 特殊勤務
                        oxlsPrintSheet.Cells[eRow, 16] = dg[cTooshi, i].Value.ToString();   // 通し勤務
                        oxlsPrintSheet.Cells[eRow, 17] = dg[cYakan, i].Value.ToString();    // 夜間手当
                        oxlsPrintSheet.Cells[eRow, 18] = dg[cShokumu, i].Value.ToString();  // 職務手当

                        if (radioButton1.Checked)
                        {
                            if (clrCode != dg[cNum, i].Value.ToString())
                            {
                                clrCode = dg[cNum, i].Value.ToString();
                                colorStatus *= -1;
                            }
                        }
                        else
                        {
                            if (clrCode != dg[cKID, i].Value.ToString())
                            {
                                clrCode = dg[cKID, i].Value.ToString();
                                colorStatus *= -1;
                            }
                        }

                        // 網掛け
                        if (colorStatus == -1)
                        {
                            rng[0] = (Excel.Range)oxlsPrintSheet.Cells[eRow, 1];
                            rng[1] = (Excel.Range)oxlsPrintSheet.Cells[eRow, 9];
                            oxlsPrintSheet.get_Range(rng[0], rng[1]).Interior.Color = Color.LightBlue;
                        }
                    }
                    
                    //マウスポインタを元に戻す
                    this.Cursor = Cursors.Default;

                    // 印刷用BOOKの1番目のシートは削除する
                    ((Excel.Worksheet)oXlsBook.Sheets[1]).Delete();

                    // 確認のためのウィンドウを表示する
                    oXls.Visible = true;

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

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     工事別社員別人件費集計表ＣＳＶデータ出力 ：2018/09/26</summary>
        ///     
        ///---------------------------------------------------------------------
        private void grid2Csv()
        {
            try
            {
                // 出力ファイル削除
                if (System.IO.File.Exists(K_DRAFT_DATA))
                {
                    System.IO.File.Delete(K_DRAFT_DATA);
                }

                // 有効データがないとき
                if (dg.Rows.Count < 3)
                {
                    linkLabel1.Visible = false;
                    return;
                }

                string[] sArray = null;

                StringBuilder sb = new StringBuilder();
                int i = 0;

                for (int iX = 0; iX < dg.Rows.Count; iX++)
                {
                    if (dg[cKID, iX].Value.ToString() == string.Empty)
                    {
                        continue;
                    }

                    sb.Clear();

                    sb.Append(dg[cNum, iX].Value.ToString()).Append(",");    // 社員番号
                    sb.Append(dg[cName, iX].Value.ToString()).Append(",");   // 氏名
                    sb.Append(dg[cKID, iX].Value.ToString()).Append(",");   // コード
                    sb.Append(dg[cKName, iX].Value.ToString()).Append(",");   // 工事名
                    sb.Append(dg[cJinkenhi, iX].Value.ToString().Replace(",", "")).Append(",");   // 人件費計
                    sb.Append(dg[cHaichiDays, iX].Value.ToString()).Append(",");   // 配置日数
                    sb.Append(dg[cGenbaDays, iX].Value.ToString()).Append(",");   // 現場
                    sb.Append(dg[cKinmuchiDays, iX].Value.ToString()).Append(",");   // 勤務地
                    sb.Append(dg[cStayDays, iX].Value.ToString()).Append(",");   // 宿泊
                    sb.Append(dg[cHolTM, iX].Value.ToString()).Append(",");   // 休日勤務
                    sb.Append(dg[cHouteiTM, iX].Value.ToString()).Append(",");   // 法休勤務
                    sb.Append(dg[cZanTM, iX].Value.ToString()).Append(",");   // 普通残業
                    sb.Append(dg[cSiTM, iX].Value.ToString()).Append(",");   // 深夜残業
                    sb.Append(dg[cJyosetsu, iX].Value.ToString()).Append(",");   // 除雪手当
                    sb.Append(dg[cTokushu, iX].Value.ToString()).Append(",");   // 特殊出勤
                    sb.Append(dg[cTooshi, iX].Value.ToString()).Append(",");   // 通し勤務
                    sb.Append(dg[cYakan, iX].Value.ToString()).Append(",");   // 夜間手当
                    sb.Append(dg[cShokumu, iX].Value.ToString());   // 職務手当

                    Array.Resize(ref sArray, i + 1);
                    sArray[i] = sb.ToString();

                    i++;
                }

                // 工事毎人件費集計表CSVファイルを出力する
                txtFileWrite(K_DRAFT_DATA, sArray);

                linkLabel1.Text = linkLabelKouji;   // ラベルを工事にする 2018/09/26
                linkLabel1.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     ＣＳＶファイルを出力する</summary>
        /// <param name="outFilePath">
        ///     出力するフォルダ</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        ///----------------------------------------------------------------------------
        private void txtFileWrite(string sPath, string[] arrayData)
        {
            //// 付加文字列（タイムスタンプ）
            //string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
            //                        DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
            //                        DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

            //// ファイル名
            //string outFileName = @"C:\RYOWA_DATA\xlsOut\koujiShainDraft" + ".csv";

            // テキストファイル出力
            System.IO.File.WriteAllLines(sPath, arrayData, System.Text.Encoding.GetEncoding(932));
        }
        
        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 2018/09/26
            if (linkLabel1.Text == linkLabelKouji)
            {
                getSumDataKouji_201809(dg, Utility.StrtoInt(txtYear.Text), Utility.StrtoInt(txtMonth.Text));
                linkLabel1.Text = linkLabelShain;   // ラベルを社員別にする 2018/09/26
            }
            else
            {
                sumShainByKouji();
            }
        }
    }
}
