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

namespace ryowa_Genba.data
{
    public partial class frmKintai : Form
    {
        public frmKintai(int pNum, DateTime pDt, int pID)
        {
            InitializeComponent();

            // 勤怠データ
            kAdp.FillByYYMM(dts.T_勤怠, pDt.Year, pDt.Month);

            // 休日設定読み込み
            hAdp.Fill(dts.M_休日);

            // 社員読み込み
            sAdp.Fill(dts.M_社員);

            // 工事読み込み
            bAdp.Fill(dts.M_工事);

            // メール設定読み込み
            mAdp.Fill(dts.メール設定);

            // パラメータ読み込み
            _pNUm = pNum;
            _pDate = pDt;
            _pID = pID;
        }

        int _pID = 0;       // ID
        int _pNUm = 0;      // 社員コード
        DateTime _pDate;    // 日付
        Utility.frmMode fMode = new Utility.frmMode();  // フォームモード
        string outCsvFileName = string.Empty;   // 出力勤怠CSVファイル名

        genbaDataSet dts = new genbaDataSet();
        genbaDataSetTableAdapters.T_勤怠TableAdapter kAdp = new genbaDataSetTableAdapters.T_勤怠TableAdapter();
        genbaDataSetTableAdapters.M_社員TableAdapter sAdp = new genbaDataSetTableAdapters.M_社員TableAdapter();
        genbaDataSetTableAdapters.M_休日TableAdapter hAdp = new genbaDataSetTableAdapters.M_休日TableAdapter();
        genbaDataSetTableAdapters.M_工事TableAdapter bAdp = new genbaDataSetTableAdapters.M_工事TableAdapter();
        genbaDataSetTableAdapters.メール設定TableAdapter mAdp = new genbaDataSetTableAdapters.メール設定TableAdapter();

        // 休日フラグ
        bool holDayFlg = false; 

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
                return;
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     勤怠データ表示 </summary>
        /// <param name="sID">
        ///     勤怠データID</param>
        ///------------------------------------------------------------------
        private void dataShow(int sID)
        {
            var s = dts.T_勤怠.Single(a => a.ID == sID);

            lblNum.Text = s.社員ID.ToString();
            lblDate.Text = s.日付.ToShortDateString() + "（" + s.日付.ToString("ddd") + "）";

            if (s.出勤印 == global.flgOn)
            {
                rBtnWork.Checked = true;
            }
            else if (s.休日出勤 == global.flgOn)
            {
                rBtnHolWork.Checked = true;
            }
            else if (s.代休 == global.flgOn)
            {
                rBtnDaiOff.Checked = true;
            }
            else if (s.休日 == global.flgOn)
            {
                rBtnOffDay.Checked = true;
            }
            else if (s.欠勤 == global.flgOn)
            {
                rBtnYasumi.Checked = true;
            }

            if (s.代休対象日 == string.Empty || s.代休対象日.Length < 8)
            {
                dtDDay.Checked = false;
            }
            else
            {
                DateTime dt;
                if (DateTime.TryParse(s.代休対象日.Substring(0, 4) + "/" + s.代休対象日.Substring(4, 2) + "/" + s.代休対象日.Substring(6, 2), out dt))
                {
                    dtDDay.Checked = true;
                    dtDDay.Value = dt;
                }
                else
                {
                    dtDDay.Checked = false;
                }
            }
            
            cmbKouji.SelectedValue = s.工事ID;

            txtInH.Text = s.出社時刻時;
            txtInM.Text = s.出社時刻分;
            txtSH.Text = s.開始時刻時;
            txtSM.Text = s.開始時刻分;
            txtEH.Text = s.終了時刻時;
            txtEM.Text = s.終了時刻分;
            txtOutH.Text = s.退出時刻時;
            txtOutM.Text = s.退出時刻分;

            txtRH.Text = (s.休憩 / 60).ToString();
            txtRM.Text = (s.休憩 % 60).ToString();

            // 早出残業 2018/07/12
            if (s.Is早出残業Null())
            {
                txtHH.Text = string.Empty;
                txtHM.Text = string.Empty;
            }
            else
            {
                txtHH.Text = (s.早出残業 / 60).ToString();
                txtHM.Text = (s.早出残業 % 60).ToString();
            }

            txtZH.Text = (s.普通残業 / 60).ToString();
            txtZM.Text = (s.普通残業 % 60).ToString();

            txtSiH.Text = (s.深夜残業 / 60).ToString();
            txtSiM.Text = (s.深夜残業 % 60).ToString();

            if (s.宿泊 == global.flgOn)
            {
                chkStay.Checked = true;
            }
            else
            {
                chkStay.Checked = false;
            }

            txtKmTuukin.Text = s.通勤業務走行.ToString();
            txtKmShiyou.Text = s.私用走行.ToString();

            if (s.Is備考Null())
            {
                txtMemo.Text = string.Empty;
            }
            else
            {
                txtMemo.Text = s.備考;
            }

            if (s.除雪当番 == global.flgOn)
            {
                chkJyosetsu.Checked = true;
            }
            else
            {
                chkJyosetsu.Checked = false;
            }

            if (s.特殊出勤 == global.flgOn)
            {
                chkTokushu.Checked = true;
            }
            else
            {
                chkTokushu.Checked = false;
            }

            if (s.通し勤務 == global.flgOn)
            {
                chkTooshi.Checked = true;
            }
            else
            {
                chkTooshi.Checked = false;
            }

            if (s.夜間手当 == global.flgOn)
            {
                chkYakan.Checked = true;
            }
            else
            {
                chkYakan.Checked = false;
            }

            if (s.職務手当 == global.flgOn)
            {
                chkShokumu.Checked = true;
            }
            else
            {
                chkShokumu.Checked = false;
            }

            if (s.確認印 == global.flgOn)
            {
                chkManager.Checked = true;
            }
            else
            {
                chkManager.Checked = false;
            }

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            //DateTime dt = DateTime.Parse(dateTimePicker1.Value.ToShortDateString());
         
        }

        private void dateChange(DateTime dt)
        {
            if (dts.M_休日.Any(a => a.日付 == dt))
            {
                // 休日の時
                holDayFlg = true;
                rBtnWork.Enabled = false;
                rBtnDaiOff.Enabled = false;
                rBtnYasumi.Enabled = false;
                rBtnHolWork.Enabled = true;
                rBtnOffDay.Enabled = true;
                rBtnOffDay.Checked = true;
                dtDDay.Checked = false;
                
                lblDate.ForeColor = Color.OrangeRed;
                lblCalender.BackColor = Color.OrangeRed;
                lblCalender.ForeColor = SystemColors.Info;

                var s = dts.M_休日.Single(a => a.日付 == dt);
                lblCalender.Text = s.摘要;
            }
            else
            {
                // 勤務日の時
                holDayFlg = false;
                rBtnWork.Enabled = true;
                rBtnHolWork.Enabled = false;
                rBtnDaiOff.Enabled = true;
                rBtnYasumi.Enabled = true;
                rBtnWork.Checked = true;
                rBtnOffDay.Enabled = false;
                dtDDay.Checked = false;
                
                lblDate.ForeColor = Color.SteelBlue;
                lblCalender.BackColor = Color.SteelBlue;
                lblCalender.ForeColor = SystemColors.Info;
                lblCalender.Text = "通常勤務";
            }
        }


        ///--------------------------------------------------------------------
        /// <summary>
        ///     前日と同じ工事／所属部署をデフォルト表示する </summary>
        /// <param name="dt">
        ///     当日の日付</param>
        ///--------------------------------------------------------------------
        private void cmbKoujiSetIndex(DateTime dt)
        {
            DateTime _dt = DateTime.Parse(dt.ToShortDateString()).AddDays(-1);

            if (dts.T_勤怠.Any(a => a.日付 == _dt && a.社員ID == _pNUm))
            {
                var s = dts.T_勤怠.Single(a => a.日付 == _dt && a.社員ID == _pNUm);
                cmbKouji.SelectedValue = s.工事ID; 
            }
        }

        ///------------------------------------------------
        /// <summary>
        ///     出勤日画面初期化 </summary>
        ///------------------------------------------------
        private void workdayInitial()
        {
            // 出勤日、休日出勤日
            dtDDay.Checked = false;
            dtDDay.Enabled = false;
            
            cmbKouji.Enabled = true;    // 工事所属部署          
            panel3.Enabled = true;      // 勤務実績
            chkStay.Enabled = true;     // 宿泊           
            txtMemo.Enabled = true;     // 所定時間内で処理できない業務内容・他 特記事項            
            panel4.Enabled = true;      // 走行距離
        }

        private void holidayInitial()
        {
            // 代休・休日・休み
            cmbKouji.Enabled = true;    // 工事所属部署        
            panel3.Enabled = false;     // 勤務実績
            chkStay.Enabled = false;    // 宿泊     
            txtMemo.Enabled = true;     // 所定時間内で処理できない業務内容・他 特記事項  
            panel4.Enabled = true;      // 走行距離

            panel2.Enabled = false;     // 休日は特殊勤務チェックは不可
        }

        private void rBtnWork_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnWork.Checked)
            {
                workdayInitial();
            }
        }

        private void rBtnHolWork_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnHolWork.Checked)
            {
                workdayInitial();
            }
        }

        private void rBtnDaiOff_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnDaiOff.Checked)
            {
                holidayInitial();

                // 代休出勤日
                dtDDay.Checked = true;
                dtDDay.Enabled = true;
            }
        }

        private void rBtnOffDay_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnOffDay.Checked)
            {
                holidayInitial();

                // 代休出勤日
                dtDDay.Checked = false;
                dtDDay.Enabled = false;
            }
        }

        private void rBtnYasumi_CheckedChanged(object sender, EventArgs e)
        {
            if (rBtnYasumi.Checked)
            {
                holidayInitial();

                // 代休出勤日
                dtDDay.Checked = false;
                dtDDay.Enabled = false;
            }
        }

        private void txtNum_TextChanged(object sender, EventArgs e)
        {

        }

        private void frmKintai_Load(object sender, EventArgs e)
        {
            // 工事コンボ
            var ds = dts.M_工事.Select(a => new
            {
                ID = a.ID,
                NAME = a.名称
            }).ToList();

            cmbKouji.DataSource = ds;
            cmbKouji.DisplayMember = "NAME";
            cmbKouji.ValueMember = "ID";
            cmbKouji.SelectedIndex = -1;

            // 画面初期化
            dispInitial();

            // 日付（勤務日・休日）による画面制御
            dateChange(_pDate);

            // 新規登録のとき
            if (_pID == global.flgOff)
            {
                // 社員コード
                lblNum.Text = _pNUm.ToString();

                // 日付
                lblDate.Text = _pDate.ToShortDateString() + "（" + _pDate.ToString("ddd") + "）";

                // フォームモード
                fMode.Mode = global.FORM_ADDMODE;
                linkLabel4.Text = "新規登録する";

                // 工事コンボ
                cmbKoujiSetIndex(_pDate);
            }
            else
            {
                // 既存データ表示
                dataShow(_pID);

                // フォームモード
                fMode.Mode = global.FORM_EDITMODE;
                linkLabel4.Text = "出勤簿の更新";
            }
        }

        ///-------------------------------------------------------
        /// <summary>
        ///     画面初期化 </summary>
        ///-------------------------------------------------------
        private void dispInitial()
        {
            txtInH.Text = string.Empty;
            txtInM.Text = string.Empty;
            txtSH.Text = Properties.Settings.Default.startH;
            txtSM.Text = Properties.Settings.Default.startM.PadLeft(2, '0');
            txtEH.Text = Properties.Settings.Default.endH;
            txtEM.Text = Properties.Settings.Default.endM.PadLeft(2, '0');
            txtOutH.Text = string.Empty;
            txtOutM.Text = string.Empty;
            txtRH.Text = string.Empty;
            txtRM.Text = string.Empty;
            txtHH.Text = string.Empty;  // 2018/07/12
            txtHM.Text = string.Empty;  // 2018/07/12
            txtZH.Text = string.Empty;
            txtZM.Text = string.Empty;
            txtSiH.Text = string.Empty;
            txtSiM.Text = string.Empty;
            chkStay.Checked = false;

            txtKmTuukin.Text = string.Empty;
            txtKmShiyou.Text = string.Empty;

            txtMemo.Text = string.Empty;

            chkJyosetsu.Checked = false;
            chkTokushu.Checked = false;
            chkTooshi.Checked = false;
            chkYakan.Checked = false;
            chkShokumu.Checked = false;

            chkManager.Checked = false;

            // 勤怠編集権限の有無
            if (global.loginType == global.flgOff)
            {
                // 自らのみの勤怠編集権限
                //panel2.Enabled = false;
                //chkManager.Enabled = false;
                label25.Visible = false;
                panel2.Visible = false;
                chkManager.Enabled = false;
            }
            else if (global.loginType == global.flgOn)
            {
                // 特殊勤務・管理者確認の勤怠編集権限あり
                label25.Visible = true;
                panel2.Visible = true;
                chkManager.Enabled = true;
            }
        }

        private void txtSH_TextChanged(object sender, EventArgs e)
        {
            DateTime dt;

            if (txtSH.Text == string.Empty || txtSM.Text == string.Empty ||
                txtEH.Text == string.Empty || txtEM.Text == string.Empty)
            {
                // 勤務開始時刻、勤務終了時刻が未入力の時、残業・深夜時間欄を空白とする
                txtZH.Text = string.Empty;
                txtZM.Text = string.Empty;
                txtSiH.Text = string.Empty;
                txtSiM.Text = string.Empty;

                // 同じく早出残業欄を空白とする   2018/07/26
                txtHH.Text = string.Empty;
                txtHM.Text = string.Empty;

                return;
            }

            if (!DateTime.TryParse(txtSH.Text + ":" + txtSM.Text, out dt) ||
                !DateTime.TryParse(txtEH.Text + ":" + txtEM.Text, out dt))
            {
                // 勤務開始時刻、勤務終了時刻が不備の時、残業・深夜時間欄を空白とする
                txtZH.Text = string.Empty;
                txtZM.Text = string.Empty;
                txtSiH.Text = string.Empty;
                txtSiM.Text = string.Empty;

                // 同じく早出残業欄を空白とする   2018/07/26
                txtHH.Text = string.Empty;
                txtHM.Text = string.Empty;

                return;
            }

            //// 残業時間計算　：　出勤日のみ（※休日出勤は該当しない）
            //DateTime eTM = global.dt1700;
            //if (rBtnWork.Checked)
            //{
            //    if (DateTime.Parse(txtSH.Text + ":" + txtSM.Text) > global.dt1700)
            //    {
            //        eTM = DateTime.Parse(txtSH.Text + ":" + txtSM.Text);
            //    }

            //    double z = getZangyoTime(eTM, Utility.StrtoInt(txtSH.Text), Utility.StrtoInt(txtSM.Text), Utility.StrtoInt(txtEH.Text), Utility.StrtoInt(txtEM.Text), Properties.Settings.Default.marume);
            //    txtZH.Text = ((int)(z / 60)).ToString();
            //    txtZM.Text = (z % 60).ToString();
            //}

            // 残業時間計算：2018/07/11
            if (rBtnWork.Checked)
            {
                double z = 0;

                // 早出残業
                z = getHayadeTime(Utility.StrtoInt(txtSH.Text), Utility.StrtoInt(txtSM.Text), Utility.StrtoInt(txtEH.Text), Utility.StrtoInt(txtEM.Text), Properties.Settings.Default.marume);
                txtHH.Text = ((int)(z / 60)).ToString();
                txtHM.Text = (z % 60).ToString();

                // 普通残業
                z = getZangyoTime2018(Utility.StrtoInt(txtSH.Text), Utility.StrtoInt(txtSM.Text), Utility.StrtoInt(txtEH.Text), Utility.StrtoInt(txtEM.Text), Properties.Settings.Default.marume);
                txtZH.Text = ((int)(z / 60)).ToString();
                txtZM.Text = (z % 60).ToString();
            }

            // 深夜残業計算
            double sz = getShinyaTime(Utility.StrtoInt(txtSH.Text), Utility.StrtoInt(txtSM.Text), Utility.StrtoInt(txtEH.Text), Utility.StrtoInt(txtEM.Text), Properties.Settings.Default.marume);
            txtSiH.Text = ((int)(sz / 60)).ToString();
            txtSiM.Text = (sz % 60).ToString();

            // 通し勤務
            chkTooshi.Checked = getTooshiKinmu(Utility.StrtoInt(txtSH.Text), Utility.StrtoInt(txtSM.Text), Utility.StrtoInt(txtEH.Text), Utility.StrtoInt(txtEM.Text));

            // 夜間手当
            chkYakan.Checked = getYakanTeate(Utility.StrtoInt(txtSH.Text), Utility.StrtoInt(txtSM.Text), Utility.StrtoInt(txtEH.Text), Utility.StrtoInt(txtEM.Text));
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間取得(22:00～05:00) </summary>
        /// <param name="cSH">
        ///     勤務開始時刻・時</param>
        /// <param name="cSM">
        ///     勤務開始時刻・分</param>
        /// <param name="cEH">
        ///     勤務終了時刻・時</param>
        /// <param name="cEM">
        ///     勤務終了時刻・分</param>
        /// <param name="Maru">
        ///     丸め単位・分</param>
        /// <returns>
        ///     深夜勤務時間・分</returns>
        ///-------------------------------------------------------------------
        private double getHayadeTime(int cSH, int cSM, int cEH, int cEM, int maru)
        {
            DateTime stTM;
            DateTime edTM;

            // 勤務開始時刻
            if (!DateTime.TryParse(cSH.ToString() + ":" + cSM.ToString(), out stTM))
            {
                return 0;
            }

            // 勤務終了時刻
            if (!DateTime.TryParse(cEH.ToString() + ":" + cEM.ToString(), out edTM))
            {
                return 0;
            }

            double spanMin = 0;

            int sTime = cSH * 100 + cSM;
            int eTime = cEH * 100 + cEM;

            // 終了が翌日のとき
            if (sTime > eTime)
            {
                edTM = edTM.AddDays(1);
            }

            //// 開始が当日8:00以前のとき
            //if (cSH < global.dt0800.Hour && cSM < 60)
            //{
            //    // 開始時刻を取得します
            //    if (cSH < global.dt0500.Hour)
            //    {
            //        stTM = global.dt0500;
            //    }
            //    else
            //    {
            //        stTM = DateTime.Parse(cSH.ToString() + ":" + cSM.ToString());
            //    }

            //    // 終了時刻を取得します
            //    if (cEH >= global.dt0800.Hour)
            //    {
            //        edTM = global.dt0800;
            //    }
            //    else
            //    {
            //        edTM = DateTime.Parse(cEH.ToString() + ":" + cEM.ToString());
            //    }

            //    // 早出勤務時間
            //    spanMin += Utility.GetTimeSpan(stTM, edTM).TotalMinutes;
            //}

            // 2018/07/26
            if (stTM < global.dt0800)
            {
                // 開始時刻を確定します
                if (stTM < global.dt0500)
                {
                    stTM = global.dt0500;
                }

                // 終了時刻を確定します
                if (edTM >= global.dt0800)
                {
                    edTM = global.dt0800;
                }

                // 早出勤務時間
                spanMin += Utility.GetTimeSpan(stTM, edTM).TotalMinutes;
            }

            // 単位時間で丸める
            spanMin -= spanMin % maru;

            return spanMin;
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     夜間手当該当チェック取得 </summary>
        /// <param name="cSH">
        ///     勤務開始時刻・時 </param>
        /// <param name="cSM">
        ///     勤務開始時刻・分</param>
        /// <param name="cEH">
        ///     勤務終了時刻・時</param>
        /// <param name="cEM">
        ///     勤務終了時刻・分</param>
        /// <returns>
        ///     true:夜間手当に該当、false:夜間手当に該当しない</returns>
        ///------------------------------------------------------------------
        private bool getYakanTeate(int cSH, int cSM, int cEH, int cEM)
        {
            if (cEH == 24)
            {
                cEH = 0;
            }

            bool rtn = false;

            int sTime = cSH * 100 + cSM;
            int eTime = cEH * 100 + cEM;

            // 勤務時間取得
            if (sTime >= eTime)
            {
                // 終了が翌日のとき
                if (eTime > 0)
                {
                    rtn = true;
                }
            }

            return rtn;
        }

        ///---------------------------------------------------------------------------
        /// <summary>
        ///     通し勤務該当チェック取得 </summary>
        /// <param name="cSH">
        ///     勤務開始時刻・時</param>
        /// <param name="cSM">
        ///     勤務開始時刻・分</param>
        /// <param name="cEH">
        ///     勤務終了時刻・時</param>
        /// <param name="cEM">
        ///     勤務終了時刻・分</param>
        /// <returns>
        ///     true:通し勤務に該当、false:通し勤務に該当しない</returns>
        ///---------------------------------------------------------------------------
        private bool getTooshiKinmu(int cSH, int cSM, int cEH, int cEM)
        {
            if (cEH == 24)
            {
                cEH = 0;
            }

            DateTime stTM = DateTime.Parse(cSH.ToString() + ":" + cSM.ToString());
            DateTime edTM = DateTime.Parse(cEH.ToString() + ":" + cEM.ToString());
            double spanMin = 0;
            bool rtn = false;

            int sTime = cSH * 100 + cSM;
            int eTime = cEH * 100 + cEM;

            // 勤務時間取得
            if (sTime >= eTime)
            {
                // 終了が翌日のとき
                spanMin += Utility.GetTimeSpan(stTM, edTM.AddDays(1)).TotalMinutes;
            }
            else
            {
                // 終了が当日のとき
                spanMin += Utility.GetTimeSpan(stTM, edTM).TotalMinutes;
            }

            // 19時間以上の時、通し勤務に該当
            if (spanMin >= (Properties.Settings.Default.tooshiH * 60))
            {
                rtn = true;
            }

            return rtn;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     残業時間取得</summary>
        /// <param name="eTm">
        ///     勤務終了定時の時刻</param>
        /// <param name="cSH">
        ///     勤務開始時刻・時</param>
        /// <param name="cSM">
        ///     勤務開始時刻・分</param>
        /// <param name="cEH">
        ///     勤務終了時刻・時</param>
        /// <param name="cEM">
        ///     勤務終了時刻・分</param>
        /// <param name="Maru">
        ///     丸め単位・分</param>
        /// <returns>
        ///     残業時間・分</returns>
        ///-------------------------------------------------------------------
        private double getZangyoTime2018(int cSH, int cSM, int cEH, int cEM, int Maru)
        {
            DateTime stTM;
            DateTime edTM;
            DateTime eTm;
            double spanMin = 0;

            if (cEH == 24)
            {
                cEH = 0;
            }

            // 勤務開始時刻
            stTM = DateTime.Parse(cSH.ToString() + ":" + cSM.ToString());

            /* 開始時間が8:00以前のとき、残業時間算出起点開始時刻は8:00とする
             * ※早出残業時間と重複するため 
             * 2018/07/26 */
            if (stTM < global.dt0800)
            {
                stTM = global.dt0800;
            }

            // 勤務終了時刻
            edTM = DateTime.Parse(cEH.ToString() + ":" + cEM.ToString());

            int sTime = cSH * 100 + cSM;
            int eTime = cEH * 100 + cEM;

            if (sTime > eTime)
            {
                // 終了が翌日のとき
                edTM = edTM.AddDays(1);
            }

            double wTM = Utility.GetTimeSpan(stTM, edTM).TotalMinutes;

            // 開始～終了10時間を超過しているとき残業
            if (wTM > 600)
            {
                // 翌日朝の5:00
                DateTime nxtDT500 = global.dt0500.AddDays(1);

                // 勤務開始から10時間の時刻を取得します
                eTm = stTM.AddHours(10);

                // 残業時間を取得します
                spanMin += Utility.GetTimeSpan(eTm, edTM).TotalMinutes;

                //if (eTm < global.dt2200)
                //{
                //    // 当日22:00まで
                //    if (edTM <= global.dt2200)
                //    {
                //        spanMin += Utility.GetTimeSpan(eTm, edTM).TotalMinutes;
                //    }
                //    else if (edTM > global.dt2200)
                //    {
                //        spanMin += Utility.GetTimeSpan(eTm, global.dt2200).TotalMinutes;
                //    }
                //}
                //else if (edTM > nxtDT500)
                //{
                //    // 翌日5:00以降
                //    spanMin += Utility.GetTimeSpan(nxtDT500, edTM).TotalMinutes;
                //}

                // 単位時間で丸める
                spanMin -= (spanMin % Maru);
            }

            return spanMin;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     残業時間取得</summary>
        /// <param name="eTm">
        ///     勤務終了定時の時刻</param>
        /// <param name="cSH">
        ///     勤務開始時刻・時</param>
        /// <param name="cSM">
        ///     勤務開始時刻・分</param>
        /// <param name="cEH">
        ///     勤務終了時刻・時</param>
        /// <param name="cEM">
        ///     勤務終了時刻・分</param>
        /// <param name="Maru">
        ///     丸め単位・分</param>
        /// <returns>
        ///     残業時間・分</returns>
        ///-------------------------------------------------------------------
        private double getZangyoTime(DateTime eTm, int cSH, int cSM, int cEH, int cEM, int Maru)
        {
            DateTime stTM;
            DateTime edTM;
            double spanMin = 0;
            
            if (cEH == 24) cEH = 0;

            stTM = DateTime.Parse(cSH.ToString() + ":" + cSM.ToString());
            edTM = DateTime.Parse(cEH.ToString() + ":" + cEM.ToString());

            int sTime = cSH * 100 + cSM;
            int eTime = cEH * 100 + cEM;

            // 終了が翌日のとき
            if (sTime > eTime)
            {
                spanMin += Utility.GetTimeSpan(eTm, global.dt2200).TotalMinutes;
            }
            else
            {
                if (edTM <= eTm)
                {
                    return 0;
                }

                if (edTM <= global.dt2200)
                {
                    spanMin += Utility.GetTimeSpan(eTm, edTM).TotalMinutes;
                }
                else if (edTM > global.dt2200)
                {
                    spanMin += Utility.GetTimeSpan(eTm, global.dt2200).TotalMinutes;
                }
            }

            // 単位時間で丸める
            spanMin -= (spanMin % Maru);

            return spanMin;
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     深夜勤務時間取得(22:00～05:00) </summary>
        /// <param name="cSH">
        ///     勤務開始時刻・時</param>
        /// <param name="cSM">
        ///     勤務開始時刻・分</param>
        /// <param name="cEH">
        ///     勤務終了時刻・時</param>
        /// <param name="cEM">
        ///     勤務終了時刻・分</param>
        /// <param name="Maru">
        ///     丸め単位・分</param>
        /// <returns>
        ///     深夜勤務時間・分</returns>
        ///-------------------------------------------------------------------
        private double getShinyaTime(int cSH, int cSM, int cEH, int cEM, int maru)
        {
            int wHour = 0;
            int wMin = 0;
            int sKyukei = 0;
            
            DateTime stTM;
            DateTime edTM;
            double spanMin = 0;

            // 開始時刻の取得
            wHour = cSH;
            wMin = cSM;

            // 開始が５：００以前のとき
            if (wHour == 24) wHour = 0;
            if (wHour < 5 && wMin < 60)
            {
                // 深夜勤務時間
                stTM = DateTime.Parse(wHour.ToString() + ":" + wMin.ToString());
                spanMin += Utility.GetTimeSpan(stTM, global.dt0500).TotalMinutes;
            }

            // 終了時刻の取得
            wHour = cEH;
            wMin = cEM;
            
            int sTime = cSH * 100 + cSM;
            int eTime = cEH * 100 + cEM;

            // 終了が翌日のとき
            if (sTime > eTime)
            {
                if (wHour >= 5)
                {
                    edTM = global.dt0500.AddDays(1);
                }
                else
                {
                    edTM = DateTime.Parse(wHour.ToString() + ":" + wMin.ToString()).AddDays(1);
                }

                spanMin += Utility.GetTimeSpan(global.dt2200, edTM).TotalMinutes;
            }
            else
            {
                // 終了が当日内で22:00以降のとき
                if (wHour >= 22)
                {
                    edTM = DateTime.Parse(wHour.ToString() + ":" + wMin.ToString());
                    spanMin += Utility.GetTimeSpan(global.dt2200, edTM).TotalMinutes;
                }
            }

            // 深夜勤務時間
            spanMin -= sKyukei;

            // 単位時間で丸める
            spanMin -= spanMin % maru;

            return spanMin;
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.Close();
        }

        private void frmKintai_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // エラーチェック
            if (!errCheck())
            {
                return;
            }

            // 登録確認メッセージ
            if (MessageBox.Show("勤怠を登録します。よろしいですか？","登録確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (fMode.Mode == global.FORM_ADDMODE)
            {
                // データ登録
                dataAdd();
            }
            else if (fMode.Mode == global.FORM_EDITMODE)
            {
                // データ更新
                dataUpdate(_pID);
            }

            // データベース更新
            kAdp.Update(dts.T_勤怠);

            // 勤怠データをメール送信
            if (System.IO.File.Exists(outCsvFileName))
            {
                // メール件名
                string sbj = "<" + _pNUm.ToString() + " " + lblName.Text + "> " + _pDate.ToShortDateString();

                // メール本文
                string sBody = lblNum.Text + " " + lblName.Text + Environment.NewLine + _pDate.Year.ToString() + "年" + _pDate.Month.ToString() + "月" + _pDate.Day.ToString() + "日 出勤簿データを送付します。";

                // 送信
                Utility.sendKintaiMail(outCsvFileName, sbj, sBody, 0);
            }

            // 画面を閉じる
            this.Close();
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     登録前チェック </summary>
        /// <returns>
        ///     true:エラーなし、false:エラー有り</returns>
        ///-------------------------------------------------------------------
        private bool errCheck()
        {
            if (cmbKouji.SelectedIndex == -1)
            {
                MessageBox.Show("工事／所属部署を選択してください", "工事部署未選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                cmbKouji.Focus();
                return false;
            }

            // 出勤日または休日出勤
            if (rBtnWork.Checked || rBtnHolWork.Checked)
            {
                if (txtInH.Text == string.Empty)
                {
                    MessageBox.Show("出社時刻を入力してください", "出社時刻未入力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtInH.Focus();
                    return false;
                }

                if (Utility.StrtoInt(txtInH.Text) > 23)
                {
                    MessageBox.Show("出社時刻は０～23の範囲で入力してください", "出社時刻不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtInH.Focus();
                    return false;
                }

                if (txtInM.Text == string.Empty)
                {
                    MessageBox.Show("出社時刻を入力してください", "出社時刻未入力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtInM.Focus();
                    return false;
                }

                if (Utility.StrtoInt(txtInM.Text) > 59)
                {
                    MessageBox.Show("時刻・分は０～59の範囲で入力してください", "出社時刻不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtInM.Focus();
                    return false;
                }

                if (txtSH.Text == string.Empty)
                {
                    MessageBox.Show("勤務開始時刻を入力してください", "勤務開始時刻未入力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtSH.Focus();
                    return false;
                }

                if (Utility.StrtoInt(txtSH.Text) > 23)
                {
                    MessageBox.Show("勤務開始時刻は０～23の範囲で入力してください", "勤務開始時刻不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtSH.Focus();
                    return false;
                }

                if (txtSM.Text == string.Empty)
                {
                    MessageBox.Show("勤務開始時刻を入力してください", "勤務開始時刻未入力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtSM.Focus();
                    return false;
                }

                if (Utility.StrtoInt(txtSM.Text) > 59)
                {
                    MessageBox.Show("時刻・分は０～59の範囲で入力してください", "開始時刻不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtSM.Focus();
                    return false;
                }

                if (txtEH.Text == string.Empty)
                {
                    MessageBox.Show("勤務終了時刻を入力してください", "勤務終了時刻未入力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtEH.Focus();
                    return false;
                }

                if (Utility.StrtoInt(txtEH.Text) > 23)
                {
                    MessageBox.Show("勤務終了時刻は０～23の範囲で入力してください", "勤務終了時刻不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtEH.Focus();
                    return false;
                }

                if (txtEM.Text == string.Empty)
                {
                    MessageBox.Show("勤務終了時刻を入力してください", "勤務終了時刻未入力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtEM.Focus();
                    return false;
                }

                if (Utility.StrtoInt(txtEM.Text) > 59)
                {
                    MessageBox.Show("時刻・分は０～59の範囲で入力してください", "勤務終了時刻不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtEM.Focus();
                    return false;
                }

                if (txtOutH.Text == string.Empty)
                {
                    MessageBox.Show("退社時刻を入力してください", "退社時刻未入力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtOutH.Focus();
                    return false;
                }

                if (Utility.StrtoInt(txtOutH.Text) > 23)
                {
                    MessageBox.Show("退社時刻は０～23の範囲で入力してください", "退社時刻不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtOutH.Focus();
                    return false;
                }

                if (txtOutM.Text == string.Empty)
                {
                    MessageBox.Show("退社時刻を入力してください", "退社時刻未入力", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtOutM.Focus();
                    return false;
                }

                if (Utility.StrtoInt(txtOutM.Text) > 59)
                {
                    MessageBox.Show("時刻・分は０～59の範囲で入力してください", "退社時刻不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtOutM.Focus();
                    return false;
                }

                if (Utility.StrtoInt(txtRM.Text) > 59)
                {
                    MessageBox.Show("時刻・分は０～59の範囲で入力してください", "休憩時間不備", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtRM.Focus();
                    return false;
                }

                // 勤務時間取得
                DateTime sDt = DateTime.Parse(txtSH.Text + ":" + txtSM.Text);
                DateTime eDt = DateTime.Parse(txtEH.Text + ":" + txtEM.Text);

                int st = Utility.StrtoInt(txtSH.Text) * 100 + Utility.StrtoInt(txtSM.Text);
                int et = Utility.StrtoInt(txtEH.Text) * 100 + Utility.StrtoInt(txtEM.Text);

                if (st >= et)
                {
                    eDt = eDt.AddDays(1);
                }

                //double w = Utility.GetTimeSpan(sDt, eDt).TotalMinutes - 75;

                // 2018/07/12
                double w = Utility.GetTimeSpan(sDt, eDt).TotalMinutes;

                if (w > Properties.Settings.Default.restTime)
                {
                    w -= Properties.Settings.Default.restTime;
                }

                double r = 0;

                // 休憩時間チェック
                r = Utility.StrtoInt(txtRH.Text) * 60 + Utility.StrtoInt(txtRM.Text);

                if (r >= w)
                {
                    // 2018/07/12
                    MessageBox.Show("休憩時間が勤務時間（" + Properties.Settings.Default.restTime + "分休憩を除く）以上になっています", "休憩時間確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtRH.Focus();
                    return false;
                }

                // 残業時間チェック
                r = Utility.StrtoInt(txtZH.Text) * 60 + Utility.StrtoInt(txtZM.Text);

                if (r >= w)
                {
                    // 2018/07/12
                    MessageBox.Show("残業時間が勤務時間（" + Properties.Settings.Default.restTime + "分休憩を除く）以上になっています", "残業時間確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtZH.Focus();
                    return false;
                }

                // 深夜残業チェック
                double siW = getShinyaTime(Utility.StrtoInt(txtSH.Text), Utility.StrtoInt(txtSM.Text), Utility.StrtoInt(txtEH.Text), Utility.StrtoInt(txtEM.Text), Properties.Settings.Default.marume);
                r = Utility.StrtoInt(txtSiH.Text) * 60 + Utility.StrtoInt(txtSiM.Text);

                if (siW != r)
                {
                    string m = "深夜時間が異なっています。よろしいですか？" + Environment.NewLine + Environment.NewLine +
                               "計算：" + Utility.intToHhMM((int)(siW)) + Environment.NewLine +
                               "入力：" + Utility.StrtoInt(txtSiH.Text).ToString() + ":" + txtSiM.Text.PadLeft(2, '0');

                    if (MessageBox.Show(m, "深夜時間確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        txtSiH.Focus();
                        return false;
                    }
                }
            }

            // 代休のとき
            if (rBtnDaiOff.Checked)
            {
                if (!dtDDay.Checked)
                {
                    if (MessageBox.Show("休日出勤日が未選択ですがよろしいですか？", "対象休日出勤日", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.No)
                    {
                        dtDDay.Focus();
                        return false;
                    }
                }
                else
                {
                    // 代休の対象となる休日出勤日
                    if (!dts.T_勤怠.Any(a => a.社員ID == _pNUm && 
                            a.日付 == DateTime.Parse(dtDDay.Value.ToShortDateString()) && 
                            a.休日出勤 == global.flgOn))
                    {
                        MessageBox.Show("休日出勤した日を選択してください", "対象休日出勤日", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        dtDDay.Focus();
                        return false;
                    }

                    // 既にその日の代休は取得済みのとき
                    if (dts.T_勤怠.Any(a => a.社員ID == _pNUm && a.代休対象日 == dtDDay.Value.ToShortDateString()))
                    {
                        MessageBox.Show("既に代休を取得した休日出勤日です", "対象休日出勤日", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        dtDDay.Focus();
                        return false;
                    }
                }
            }

            return true;
        }

        private bool dataAny()
        {
            if (dts.T_勤怠.Any(a => a.社員ID == _pNUm && a.日付 == DateTime.Parse(lblDate.Text)))
            {
                MessageBox.Show(lblName.Text + "さんの" + lblDate.Text + "の勤怠は入力済みです","", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return false;
            }

            return true;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     データ更新 </summary>
        /// <param name="sID">
        ///     T_勤怠@ID</param>
        ///----------------------------------------------------------------------
        private void dataUpdate(int sID)
        {
            genbaDataSet.T_勤怠Row s = dts.T_勤怠.Single(a => a.ID == sID);
            rowDataSet(s);

            // 自身のメールアドレスを取得
            string mlAdd = string.Empty;
            if (dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                var ml = dts.メール設定.Single(a => a.ID == global.mailKey);
                mlAdd = ml.メールアドレス;
            }

            // 勤怠CSVデータ出力
            outCsvFileName = putKintaiCsv(s, mlAdd);
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     データ登録 </summary>
        ///------------------------------------------------------------------
        private void dataAdd()
        {
            genbaDataSet.T_勤怠Row s = dts.T_勤怠.NewT_勤怠Row();
            rowDataSet(s);
            dts.T_勤怠.AddT_勤怠Row(s);

            // 自身のメールアドレスを取得
            string mlAdd = string.Empty;
            if (dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                var ml = dts.メール設定.Single(a => a.ID == global.mailKey);
                mlAdd = ml.メールアドレス;
            }

            // 勤怠CSVデータ出力
            outCsvFileName = putKintaiCsv(s, mlAdd);
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     送信用出勤簿CSVデータ作成 </summary>
        /// <param name="s">
        ///     genbaDataSet.T_勤怠Row</param>
        /// <param name="mlAdd">
        ///     送信元メールアドレス</param>
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

            string[] st = new string[1];
            st[0] = sb.ToString();

            // 受け渡しデータパスファイル名
            string outFileName = Properties.Settings.Default.attachPath + s.社員ID.ToString() + "_" + s.日付.Year.ToString() + s.日付.Month.ToString().PadLeft(2, '0') + s.日付.Day.ToString().PadLeft(2, '0') + ".csv";

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


        ///------------------------------------------------------------------
        /// <summary>
        ///     T_勤怠Rowにデータをセットします </summary>
        /// <param name="s">
        ///     genbaDataSet.T_勤怠Row </param>
        ///------------------------------------------------------------------
        private void rowDataSet(genbaDataSet.T_勤怠Row s)
        {
            s.日付 = _pDate;
            s.社員ID = _pNUm;
            s.工事ID = Utility.StrtoInt(cmbKouji.SelectedValue.ToString());
            
            if (rBtnWork.Checked)
            {
                s.出勤印 = global.flgOn;
            }
            else
            {
                s.出勤印 = global.flgOff;
            }

            if (rBtnWork.Checked || rBtnHolWork.Checked)
            {
                s.出社時刻時 = txtInH.Text;
                s.出社時刻分 = txtInM.Text;
                s.開始時刻時 = txtSH.Text;
                s.開始時刻分 = txtSM.Text;
                s.終了時刻時 = txtEH.Text;
                s.終了時刻分 = txtEM.Text;
                s.退出時刻時 = txtOutH.Text;
                s.退出時刻分 = txtOutM.Text;
                s.休憩 = Utility.StrtoInt(txtRH.Text) * 60 + Utility.StrtoInt(txtRM.Text);
                s.普通残業 = Utility.StrtoInt(txtZH.Text) * 60 + Utility.StrtoInt(txtZM.Text);
                s.深夜残業 = Utility.StrtoInt(txtSiH.Text) * 60 + Utility.StrtoInt(txtSiM.Text);
                s.早出残業 = Utility.StrtoInt(txtHH.Text) * 60 + Utility.StrtoInt(txtHM.Text);      // 2018/07/12
            }
            else
            {
                s.出社時刻時 = string.Empty;
                s.出社時刻分 = string.Empty;
                s.開始時刻時 = string.Empty;
                s.開始時刻分 = string.Empty;
                s.終了時刻時 = string.Empty;
                s.終了時刻分 = string.Empty;
                s.退出時刻時 = string.Empty;
                s.退出時刻分 = string.Empty;
                s.休憩 = global.flgOff;
                s.普通残業 = global.flgOff;
                s.深夜残業 = global.flgOff;
                s.早出残業 = global.flgOff;     // 2018/07/11
            }

            if (rBtnHolWork.Checked)
            {
                s.休日出勤 = global.flgOn;
            }
            else
            {
                s.休日出勤 = global.flgOff;
            }

            if (rBtnDaiOff.Checked)
            {
                s.代休 = global.flgOn;
            }
            else
            {
                s.代休 = global.flgOff;
            }

            if (rBtnOffDay.Checked)
            {
                s.休日 = global.flgOn;
            }
            else
            {
                s.休日 = global.flgOff;
            }

            if (rBtnYasumi.Checked)
            {
                s.欠勤 = global.flgOn;
            }
            else
            {
                s.欠勤 = global.flgOff;
            }

            if (chkStay.Checked)
            {
                s.宿泊 = global.flgOn;
            }
            else
            {
                s.宿泊 = global.flgOff;
            }

            s.備考 = txtMemo.Text.Replace(",","");

            if (chkJyosetsu.Checked)
            {
                s.除雪当番 = global.flgOn;
            }
            else
            {
                s.除雪当番 = global.flgOff;
            }

            if (chkTokushu.Checked)
            {
                s.特殊出勤 = global.flgOn;
            }
            else
            {
                s.特殊出勤 = global.flgOff;
            }

            if (chkTooshi.Checked)
            {
                s.通し勤務 = global.flgOn;
            }
            else
            {
                s.通し勤務 = global.flgOff;
            }

            if (chkYakan.Checked)
            {
                s.夜間手当 = global.flgOn;
            }
            else
            {
                s.夜間手当 = global.flgOff;
            }

            if (chkShokumu.Checked)
            {
                s.職務手当 = global.flgOn;
            }
            else
            {
                s.職務手当 = global.flgOff;
            }

            s.通勤業務走行 = Utility.StrtoInt(txtKmTuukin.Text);
            s.私用走行 = Utility.StrtoInt(txtKmShiyou.Text);

            if (dtDDay.Checked)
            {
                s.代休対象日 = dtDDay.Value.Year.ToString() + dtDDay.Value.Month.ToString().PadLeft(2, '0') + dtDDay.Value.Day.ToString().PadLeft(2, '0');
            }
            else
            {
                s.代休対象日 = string.Empty;
            }

            if (chkManager.Checked)
            {
                s.確認印 = global.flgOn;
            }
            else
            {
                s.確認印 = global.flgOff;
            }

            if (fMode.Mode == global.FORM_ADDMODE)
            {
                s.登録年月日 = DateTime.Now;
                s.登録ユーザーID = global.loginUserID;
            }

            s.更新年月日 = DateTime.Now;
            s.更新ユーザーID = global.loginUserID;
        }

        private void txtInM_Leave(object sender, EventArgs e)
        {
            TextBox tb = (TextBox)sender;

            tb.Text = tb.Text.PadLeft(2, '0');
        }
        
        private void lblNum_TextChanged(object sender, EventArgs e)
        {
            if (dts.M_社員.Any(a => a.ID == Utility.StrtoInt(lblNum.Text)))
            {
                var s = dts.M_社員.Single(a => a.ID == Utility.StrtoInt(lblNum.Text));
                lblName.Text = s.氏名;
                lblName.ForeColor = Color.Black;
            }
            else
            {
                lblName.Text = "未登録";
                lblName.ForeColor = Color.Red;
            }
        }

        private void txtSiM_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtSiH_TextChanged(object sender, EventArgs e)
        {

        }

        private void label18_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void label17_Click(object sender, EventArgs e)
        {

        }
    }
}
