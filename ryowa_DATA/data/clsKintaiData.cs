using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ryowa_DATA.common;
using System.Windows.Forms;

namespace ryowa_DATA.data
{
    class clsKintaiData
    {
        ryowaDataSet dts = new ryowaDataSet();
        ryowaDataSetTableAdapters.T_勤怠TableAdapter adp = new ryowaDataSetTableAdapters.T_勤怠TableAdapter();
        ryowaDataSetTableAdapters.M_社員TableAdapter sAdp = new ryowaDataSetTableAdapters.M_社員TableAdapter();
        ryowaDataSetTableAdapters.M_工事TableAdapter pAdp = new ryowaDataSetTableAdapters.M_工事TableAdapter();
        ryowaDataSetTableAdapters.M_休日TableAdapter hAdp = new ryowaDataSetTableAdapters.M_休日TableAdapter();
        ryowaDataSetTableAdapters.環境設定TableAdapter cAdp = new ryowaDataSetTableAdapters.環境設定TableAdapter();

        int _sYY = 0;
        int _sMM = 0;

        // 特殊勤務
        int tJyo = 0;
        int tToku1 = 0;
        int tToku2 = 0;
        int tTooshi = 0;
        int tYakan = 0;
        int tShokumu = 0;

        // 給与大臣用CSVデータ作成パス
        string csvPath = string.Empty;

        public workData[] wd = null;

        public clsKintaiData(int sYY, int sMM)
        {
            // データ読み込み : 対象月のみ取得 2018/07/20
            //adp.Fill(dts.T_勤怠);
            adp.FillByYYMM(dts.T_勤怠, sYY, sMM);

            sAdp.Fill(dts.M_社員);
            pAdp.Fill(dts.M_工事);
            hAdp.Fill(dts.M_休日);
            cAdp.Fill(dts.環境設定);

            _sYY = sYY; // 対象年
            _sMM = sMM; // 対象月

            // 環境設定ファイルより各種手当単価を読み込み
            var d = dts.環境設定.Single(a => a.ID == global.configKEY);
            tJyo = d.除雪当番単価;
            tToku1 = d.特殊勤務単価1;
            tToku2 = d.特殊勤務単価2;
            tYakan = d.夜間手当単価;
            tShokumu = d.職務手当単価;
            csvPath = d.受け渡しデータ作成パス;

            // 勤怠データ集計
            getCheckData();
        }

        ///-------------------------------------------------------------------
        /// <summary>
        ///     勤怠データ取得 </summary>
        ///-------------------------------------------------------------------
        public void getCheckData()
        {
            var s = dts.T_勤怠
                .Where(a => a.日付.Year == _sYY && a.日付.Month == _sMM)
                .GroupBy(a => a.社員ID)
                .Select(gr => new
                {
                    sID = gr.Key,
                    sCnt = gr.Count(),
                    sZan = gr.Sum(a => a.普通残業 + a.早出残業),    // 2018/07/20 早出残業を含める
                    sSinya = gr.Sum(a => a.深夜残業)
                })
                .OrderBy(a => a.sID);

            // 対象データなしのとき
            if (s.Count() == 0)
            {
                return;
            }
            else
            {
                // 勤怠データ配列生成
                wd = new workData[s.Count()];
            }

            int iX = 0;

            // 社員ごとに読み込む
            foreach (var t in s)
            {
                //int hol = 0;        // 休日勤務時間
                //int houtei = 0;     // 法定休日勤務時間

                int holD = 0;       // 休日代休時間
                int houteiD = 0;    // 法定休日代休時間

                int shaink10 = 0;
                int shainTooshi = 0;
                string sName = string.Empty;

                int sJyo = 0;
                int sTooshi = 0; 
                int sYakan = 0; 
                decimal sShokushu = 0;
                int sZanStatus = 0; // 残業有無
                int sGenbaKbn = 0;  // 現場手当有無 2018/09/11
                double fZan = 0;    // 固定残業時間 2018/09/21

                // 社員情報を取得
                if (dts.M_社員.Any(a => a.ID == t.sID))
                {
                    var ss = dts.M_社員.Single(a => a.ID == t.sID);
                    sName = ss.氏名;
                    shaink10 = ss.基本給10;
                    shainTooshi = ss.通し勤務単価;
                    sZanStatus = ss.残業有無;
                    sGenbaKbn = ss.現場手当有無;  // 2018/09/11
                    fZan = ss.固定残業時間;       // 2018/09/21
                }
                else
                {
                    // 社員マスター該当なしのとき
                    sName = string.Empty;
                    shaink10 = 0;
                    shainTooshi = 0;
                    sGenbaKbn = 0;  // 2018/09/11
                    fZan = 0;       // 2018/09/21
                }
                
                wd[iX] = new workData();    // 配列インスタンスの生成

                wd[iX].code = t.sID;    // 個人コード
                wd[iX].name = sName;    // 氏名

                // 残業有無ステータス
                if (sZanStatus == global.flgOn)
                {
                    // 工事毎の残業時間 2018/09/21
                    var ss = dts.T_勤怠
                        .Where(a => a.社員ID == t.sID && a.日付.Year == _sYY && a.日付.Month == _sMM)
                        .GroupBy(a => a.工事ID)
                        .Select(gr => new
                        {
                            pID = gr.Key,
                            pHayade = gr.Sum(a => a.早出残業),
                            pZan = gr.Sum(a => a.普通残業),
                            pSinya = gr.Sum(a => a.深夜残業)
                        }).OrderBy(a => a.pID);     // 工事IDでソード 2018/09/25
                    
                    double tlZan = 0;   // 固定残業時間と比較する累積残業時間 2018/09/17
                    fZan *= 60;         // 固定残業時間を分単位に変換 2018/09/20

                    // 残業内訳クラス配列
                    clsZangyoArray[] zArray = new clsZangyoArray[ss.Count()];

                    int ij = 0;

                    // 以下、2018/09/21
                    foreach (var kk in ss)
                    {
                        // 残業内訳クラス配列初期化
                        zArray[ij] = new clsZangyoArray();
                        zArray[ij].sHol = 0;
                        zArray[ij].sHoutei = 0;
                        zArray[ij].sZangyo = 0;
                        zArray[ij].sShinya = 0;

                        // 工事部署ごとの休日勤務時間・法定休日勤務時間を求める
                        int _hol = 0;
                        int _houtei = 0;
                        Utility.getHolTime(dts.T_勤怠, out _hol, out _houtei, kk.pID, _sYY, _sMM, t.sID);

                        // 工事部署ごとの代休時間を求める
                        int _holD = 0;
                        int _houteiD = 0;
                        Utility.getdaikyuTime(dts, out _holD, out _houteiD, _sYY, _sMM, t.sID, kk.pID);
                        
                        // 代休取得した時間を差し引く
                        _hol -= _holD;
                        _houtei -= _houteiD;

                        /* ------------------------------------------------------------------
                         * 以下、固定残業時間を超過した時間を残業時間を計算する
                         * ------------------------------------------------------------------
                         */
                        zArray[ij].sHol = (int)fix2Zan(_hol, ref tlZan, fZan);    // 休日勤務時間
                        zArray[ij].sHoutei = (int)fix2Zan(_houtei, ref tlZan, fZan);     // 法定休日勤務時間
                        zArray[ij].sZangyo = (int)fix2Zan(kk.pHayade + kk.pZan, ref tlZan, fZan);   // 早出 + 普通残業
                        zArray[ij].sShinya = (int)fix2Zan(kk.pSinya, ref tlZan, fZan);    // 深夜残業

                        //--------------------------------------------------------------------

                        // 休日代休時間
                        holD += _holD;

                        // 法定休日代休時間
                        houteiD += _houteiD;

                        ij++;
                    }


                    //// 残業支給対象 2018/09/21 コメント化
                    //wd[iX].zanTM = t.sZan;  // 普通残業（早出残業含む 2018/07/20）
                    //wd[iX].siTM = t.sSinya; // 深夜残業

                    // 残業支給対象
                    int tHol = 0;
                    int tHoutei = 0;
                    int tZan = 0;
                    int tShinya = 0;

                    for (int i = 0; i < zArray.Length; i++)
                    {
                        tHol += zArray[i].sHol;
                        tHoutei += zArray[i].sHoutei;
                        tZan += zArray[i].sZangyo;
                        tShinya += zArray[i].sShinya;
                    }

                    wd[iX].zanTM = tZan;    // 普通残業（早出残業含む 2018/09/21）
                    wd[iX].siTM = tShinya;  // 深夜残業 2018/09/21

                    //// 月間の休日勤務時間、法定休日勤務時間取得 2018/09/21 コメント化
                    //Utility.getHolTime(dts.T_勤怠, out hol, out houtei, _sYY, _sMM, t.sID);

                    //// 休日代休時間、法定休日代休時間取得 2018/09/21 コメント化
                    //getdaikyuTime(out holD, out houteiD, _sYY, _sMM, t.sID);
                    //hol -= holD;       // 休日勤務時間より休日代休時間を差し引く
                    //houtei -= houteiD; // 法定休日勤務時間より法定休日代休時間を差し引く

                    //wd[iX].holTM = hol;                 // 休日勤務時間 2018/09/21 コメント化
                    //wd[iX].houteiTM = houtei;           // 法定休日勤務時間 2018/09/21 コメント化

                    wd[iX].holTM = tHol;         // 休日勤務時間 2018/09/21
                    wd[iX].houteiTM = tHoutei;   // 法定休日勤務時間 2018/09/21

                    wd[iX].holDaikyuTM = holD;          // 休日代休時間
                    wd[iX].houteiDaikyuTM = houteiD;    // 法定休日勤務時間
                }
                else
                {
                    // 残業支給対象外
                    wd[iX].zanTM = global.flgOff;           // 普通残業
                    wd[iX].siTM = global.flgOff;            // 深夜残業
                    wd[iX].holTM = global.flgOff;           // 休日勤務時間
                    wd[iX].houteiTM = global.flgOff;        // 法定休日勤務時間
                    wd[iX].holDaikyuTM = global.flgOff;     // 休日代休時間
                    wd[iX].houteiDaikyuTM = global.flgOff;  // 法定休日勤務時間
                }

                wd[iX].kinmuchiTe = getKinmuchiT(t.sID, shaink10);  // 勤務地手当
                wd[iX].sonota2 = getSonotaT(t.sID, shainTooshi, out sJyo, out sTooshi, out sYakan, out sShokushu);    // その他支給2

                // 現場回数
                // 社員マスターの「現場手当有無」が「有り」を対象とする 2018/09/11
                if (sGenbaKbn == global.flgOn)
                {
                    wd[iX].genbaNum = dts.T_勤怠.Count(a => a.社員ID == t.sID && a.日付.Year == _sYY &&
                                            a.日付.Month == _sMM && a.M_工事Row.現場区分 == global.flgOff &&
                                            (a.休日出勤 == global.flgOn || a.出勤印 == global.flgOn));
                } 
                else
                {
                    wd[iX].genbaNum = global.flgOff;
                }

                // 県内宿泊回数
                wd[iX].inStayNum = dts.T_勤怠.Count(a => a.社員ID == t.sID && a.日付.Year == _sYY &&
                                    a.日付.Month == _sMM && a.宿泊 == global.flgOn &&
                                    a.M_工事Row.勤務地区分 == global.flgOff);

                // 県外宿泊回数
                wd[iX].outStayNum = dts.T_勤怠.Count(a => a.社員ID == t.sID && a.日付.Year == _sYY &&
                                    a.日付.Month == _sMM && a.宿泊 == global.flgOn &&
                                    a.M_工事Row.勤務地区分 == global.flgOn);

                // 遠隔地回数
                wd[iX].enkakuchiNum = dts.T_勤怠.Count(a => a.社員ID == t.sID && a.日付.Year == _sYY &&
                                    a.日付.Month == _sMM && a.M_工事Row.勤務地区分 == 2 &&
                                   (a.休日出勤 == global.flgOn || a.出勤印 == global.flgOn));

                wd[iX].jyosekiNum = sJyo;   // 除雪当番数

                // 特殊出勤数
                wd[iX].tokushuNum = dts.T_勤怠.Count(a => a.社員ID == t.sID && a.日付.Year == _sYY &&
                                    a.日付.Month == _sMM && a.特殊出勤 == global.flgOn);

                wd[iX].tooshiNum = sTooshi;     // 通し勤務数
                wd[iX].yakanNum = sYakan;       // 夜間手当数
                wd[iX].shokumuNum = sShokushu;  // 職務手当日数

                iX++;
            }
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     固定残業時間を差し引いた時間外勤務を計算 </summary>
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

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     対象年月の休日代休時間、法定休日代休時間取得 </summary>
        /// <param name="holDTM">
        ///     休日代休時間</param>
        /// <param name="houteiDTM">
        ///     法定休日代休時間</param>
        /// <param name="sYY">
        ///     対象年</param>
        /// <param name="sMM">
        ///     対象月</param>
        /// <param name="sNum">
        ///     個人コード</param>
        ///----------------------------------------------------------------------------
        private void getdaikyuTime(out int holDTM, out int houteiDTM, int sYY, int sMM, int sNum)
        {
            holDTM = 0;
            houteiDTM = 0;

            var s = dts.T_勤怠.Where(a => a.日付.Year == sYY && a.日付.Month == sMM &&
                                          a.社員ID == sNum && a.代休 == global.flgOn);

            if (s.Count() == 0)
            {
                return;
            }

            foreach (var t in s)
            {
                DateTime dt;

                if (DateTime.TryParse(t.代休対象日.Substring(0, 4) + "/" + t.代休対象日.Substring(4, 2) + "/" + t.代休対象日.Substring(6, 2), out dt))
                {
                    if (dts.M_休日.Any(a => a.日付 == dt))
                    {
                        // 休日勤務時間を取得
                        int wTM = getHolWorkTime(dt, sNum);

                        if (wTM > Properties.Settings.Default.workTime)
                        {
                            wTM = Properties.Settings.Default.workTime;
                        }

                        var d = dts.M_休日.Single(a => a.日付 == dt);
                        if (d.法定休日 == global.flgOn)
                        {
                            // 法定休日代休のとき
                            houteiDTM += wTM;
                        }
                        else
                        {
                            // 休日代休のとき
                            holDTM += wTM;
                        }
                    }
                }
            }
        }


        ///------------------------------------------------------------------
        /// <summary>
        ///     勤務時間取得 </summary>
        /// <param name="dt">
        ///     日付</param>
        /// <param name="sNum">
        ///     個人コード</param>
        /// <returns>
        ///     休日勤務時間・分</returns>
        ///------------------------------------------------------------------
        private int getHolWorkTime(DateTime dt, int sNum)
        {
            if (!dts.T_勤怠.Any(a => a.日付 == dt && a.社員ID == sNum))
            {
                return 0;
            }

            var s = dts.T_勤怠.Single(a => a.日付 == dt && a.社員ID == sNum);

            // 勤務時間を取得する : 2019/01/04
            double spanMin = Utility.getWorkTime(s, false);
            return (int)spanMin;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     その他支給2取得 </summary>
        /// <param name="sNum">
        ///     個人コード</param>
        /// <param name="tooshi">
        ///     通し勤務単価</param>
        /// <param name="sJyo">
        ///     除雪当番回数</param>
        /// <param name="sTooshi">
        ///     通し勤務回数</param>
        /// <param name="sYakan">
        ///     夜間手当回数</param>
        /// <param name="sShokumu">
        ///     職務手当回数</param>
        /// <returns>
        ///     その他支給2金額</returns>
        ///----------------------------------------------------------------------
        private int getSonotaT(int sNum, int tooshi, out int sJyo, out int sTooshi, out int sYakan, out decimal sShokumu)
        {
            int sonotaT = 0;

            // 除雪手当
            sJyo = dts.T_勤怠.Count(a => a.日付.Year == _sYY && a.日付.Month == _sMM && a.社員ID == sNum &&
                        a.除雪当番 == global.flgOn);

            sonotaT += sJyo * tJyo;

            // 特殊勤務手当
            sonotaT += getTokushuT(sNum);

            // 通し勤務手当
            sTooshi = dts.T_勤怠.Count(a => a.日付.Year == _sYY && a.日付.Month == _sMM && a.社員ID == sNum &&
                        a.通し勤務 == global.flgOn);

            sonotaT += sTooshi * tooshi;

            // 夜間手当
            sYakan = dts.T_勤怠.Count(a => a.日付.Year == _sYY && a.日付.Month == _sMM && a.社員ID == sNum &&
                        a.夜間手当 == global.flgOn);

            sonotaT += sYakan * tYakan;

            // 職務手当
            sShokumu = dts.T_勤怠.Count(a => a.日付.Year == _sYY && a.日付.Month == _sMM && a.社員ID == sNum &&
                        a.職務手当 == global.flgOn);

            // 職務手当日数は上限を21.6日とする 2019/01/09
            if (sShokumu > Properties.Settings.Default.tempdays)
            {
                sShokumu = Properties.Settings.Default.tempdays;
            }

            // 職務手当は最大15,000円／月
            if ((sShokumu * tShokumu) > 15000)
            {
                sonotaT += 15000;
            }
            else
            {
                sonotaT += (int)(sShokumu * tShokumu);  // 2019/01/07
            }

            return sonotaT;
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     特殊出勤手当取得 </summary>
        /// <param name="sNum">
        ///     個人コード</param>
        /// <returns>
        ///     特殊出勤手当</returns>
        ///----------------------------------------------------------------------
        private int getTokushuT(int sNum)
        {
            int tt = 0;

            var s = dts.T_勤怠.Where(a => a.日付.Year == _sYY && a.日付.Month == _sMM && a.社員ID == sNum &&
                        a.特殊出勤 == global.flgOn && (a.出勤印 == global.flgOn || a.休日出勤 == global.flgOn));

            foreach (var t in s)
            {
                // 勤務時間を取得する : 2019/01/04
                if (Utility.getWorkTime(t, true) < 240)
                {
                    tt += tToku1;   // 4.0時間／日 未満
                }
                else
                {
                    tt += tToku2;   // 4.0時間／日 以上
                }
            }

            return tt;
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     勤務地手当取得 </summary>
        /// <param name="sNum">
        ///     個人コード</param>
        /// <param name="k10">
        ///     個人コード対象社員の基本給の10％</param>
        /// <returns>
        ///     勤務地手当</returns>
        ///-------------------------------------------------------------------------
        private int getKinmuchiT(int sNum, int k10)
        {
            if (k10 == global.flgOff)
            {
                return 0;
            }
            else
            {
                return getEnkakuchiT(sNum, k10) + getKengaiT(sNum, k10);
            }
        }

        ///--------------------------------------------------------------------------
        /// <summary>
        ///     遠隔地手当取得 </summary>
        /// <param name="sNum">
        ///     個人コード</param>
        /// <param name="k10">
        ///     個人コード対象社員の基本給の10％</param>
        /// <returns>
        ///     遠隔地手当</returns>
        ///--------------------------------------------------------------------------
        private int getEnkakuchiT(int sNum, int k10)
        {
            // 遠隔地手当
            int en = 0;
            decimal n = 0;

            int s = dts.T_勤怠.Count(a => a.M_工事Row.勤務地区分 == 2 && a.社員ID == sNum && 
                        a.日付.Year == _sYY && a.日付.Month == _sMM && 
                        (a.出勤印 == global.flgOn || a.休日出勤 == global.flgOn));
            
            if (s == 0)
            {
                // 遠隔地勤務なしのとき
                return en;
            }
            else if (s >= Properties.Settings.Default.tempdays)
            {
                n = Properties.Settings.Default.tempdays;
            }
            else
            {
                n = s;
            }

            en = (int)(k10 * n / Properties.Settings.Default.tempdays);

            return en;
        }

        ///---------------------------------------------------------------------------
        /// <summary>
        ///     県外手当取得 </summary>
        /// <param name="sNum">
        ///     個人コード</param>
        /// <param name="k10">
        ///     個人コード対象社員の基本給の10％</param>
        /// <returns>
        ///     県外手当</returns>
        ///---------------------------------------------------------------------------
        private int getKengaiT(int sNum, int k10)
        {
            int kn = 0;

            // 全ての出勤日を抽出
            var s = dts.T_勤怠.Where(a => a.社員ID == sNum && a.日付.Year == _sYY && a.日付.Month == _sMM && 
                        (a.出勤印 == global.flgOn || a.休日出勤 == global.flgOn));

            if (s.Count() == global.flgOff)
            {
                // 出勤なし
                return kn;
            }

            // 出勤数を取得
            int wCnt = s.Count(); 

            // 県外をカウント
            decimal sc = s.Count(a => a.M_工事Row.勤務地区分 == global.flgOn);

            // 全ての勤務日が県外か？
            if (sc == wCnt)
            {
                // 全てが県外のとき基本給の10％
                kn = k10;
            }
            else
            {
                // 最高で25日
                // 最高で21.6日 2018/07/10
                if (sc >= Properties.Settings.Default.tempdays)
                {
                    sc = Properties.Settings.Default.tempdays;
                }

                kn = (int)(k10 * sc / Properties.Settings.Default.tempdays);
            }

            return kn;
        }

        ///---------------------------------------------------------------
        /// <summary>
        ///     給与大臣用CSVデータ作成 </summary>
        ///---------------------------------------------------------------
        public void csvOutput()
        {
            if (wd == null)
            {
                return;
            }

            string[] cd = new string[wd.Length];
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < wd.Length; i++)
            {
                sb.Clear();

                sb.Append(wd[i].code.ToString()).Append(",");
                sb.Append(Utility.intToHhMM(wd[i].zanTM)).Append(",");
                sb.Append(Utility.intToHhMM(wd[i].siTM)).Append(",");
                sb.Append(Utility.intToHhMM(wd[i].holTM)).Append(",");
                sb.Append(Utility.intToHhMM(wd[i].houteiTM)).Append(",");
                sb.Append(Utility.intToHhMM(wd[i].holDaikyuTM)).Append(",");
                sb.Append(Utility.intToHhMM(wd[i].houteiDaikyuTM)).Append(",");
                sb.Append(wd[i].kinmuchiTe.ToString()).Append(",");
                sb.Append(wd[i].sonota2.ToString()).Append(",");
                sb.Append(wd[i].genbaNum.ToString()).Append(",");
                sb.Append(wd[i].inStayNum.ToString()).Append(",");
                sb.Append(wd[i].outStayNum.ToString()).Append(",");
                sb.Append("0,");
                sb.Append("0,");
                sb.Append(wd[i].enkakuchiNum.ToString()).Append(",");

                cd[i] = sb.ToString();
            }

            // 受け渡しデータファイル名
            string fName = _sYY.ToString() + "年" + _sMM + "月" + Properties.Settings.Default.csvFileName;

            // 受け渡しデータパス出力
            Utility.csvFileWrite(csvPath, cd, fName);

            // 終了
            MessageBox.Show("終了しました");
        }


        ///-----------------------------------------------------------------------
        /// <summary>
        ///     勤怠チェックデータクラス </summary>
        ///-----------------------------------------------------------------------
        public class workData
        {
            public int code { get; set; }           // 個人コード
            public string name { get; set; }        // 氏名
            public int zanTM { get; set; }          // 残業時間（早出残業含む 2018/07/20）
            public int siTM { get; set; }           // 深夜残業
            public int holTM { get; set; }          // 休日勤務時間
            public int houteiTM { get; set; }       // 法定休日勤務時間
            public int holDaikyuTM { get; set; }    // 休日代休時間
            public int houteiDaikyuTM { get; set; } // 法定休日代休時間
            public int kinmuchiTe { get; set; }     // 勤務地手当
            public int sonota2 { get; set; }        // その他支給２
            public int genbaNum { get; set; }       // 現場回数
            public int inStayNum { get; set; }      // 県内宿泊回数
            public int outStayNum { get; set; }     // 県外宿泊回数
            public int enkakuchiNum { get; set; }   // 遠隔地回数
            public int jyosekiNum { get; set; }     // 除籍手当
            public int tokushuNum { get; set; }     // 特殊出勤日数
            public int tooshiNum { get; set; }      // 通し勤務日数
            public int yakanNum { get; set; }       // 夜間手当日数
            public decimal shokumuNum { get; set; }     // 職務手当日数
        }
    }
}
