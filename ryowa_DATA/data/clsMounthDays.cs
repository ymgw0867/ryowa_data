using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ryowa_DATA.common;

namespace ryowa_DATA.data
{
    class clsMounthDays
    {
        ryowaDataSet dts = new ryowaDataSet();
        ryowaDataSetTableAdapters.T_勤怠TableAdapter adp = new ryowaDataSetTableAdapters.T_勤怠TableAdapter();

        int _sYY = 0;
        int _sMM = 0;

        public mDays[] md = null;       // 個人別月間出勤日数
        public haichiDays[] hd = null;  // 個人別工事別配置日数

        ///-----------------------------------------------------------
        /// <summary>
        ///     個人別月間出勤日数配列作成 </summary>
        ///-----------------------------------------------------------
        private void getSumMonthDays()
        {
            // 社員ID毎の月間出勤日数取得
            var dc = dts.T_勤怠.Where(a => (a.日付.Year == _sYY && a.日付.Month == _sMM) &&
                                           (a.出勤印 == global.flgOn || a.休日出勤 == global.flgOn))
                .GroupBy(a => a.社員ID)
                .Select(n => new
                {
                    sID = n.Key,
                    sDayCnt = n.Count()
                });

            // 月間出勤日数配列
            md = new mDays[dc.Count()];

            int iX = 0;

            // 配列に格納
            foreach (var item in dc)
            {
                md[iX] = new mDays();
                md[iX].sNum = item.sID;
                md[iX].sDays = item.sDayCnt;
                iX++;
            }
        }

        public clsMounthDays(int sYY, int sMM)
        {
            _sYY = sYY;
            _sMM = sMM;

            adp.Fill(dts.T_勤怠);

            getSumMonthDays();  // 月間出勤日数取得
            setHaichiDays();    // 配置日数配列作成
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     配置日数配列作成 </summary>
        ///-------------------------------------------------------------
        private void setHaichiDays()
        {
            hd = null;
            decimal hDaysTl = 0;
            int iX = 0;
            decimal monthWorkDays = 0;

            // 社員ID、工事ID毎の出勤日集計
            var s = dts.T_勤怠.Where(a => (a.日付.Year == _sYY && a.日付.Month == _sMM) &&
                                          (a.出勤印 == global.flgOn || a.休日出勤 == global.flgOn))
                .GroupBy(a => a.社員ID)
                .Select(n => new
                {
                    sID = n.Key,
                    pID = n.GroupBy(a => a.工事ID)
                    .Select(k => new
                    {
                        kID = k.Key,
                        cnt = k.Count()
                    })
                    .OrderByDescending(a => a.cnt)
                })
                .OrderBy(a => a.sID);

            foreach (var t in s)
            {
                // 月間出勤日数取得
                for (int i = 0; i < md.Length; i++)
                {
                    if (md[i].sNum == t.sID)
                    {
                        monthWorkDays = md[i].sDays;
                        break;
                    }
                }

                // 個人別配置日数計初期化
                hDaysTl = 0;

                foreach (var j in t.pID)
                {
                    // 配置日数（小数点以下第一位四捨五入） 2018/07/10
                    int hh = (int)((j.cnt * Properties.Settings.Default.tempdays / monthWorkDays * 100 + 5) / 10);
                    decimal h = (decimal)hh / 10;
                    if ((hDaysTl + h) > Properties.Settings.Default.tempdays)
                    {
                        h = (int)(Properties.Settings.Default.tempdays - hDaysTl);
                    }
                    else
                    {
                        hDaysTl += h;
                    }

                    // 配列作成
                    Array.Resize(ref hd, iX + 1);
                    hd[iX] = new haichiDays();
                    hd[iX].sNum = t.sID;
                    hd[iX].kID = j.kID;
                    hd[iX].sDays = h;

                    iX++;
                }
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     個人別工事別配置日数取得 </summary>
        /// <param name="sNum">
        ///     個人コード</param>
        /// <param name="kCode">
        ///     工事コード</param>
        /// <returns>
        ///     配置日数</returns>
        ///---------------------------------------------------------------------
        public decimal getHaichidays(int sNum, int kCode)
        {
            decimal rVal = 0;

            for (int i = 0; i < hd.Length; i++)
            {
                if (hd[i].sNum == sNum && hd[i].kID == kCode)
                {
                    rVal = hd[i].sDays;
                    break;
                }
            }
            return rVal;
        }
        ///-------------------------------------------------------------
        /// <summary>
        ///     月間出勤日数配列 </summary>
        ///-------------------------------------------------------------
        public class mDays
        {
            public int sNum { get; set; }   // 個人コード
            public int sDays { get; set; }  // 月間出勤日数
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     個人別工事別配置日数配列 </summary>
        ///-------------------------------------------------------------
        public class haichiDays
        {
            public int sNum { get; set; }   // 個人コード
            public int kID { get; set; }    // 工事コード
            public decimal sDays { get; set; }  // 配置日数
        }
    }
}
