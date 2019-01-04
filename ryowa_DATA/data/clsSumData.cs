using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ryowa_DATA.data
{
    class clsSumData
    {
        public sumData[] sArray = new sumData[3];

        public clsSumData()
        {
            sumDatainitial(sArray);
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     集計データ配列初期化 </summary>
        /// <param name="s">
        ///     集計配列</param>
        ///----------------------------------------------------------------------
        public void sumDatainitial(sumData[] s)
        {
            for (int i = 0; i < 3; i++)
            {
                s[i] = new sumData();

                s[i].sID = "";
                s[i].sName = "";
                s[i].pID = "";
                s[i].pName = "";
                s[i].sJinkanhi = 0;
                s[i].sHaichiDays = 0;
                s[i].sGanbaDays = 0;
                s[i].sKinmuchiDays = 0;
                s[i].sStayDays = 0;
                s[i].sHolTM = 0;
                s[i].sHouteiTM = 0;
                s[i].sZanTM = 0;
                s[i].sSiTM = 0;
                s[i].sJyosetsu = 0;
                s[i].sTokushu = 0;
                s[i].sTooshi = 0;
                s[i].sYakan = 0;
                s[i].sShokumu = 0;
            }
        }


        ///------------------------------------------------------------------------
        /// <summary>
        ///     集計データ配列初期化（工事別、社員別）</summary>
        /// <param name="s">
        ///     集計データ配列</param>
        ///------------------------------------------------------------------------
        public void sumDataCrear(sumData[] s)
        {
            for (int i = 0; i < 2; i++)
            {
                s[i].sID = "";
                s[i].sName = "";
                s[i].pID = "";
                s[i].pName = "";
                s[i].sJinkanhi = 0;
                s[i].sHaichiDays = 0;
                s[i].sGanbaDays = 0;
                s[i].sKinmuchiDays = 0;
                s[i].sStayDays = 0;
                s[i].sHolTM = 0;
                s[i].sHouteiTM = 0;
                s[i].sZanTM = 0;
                s[i].sSiTM = 0;
                s[i].sJyosetsu = 0;
                s[i].sTokushu = 0;
                s[i].sTooshi = 0;
                s[i].sYakan = 0;
                s[i].sShokumu = 0;
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     集計データ配列クラス </summary>
        ///----------------------------------------------------------------------
        public class sumData
        {
            public string sID { get; set; }
            public string sName { get; set; }
            public string pID { get; set; }
            public string pName { get; set; }
            public int sJinkanhi { get; set; }
            public decimal sHaichiDays { get; set; }
            public int sGanbaDays { get; set; }
            public decimal sKinmuchiDays { get; set; }
            public int sStayDays { get; set; }
            public int sHolTM { get; set; }
            public int sHouteiTM { get; set; }
            public int sZanTM { get; set; }
            public int sSiTM { get; set; }
            public int sJyosetsu { get; set; }
            public int sTokushu { get; set; }
            public int sTooshi { get; set; }
            public int sYakan { get; set; }
            public int sShokumu { get; set; }
        }
    }
}
