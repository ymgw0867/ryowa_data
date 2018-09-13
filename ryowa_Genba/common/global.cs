using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ryowa_Genba.common
{
    class global
    {
        public static string pblImagePath;
        
        //和暦西暦変換
        public const int rekiCnv = 1988;    //西暦、和暦変換
        
        #region ローカルMDB関連定数
        public const string MDBFILE = "ryowa.mdb";         // MDBファイル名
        public const string MDBTEMP = "ryowa_Temp.mdb";    // 最適化一時ファイル名
        public const string MDBBACK = "ryowa_Back.mdb";    // 最適化後バックアップファイル名
        #endregion

        #region フラグオン・オフ定数
        public const int flgOn = 1;            //フラグ有り(1)
        public const int flgOff = 0;           //フラグなし(0)
        public const string FLGON = "1";
        public const string FLGOFF = "0";
        public const string chkBoxTrue = "true";    // データグリッドビューのチェックボックスのTrue値
        public const string chkBoxFalse = "false";  // データグリッドビューのチェックボックスのFalse値
        #endregion

        public static int pblDenNum;            // データ数

        public const int configKEY = 1;        // 環境設定データキー
        public const int mailKey = 1;          // メール設定データキー

        //ＯＣＲ処理ＣＳＶデータの検証要素
        public const int CSVLENGTH = 197;          //データフィールド数 2011/06/11
        public const int CSVFILENAMELENGTH = 21;   //ファイル名の文字数 2011/06/11  
 
        // 勤務記録表
        public const int STARTTIME = 8;            // 単位記入開始時間帯
        public const int ENDTIME = 22;             // 単位記入終了時間帯
        public const int TANNIMAX = 4;             // 単位最大値
        public const int WEEKLIMIT = 160;          // 週労働時間基準単位：40時間
        public const int DAYLIMIT = 32;            // 一日あたり労働時間基準単位：8時間

        #region 環境設定項目
        public static int cnfYear;                  // 対象年
        public static int cnfMonth;                 // 対象月
        public static string cnfPath;               // 受け渡しデータ作成パス
        public static int cnfArchived;              // データ保管期間（月数）
        #endregion

        #region コード桁数定数
        public const int ShozokuLength = 0;                 // 所属コード桁数
        public const int ShainLength = 0;                   // 社員コード桁数
        public const int ShozokuMaxLength = 4;              // 所属コードＭＡＸ桁数
        public const int ShainMaxLength = 4;                // 社員コードＭＡＸ桁数
        #endregion  
        
        // 深夜時間帯チェック用
        public static DateTime dt2200 = DateTime.Parse("22:00");
        public static DateTime dt0000 = DateTime.Parse("0:00");
        public static DateTime dt0500 = DateTime.Parse("05:00");
        public static DateTime dt0800 = DateTime.Parse("08:00");
        public static DateTime dt2359 = DateTime.Parse("23:59");
        public static DateTime dt1700 = DateTime.Parse("17:00");
        public const int TOUJITSU_SINYATIME = 120;      // 終了時刻が翌日のときの当日の深夜勤務時間

        // ChangeValueStatus
        public static bool ChangeValueStatus = true;

        public const int MAX_GYO = 31;
        public const int MAX_MIN = 1;
        
        #region 勤務管理表種別ID定数
        public const string SHAIN_ID = "1";
        public const string PART_ID = "2";
        public const string SHUKKOU_ID = "3";
        #endregion

        public string[] arrayChohyoID = { "社員","パート","出向社員" };

        // データ作成画面datagridview表示行数
        public const int _MULTIGYO = 31;

        // フォーム登録モード
        public const int FORM_ADDMODE = 0;
        public const int FORM_EDITMODE = 1;

        // 社員マスター検索該当者なしの戻り値
        public const string NO_MASTER = "NonMaster";
        public const string NO_ZAISEKI = "NonZaiseki";
        public const string NO_TAISHOKU = "NonTaishoku";
        public const string NO_KYUSHOKU = "NonKyushoku";
        
        // 現場区分
        public string[,] arrStyle = new string[3, 2] { { "0", "現場" }, { "1", "講習会" }, { "2", "その他" }};

        // 勤務地区分
        public string[,] kArrStyle = new string[3, 2] { { "0", "県内" }, { "1", "県外" }, { "2", "遠隔地" } };

        // 残業有無区分
        public string[,] zArrStyle = new string[2, 2] { { "0", "なし" }, { "1", "あり" } };

        // 勤怠編集権限区分
        public string[,] eArrStyle = new string[2, 2] { { "0", "自らの勤怠のみ" }, { "1", "他者の勤怠編集可能" } };

        // システムユーザー区分
        public string[,] uArrStyle = new string[2, 2] { { "0", "スタッフメニュー" }, { "1", "スタッフ＋本部メニュー" } };

        // 年月日未設定値
        public static DateTime NODATE = DateTime.Parse("1900/01/01");
        
        // ログインステータス
        public static bool loginStatus;

        // ログインユーザー情報
        public static int loginUserID = 0;
        public static int loginType = 0;
        public static int loginSysUser = 0;

        // 祝日配列
        public static string[,] wHoriDay = new string[10, 2] { {"01/01", "元旦" }, {"02/11","建国記念の日"}, { "04/29", "昭和の日" }, { "05/03", "憲法記念日" }, { "05/04", "みどりの日" }, { "05/05", "こどもの日" }, 
            { "08/11", "山の日" }, { "11/03", "文化の日" }, { "11/23", "勤労感謝の日" }, { "12/23", "天皇誕生日" }};

    }
}
