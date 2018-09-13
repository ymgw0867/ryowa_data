using System;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Net.Mail;

namespace ryowa_Genba.common
{
    class Utility
    {
        /// ---------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     Height</param>
        /// ---------------------------------------------------------------------
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new Size(wSize, hSize);
        }

        /// ---------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     height</param>
        /// --------------------------------------------------------------------
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new Size(wSize, hSize);
        }

        /// <summary>
        /// フォームのデータ登録モード
        /// </summary>
        public class frmMode
        {
            public int Mode { get; set; }
            public string ID { get; set; }
            public int rowIndex { get; set; }
            public int closeMode { get; set; }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     工事コンボボックスクラス </summary>
        /// -------------------------------------------------------------------
        public class comboKouji
        {
            public string ID { get; set; }
            public string Name { get; set; }

            /// -------------------------------------------------------------------------
            /// <summary>
            ///     工事コンボボックスデータロード </summary>
            /// <param name="tempBox">
            ///     ロード先コンボボックスオブジェクト名</param>
            /// -------------------------------------------------------------------------
            public static void Load(ComboBox tempBox)
            {
                genbaDataSet dts = new genbaDataSet();
                genbaDataSetTableAdapters.M_工事TableAdapter adp = new genbaDataSetTableAdapters.M_工事TableAdapter();
                adp.Fill(dts.M_工事);

                try
                {
                    comboKouji cmb1;

                    tempBox.Items.Clear();
                    tempBox.DisplayMember = "Name";
                    tempBox.ValueMember = "ID";

                    foreach (var t in dts.M_工事.OrderBy(a => a.ID))
                    {
                        cmb1 = new comboKouji();
                        cmb1.ID = t.ID.ToString();
                        cmb1.Name = t.名称;
                        tempBox.Items.Add(cmb1);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "工事コンボボックスロード");
                }
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     文字列の値が数字かチェックする </summary>
        /// <param name="tempStr">
        ///     検証する文字列</param>
        /// <returns>
        ///     数字:true,数字でない:false</returns>
        /// ------------------------------------------------------------------------------
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }
        
        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     emptyを"0"に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// ------------------------------------------------------------------------------
        public static string EmptytoZero(string tempStr)
        {
            if (tempStr == string.Empty)
            {
                return "0";
            }
            else
            {
                return tempStr;
            }
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// -------------------------------------------------------------------------------
        public static string nulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// -------------------------------------------------------------------------------
        public static string nulltoStr2(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をIntへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Int型の値</returns>
        /// --------------------------------------------------------------------------------
        public static int StrtoInt(string tempStr)
        {
            if (NumericCheck(tempStr)) return int.Parse(tempStr);
            else return 0;
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をDoubleへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     double型の値</returns>
        /// --------------------------------------------------------------------------------
        public static double StrtoDouble(string tempStr)
        {
            if (NumericCheck(tempStr)) return double.Parse(tempStr);
            else return 0;
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をdecimalへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     decimal型の値</returns>
        /// --------------------------------------------------------------------------------
        public static decimal StrtoDecimal(string tempStr)
        {
            if (NumericCheck(tempStr))
            {
                return decimal.Parse(tempStr);
            }
            else
            {
                return 0;
            }
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     経過時間を返す </summary>
        /// <param name="s">
        ///     開始時間</param>
        /// <param name="e">
        ///     終了時間</param>
        /// <returns>
        ///     経過時間</returns>
        /// --------------------------------------------------------------------------------
        public static TimeSpan GetTimeSpan(DateTime s, DateTime e)
        {
            TimeSpan ts;
            if (s > e)
            {
                TimeSpan j = new TimeSpan(24, 0, 0);
                ts = e + j - s;
            }
            else
            {
                ts = e - s;
            }

            return ts;
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     指定した精度の数値に切り捨てます。</summary>
        /// <param name="dValue">
        ///     丸め対象の倍精度浮動小数点数。</param>
        /// <param name="iDigits">
        ///     戻り値の有効桁数の精度。</param>
        /// <returns>
        ///     iDigits に等しい精度の数値に切り捨てられた数値。</returns>
        /// ------------------------------------------------------------------------
        public static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = System.Math.Pow(10, iDigits);

            return dValue > 0 ? System.Math.Floor(dValue * dCoef) / dCoef :
                                System.Math.Ceiling(dValue * dCoef) / dCoef;
        }


        //部門コンボボックスクラス
        public class ComboBumon
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string NameShow { get; set; }
            public string code { get; set; }
            
            ////部門マスターロード
            //public static void load(ComboBox tempObj, int tempLen, string dbName)
            //{
            //    try
            //    {
            //        ComboBumon cmb1;
            //        string sqlSTRING = string.Empty;
            //        dbControl.DataControl sdcon = new dbControl.DataControl(dbName);
            //        OleDbDataReader dR;

            //        sqlSTRING += "select DepartmentID,DepartmentCode,DepartmentName from tbDepartment ";
            //        sqlSTRING += "where DepartmentID <> 1 ";
            //        sqlSTRING += "order by DepartmentCode ";

            //        //scom.CommandText = sqlSTRING;

            //        //データリーダーを取得する
            //        //dR = scom.ExecuteReader();
            //        dR = sdcon.FreeReader(sqlSTRING);

            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "Name";
            //        tempObj.ValueMember = "ID";

            //        while (dR.Read())
            //        {
            //            cmb1 = new ComboBumon();
            //            cmb1.ID = dR["DepartmentCode"].ToString();
            //            //cmb1.Name = string.Format("{0:000000000000000}", Int64.Parse(dR["DepartmentCode"].ToString())).Substring(15 - tempLen, tempLen) + " " + dR["DepartmentName"].ToString() + "";

            //            if (Utility.NumericCheck(dR["DepartmentCode"].ToString()))
            //                cmb1.Name = int.Parse(dR["DepartmentCode"].ToString()).ToString().PadLeft(tempLen, '0') + " " + dR["DepartmentName"].ToString() + "";
            //            else cmb1.Name = (dR["DepartmentCode"].ToString() + "    ").Substring(0, tempLen) + " " + dR["DepartmentName"].ToString() + "";

            //            cmb1.NameShow = dR["DepartmentName"].ToString() + "";
            //            tempObj.Items.Add(cmb1);
            //        }

            //        dR.Close();
            //        sdcon.Close();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "部門コンボボックスロード");
            //    }

            //}

            //部門コンボ表示
            public static void selectedIndex(ComboBox tempObj, int id)
            {
                ComboBumon cmbS = new ComboBumon();
                Boolean Sh;

                Sh = false;

                for (int iX = 0; iX <= tempObj.Items.Count - 1; iX++)
                {
                    tempObj.SelectedIndex = iX;
                    cmbS = (ComboBumon)tempObj.SelectedItem;

                    if (cmbS.ID == id.ToString())
                    {
                        Sh = true;
                        break;
                    }
                }

                if (Sh == false)
                {
                    tempObj.SelectedIndex = -1;
                }

            }
        }

        // 社員コンボボックスクラス
        public class ComboShain
        {
            public int ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }
            public int YakushokuType { get; set; }
            public string BumonName { get; set; }
            public string BumonCode { get; set; }

            //// 社員マスターロード
            //public static void load(ComboBox tempObj, string dbName)
            //{
            //    try
            //    {
            //        ComboShain cmb1;
            //        string sqlSTRING = string.Empty;
            //        dbControl.DataControl dCon = new dbControl.DataControl(dbName);
            //        OleDbDataReader dR;

            //        sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
            //        sqlSTRING += "where Shurojokyo = 1 ";
            //        sqlSTRING += "order by Code";

            //        //データリーダーを取得する
            //        dR = dCon.FreeReader(sqlSTRING);

            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";
            //        tempObj.ValueMember = "code";

            //        while (dR.Read())
            //        {
            //            cmb1 = new ComboShain();
            //            cmb1.ID = int.Parse(dR["Id"].ToString());
            //            cmb1.DisplayName = dR["Code"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.code = (dR["Code"].ToString() + "").Trim();
            //            cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
            //            tempObj.Items.Add(cmb1);
            //        }

            //        dR.Close();
            //        dCon.Close();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "社員コンボボックスロード");
            //    }

            //}

            // パートタイマーロード
            //public static void loadPart(ComboBox tempObj, string dbName)
            //{
            //    try
            //    {
            //        ComboShain cmb1;
            //        string sqlSTRING = string.Empty;
            //        dbControl.DataControl dCon = new dbControl.DataControl(dbName);
            //        OleDbDataReader dR;
            //        sqlSTRING += "select Bumon.Code as bumoncode,Bumon.Name as bumonname,Shain.Id as shainid,";
            //        sqlSTRING += "Shain.Code as shaincode,Shain.Sei,Shain.Mei, Shain.YakushokuType ";
            //        sqlSTRING += "from Shain left join Bumon ";
            //        sqlSTRING += "on Shain.BumonId = Bumon.Id ";
            //        sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
            //        sqlSTRING += "order by Shain.Code";
                    
            //        //sqlSTRING += "select Id,Code, Sei, Mei, YakushokuType from Shain ";
            //        //sqlSTRING += "where Shurojokyo = 1 and YakushokuType = 1 ";
            //        //sqlSTRING += "order by Code";

            //        //データリーダーを取得する
            //        dR = dCon.FreeReader(sqlSTRING);

            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";
            //        tempObj.ValueMember = "code";

            //        while (dR.Read())
            //        {
            //            cmb1 = new ComboShain();
            //            cmb1.ID = int.Parse(dR["shainid"].ToString());
            //            cmb1.DisplayName = dR["shaincode"].ToString().Trim() + " " + dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.Name = dR["Sei"].ToString().Trim() + "　" + dR["Mei"].ToString().Trim();
            //            cmb1.code = (dR["shaincode"].ToString() + "").Trim();
            //            cmb1.YakushokuType = int.Parse(dR["YakushokuType"].ToString());
            //            cmb1.BumonCode = dR["bumoncode"].ToString().PadLeft(3, '0');
            //            cmb1.BumonName = dR["bumonname"].ToString();
            //            tempObj.Items.Add(cmb1);
            //        }

            //        dR.Close();
            //        dCon.Close();
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show(ex.Message, "社員コンボボックスロード");
            //    }

            //}
        }


        // データ領域コンボボックスクラス
        public class ComboDataArea
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            //// データ領域ロード
            //public static void load(ComboBox tempObj)
            //{
            //    dbControl.DataControl dcon = new dbControl.DataControl(Properties.Settings.Default.SQLDataBase);
            //    OleDbDataReader dR = null;

            //    try
            //    {
            //        ComboDataArea cmb;

            //        // データリーダー取得
            //        string mySql = string.Empty;
            //        mySql += "SELECT * FROM Common_Unit_DataAreaInfo ";
            //        mySql += "where CompanyTerm = " + DateTime.Today.Year.ToString();
            //        dR = dcon.FreeReader(mySql);

            //        //会社情報がないとき
            //        if (!dR.HasRows)
            //        {
            //            MessageBox.Show("会社領域情報が存在しません", "会社領域選択", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //            return;
            //        }

            //        // コンボボックスにアイテムを追加します
            //        tempObj.Items.Clear();
            //        tempObj.DisplayMember = "DisplayName";

            //        while (dR.Read())
            //        {
            //            cmb = new ComboDataArea();
            //            // "CompanyCode"が数字のレコードを対象とする
            //            if (Utility.NumericCheck(dR["CompanyCode"].ToString()))
            //            {
            //                cmb.DisplayName = dR["CompanyName"].ToString().Trim();
            //                cmb.ID = dR["Name"].ToString().Trim();
            //                cmb.code = dR["CompanyCode"].ToString().Trim();
            //                tempObj.Items.Add(cmb);
            //            }
            //        }
            //    }
            //    catch (Exception e)
            //    {
            //        MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            //    }
            //    finally
            //    {
            //        if (!dR.IsClosed) dR.Close();
            //        dcon.Close();
            //    }

            //}

        }


        ///--------------------------------------------------------
        /// <summary>
        /// 会社情報より部門コード桁数、社員コード桁数を取得
        /// </summary>
        /// -------------------------------------------------------
        public class BumonShainKetasu
        {
            public string ID { get; set; }
            public string DisplayName { get; set; }
            public string Name { get; set; }
            public string code { get; set; }

            //// 会社情報取得
            //public static void GetKetasu(string dbName)
            //{
            //    dbControl.DataControl dcon = new dbControl.DataControl(dbName);
            //    OleDbDataReader dR = null;

            //    try
            //    {
            //        // データリーダー取得
            //        string mySql = string.Empty;
            //        mySql += "SELECT BumonCodeKeta,ShainCodeKeta FROM Kaisha ";
            //        dR = dcon.FreeReader(mySql);

            //        // 部門コード桁数、社員コード桁数を取得
            //        while (dR.Read())
            //        {
            //            global.ShozokuLength = int.Parse(dR["BumonCodeKeta"].ToString());
            //            global.ShainLength = int.Parse(dR["ShainCodeKeta"].ToString());
            //        }
            //    }
            //    catch (Exception e)
            //    {
            //        MessageBox.Show(e.Message, "エラー", MessageBoxButtons.OK);
            //    }
            //    finally
            //    {
            //        if (!dR.IsClosed) dR.Close();
            //        dcon.Close();
            //    }

            //}
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     任意のディレクトリのファイルを削除する </summary>
        /// <param name="sPath">
        ///     指定するディレクトリ</param>
        /// <param name="sFileType">
        ///     ファイル名及び形式</param>
        /// --------------------------------------------------------------------
        public static void FileDelete(string sPath, string sFileType)
        {
            //sFileTypeワイルドカード"*"は、すべてのファイルを意味する
            foreach (string files in System.IO.Directory.GetFiles(sPath, sFileType))
            {
                // ファイルを削除する
                System.IO.File.Delete(files);
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     文字列を指定文字数をＭＡＸとして返します</summary>
        /// <param name="s">
        ///     文字列</param>
        /// <param name="n">
        ///     文字数</param>
        /// <returns>
        ///     文字数範囲内の文字列</returns>
        /// --------------------------------------------------------------------
        public static string GetStringSubMax(string s, int n)
        {
            string val = string.Empty;
            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     8ケタ左詰め右空白埋めの給与大臣検索用の社員コード文字列を返す
        /// </summary>
        /// <param name="sCode">
        ///     コード</param>
        /// <returns>
        ///     給与大臣検索用の社員コード文字列</returns>
        /// --------------------------------------------------------------------
        public static string bldShainCode(string sCode)
        {
            return sCode.PadLeft(4, '0').PadRight(8, ' ').Substring(0, 8);
        }
        

        ///---------------------------------------------------------------------------
        /// <summary>
        ///     国内電話番号および携帯番号の正規かチェック </summary>
        /// <param name="tVal">
        ///     対象となるTEL/携帯番号</param>
        /// <returns>true:正しい、false:正しくない</returns>
        ///---------------------------------------------------------------------------
        public static bool regexTelNum(string tVal)
        {
            if (!Regex.IsMatch(tVal, @"^\d{2}-\d{4}-\d{4}$"))
            {
                if (!Regex.IsMatch(tVal, @"^\d{3}-\d{3}-\d{4}$"))
                {
                    if (!Regex.IsMatch(tVal, @"^\d{4}-\d{2}-\d{4}$"))
                    {
                        if (!Regex.IsMatch(tVal, @"^\d{5}-\d{1}-\d{4}$"))
                        {
                            // 携帯番号
                            if (!Regex.IsMatch(tVal, @"^\d{3}-\d{4}-\d{4}$"))
                            {
                                return false;
                            }

                        }
                    }
                }
            }

            return true;
        }


        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルを出力する</summary>
        /// <param name="sPath">
        ///     出力するパス</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        /// <param name="sFileName">
        ///     CSVファイル名</param>
        ///----------------------------------------------------------------------------
        public static void csvFileWrite(string sPath, string[] arrayData, string sFileName)
        {
            // ファイル名
            string outFileName = sPath + sFileName + ".csv";

            // 出力ファイルが存在するとき
            if (System.IO.File.Exists(outFileName))
            {
                // リネーム付加文字列（タイムスタンプ）
                string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

                // リネーム後ファイル名
                string reFileName = sPath + sFileName + newFileName + ".csv";

                // 既存のファイルをリネーム
                System.IO.File.Move(outFileName, reFileName);
            }

            // CSVファイル出力
            System.IO.File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding("utf-8"));
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     半角カナの小文字を大文字に変換 </summary>
        /// <param name="s">
        ///     変換する文字列</param>
        /// <returns>
        ///     変換後の文字列</returns>
        ///-------------------------------------------------------------
        public static string strSmallTolarge(string s)
        {
            s = s.Replace("ｬ", "ﾔ");
            s = s.Replace("ｭ", "ﾕ");
            s = s.Replace("ｮ", "ﾖ");
            s = s.Replace("ｯ", "ﾂ");

            return s;
        }

        //--------------------------------------------------------------
        /// <summary>
        ///     郵便番号住所で住所文字列を更新する </summary>
        /// <param name="st1">
        ///     住所文字列</param>
        /// <param name="st2">
        ///     郵便番号住所</param>
        /// <returns>
        ///     更新後文字列</returns>
        //--------------------------------------------------------------
        public static string addressUpdate(string st1, string st2)
        {
            st1 = Utility.strSmallTolarge(st1).Replace(" ", "").Trim();
            st2 = Utility.strSmallTolarge(st2).Replace(" ", "").Trim();

            return st2 + " " + st1.Replace(st2, "");
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     分単位を時：分形式の文字列に変換する </summary>
        /// <param name="tm">
        ///     分</param>
        /// <returns>
        ///     時：分形式の文字列</returns>
        ///-------------------------------------------------------------
        public static string intToHhMM(int tm)
        {
            string hm = ((int)(tm / 60)).ToString() + ":" + (tm % 60).ToString().PadLeft(2, '0');
            return hm;
        }

        ///-------------------------------------------------------------
        /// <summary>
        ///     分単位を時：分（10進数）形式の文字列に変換する </summary>
        /// <param name="tm">
        ///     分</param>
        /// <param name="marume">
        ///     丸め分</param>
        /// <returns>
        ///     時.分形式の文字列<</returns>
        ///-------------------------------------------------------------
        public static string intToHhMM10(int tm, int marume)
        {
            int tmm = tm - (tm % marume);
            double hm = (double)tmm / 60;

            return hm.ToString("0.00");
        }


        ///-----------------------------------------------------------------------------------
        /// <summary>
        ///     休日勤務時間、法定休日勤務時間取得【個人・工事毎】　</summary>
        /// <param name="tbl">
        ///     ryowaDataSet.T_勤怠DataTable</param>
        /// <param name="hol">
        ///     休日勤務時間</param>
        /// <param name="hotei">
        ///     法定休日勤務時間取得</param>
        /// <param name="pID">
        ///     工事ID</param>
        /// <param name="sYY">
        ///     対象年</param>
        /// <param name="sMM">
        ///     対象月</param>
        /// <param name="sNum">
        ///     社員コード</param>
        /// <returns>
        ///     true </returns>
        ///-----------------------------------------------------------------------------------
        public static bool getHolTime(genbaDataSet.T_勤怠DataTable tbl, out int hol, out int hotei, int pID, int sYY, int sMM, int sNum)
        {
            var ff = tbl
                .Where(a => a.工事ID == pID && a.日付.Year == sYY && a.日付.Month == sMM &&
                            a.社員ID == sNum && a.休日出勤 == global.flgOn);

            hol = 0;
            hotei = 0;

            foreach (var item in ff)
            {
                // 勤務時間を取得する
                double spanMin = getWorkTime(item);

                if (item.M_休日Row.法定休日 == global.flgOff)
                {
                    // 休日出勤
                    hol += (int)spanMin;
                }
                else
                {
                    // 法定休日勤務
                    hotei += (int)spanMin;
                }
            }

            return true;
        }


        ///-----------------------------------------------------------------------------------
        /// <summary>
        ///     休日勤務時間、法定休日勤務時間取得【個人毎】　</summary>
        /// <param name="tbl">
        ///     ryowaDataSet.T_勤怠DataTable</param>
        /// <param name="hol">
        ///     休日勤務時間</param>
        /// <param name="hotei">
        ///     法定休日勤務時間取得</param>
        /// <param name="sYY">
        ///     対象年</param>
        /// <param name="sMM">
        ///     対象月</param>
        /// <param name="sNum">
        ///     社員コード</param>
        /// <returns>
        ///     true </returns>
        ///-----------------------------------------------------------------------------------
        public static bool getHolTime(genbaDataSet.T_勤怠DataTable tbl, out int hol, out int hotei, int sYY, int sMM, int sNum)
        {
            var ff = tbl
                .Where(a => a.日付.Year == sYY && a.日付.Month == sMM &&
                            a.社員ID == sNum && a.休日出勤 == global.flgOn);

            hol = 0;
            hotei = 0;

            foreach (var item in ff)
            {
                // 勤務時間を取得する
                double spanMin = getWorkTime(item);

                if (item.M_休日Row.法定休日 == global.flgOff)
                {
                    // 休日出勤
                    hol += (int)spanMin;
                }
                else
                {
                    // 法定休日勤務
                    hotei += (int)spanMin;
                }
            }

            return true;
        }

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     勤務時間を取得する </summary>
        /// <param name="item">
        ///     ryowaDataSet.T_勤怠Row </param>
        /// <returns>
        ///     勤務時間・分</returns>
        ///-----------------------------------------------------------------------------
        public static double getWorkTime(genbaDataSet.T_勤怠Row item)
        {
            int cSH = Utility.StrtoInt(item.開始時刻時);
            int cSM = Utility.StrtoInt(item.開始時刻分);
            int cEH = Utility.StrtoInt(item.終了時刻時);
            int cEM = Utility.StrtoInt(item.終了時刻分);

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

            // 休憩時間を差し引く
            int rTM = Properties.Settings.Default.restTime + item.休憩;
            if (spanMin > rTM)
            {
                spanMin -= rTM;
            }

            return spanMin;
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     対象年月の休日代休時間、法定休日代休時間取得 【工事指定】</summary>
        /// <param name="dts">
        ///     ryowaDataSet</param>
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
        /// <param name="pID">
        ///     工事コード</param>
        ///----------------------------------------------------------------------------
        public static void getdaikyuTime(genbaDataSet dts, out int holDTM, out int houteiDTM, int sYY, int sMM, int sNum, int pID)
        {
            holDTM = 0;
            houteiDTM = 0;

            var s = dts.T_勤怠.Where(a => a.日付.Year == sYY && a.日付.Month == sMM &&
                                          a.社員ID == sNum && a.代休対象日 != string.Empty);

            if (s.Count() == 0)
            {
                return;
            }

            foreach (var t in s)
            {
                DateTime dt;

                if (DateTime.TryParse(t.代休対象日.Substring(0, 4) + "/" + t.代休対象日.Substring(4, 2) + "/" + t.代休対象日.Substring(6, 2), out dt))
                {
                    // 該当する工事の休日勤務か？
                    if (dts.T_勤怠.Any(a => a.社員ID == sNum && a.日付 == dt && a.工事ID == pID))
                    {
                        // 休日か？
                        if (dts.M_休日.Any(a => a.日付 == dt))
                        {
                            // 休日勤務時間を取得
                            int wTM = getHolWorkTime(dts, dt, sNum);

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
        private static int getHolWorkTime(genbaDataSet dts, DateTime dt, int sNum)
        {
            if (!dts.T_勤怠.Any(a => a.日付 == dt && a.社員ID == sNum))
            {
                return 0;
            }

            var s = dts.T_勤怠.Single(a => a.日付 == dt && a.社員ID == sNum);

            // 勤務時間を取得する
            double spanMin = Utility.getWorkTime(s);
            return (int)spanMin;
        }


        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     依頼メールを送信する </summary>
        /// <param name="attachFile">
        ///     添付ファイル名</param>
        /// <param name="sSubject">
        ///     件名</param>
        /// <param name="sBody">
        ///     メール本文</param>
        /// <param name="testFlg">
        ///     0:本番送信、1:テスト送信</param>
        ///-----------------------------------------------------------------------------
        public static void sendKintaiMail(string attachFile, string sSubject, string sBody, int testFlg)
        {
            genbaDataSet dts = new genbaDataSet();
            genbaDataSetTableAdapters.メール設定TableAdapter adp = new genbaDataSetTableAdapters.メール設定TableAdapter();
            adp.Fill(dts.メール設定);

            // メール設定情報
            if (!dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                MessageBox.Show("メール設定情報が未登録のためメール送信はできません", "メール設定未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            genbaDataSet.メール設定Row r = dts.メール設定.Single(a => a.ID == global.mailKey);

            // smtpサーバーを指定する
            SmtpClient client = new SmtpClient();
            client.Host = r.SMTPサーバー;
            client.Port = r.SMTPポート番号;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Credentials = new System.Net.NetworkCredential(r.ログイン名, r.パスワード);
            client.EnableSsl = false;
            client.Timeout = 10000;

            //メッセージインスタンス生成
            MailMessage message = new MailMessage();

            //送信元
            message.From = new MailAddress(r.メールアドレス, r.メール名称);

            //件名
            message.Subject = sSubject;

            //// Cc
            //if (txtCc.Text != string.Empty)
            //{
            //    message.CC.Add(new MailAddress(txtCc.Text));
            //}

            //// Bcc
            //if (txtBcc.Text != string.Empty)
            //{
            //    message.Bcc.Add(new MailAddress(txtBcc.Text));
            //}

            // 送信メールカウント
            int mCnt = 0;

            string toAdd = "";
            string toName = "";

            try
            {
                // 送信先テストモード
                if (Properties.Settings.Default.mailTest == global.FLGON)
                {
                    //// テスト送信先
                    //string[] toAdd = { "kyamagiwa@gmail.com", "yamagiwak@ybb.ne.jp", "e.moshichi-1212@i.softbank.jp" };

                    //for (int i = 0; i < toAdd.Length; i++)
                    //{
                    //    // 宛先
                    //    message.To.Clear();
                    //    message.To.Add(new MailAddress(toAdd[i]));

                    //    // 複数送信の時、2件目以降のCc/Bcc設定はクリアする
                    //    if (mCnt > 0)
                    //    {
                    //        message.CC.Clear();
                    //        message.Bcc.Clear();
                    //    }

                    //    // 送信する
                    //    client.Send(message);

                    //    // 送信ログ書き込み
                    //    DateTime nDt = DateTime.Now;
                    //    mllogUpdate("", toAdd[i], r.メール名称, r.メールアドレス, sSubject, sMailText, nDt);

                    //    // カウント
                    //    mCnt++;
                    //}
                }
                else if (Properties.Settings.Default.mailTest == global.FLGOFF)
                {
                    toAdd = r.送信先アドレス;
                    toName = "";

                    //宛先
                    message.To.Clear();
                    message.To.Add(new MailAddress(toAdd, toName));

                    // 複数送信の時、2件目以降のCc/Bcc設定はクリアする
                    if (mCnt > 0)
                    {
                        message.CC.Clear();
                        message.Bcc.Clear();
                    }

                    //本文
                    message.Body = sBody;

                    // 添付ファイル・勤怠データ
                    Attachment att = new Attachment(attachFile, "text/csv");
                    message.Attachments.Add(att);

                    message.BodyEncoding = System.Text.Encoding.GetEncoding(50220);

                    // 送信する
                    client.Send(message);

                    // 本送信のとき
                    if (testFlg == global.flgOff)
                    {
                        MessageBox.Show("出勤簿メールを送信しました","結果",MessageBoxButtons.OK,MessageBoxIcon.Information);

                        // ログ書き込み
                        putMaillog(toAdd, sSubject, "送信しました");
                    }

                    // カウント
                    mCnt++;
                }
            }
            catch (SmtpException ex)
            {
                //エラーメッセージ
                string errMsg = ex.Message + Environment.NewLine + Environment.NewLine + "メール情報設定画面で設定内容を確認してください。その後、再実行してください。";
                MessageBox.Show(errMsg, "メール送信失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                // ログ書き込み
                putMaillog(toAdd, sSubject, ex.Message);
            }
            finally
            {
                //// 送信あり
                //if (mCnt > 0)
                //{
                //    MessageBox.Show(mCnt.ToString() + "件の出勤簿メールを送信しました");

                //    // 本送信の時
                //    if (testFlg == global.flgOff)
                //    {
                //        // ログ書き込み
                //        putMaillog(toAdd, sSubject, "送信しました");
                //    }
                //}

                // 後片付け
                message.Dispose();
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     メールログ書き込み </summary>
        /// <param name="sto">
        ///     送信先</param>
        /// <param name="sSubject">
        ///     メール件名</param>
        /// <param name="result">
        ///     結果</param>
        ///--------------------------------------------------------------
        private static void putMaillog(string sto, string sSubject, string result)
        {
            string sLog = "送信," + DateTime.Now.ToString() + "," + sto + "," + sSubject + "," + result;
            string[] st = new string[1];
            st[0] = sLog;

            // メール送信ログデータパスファイル名
            string outMailLogFile = Properties.Settings.Default.mailLogFilePath;

            // CSVファイル出力
            var sw = new System.IO.StreamWriter(outMailLogFile, true, Encoding.UTF8);
            sw.WriteLine(st[0]);
            sw.Close();
        }

    }
}
