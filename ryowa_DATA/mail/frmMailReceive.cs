using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using ryowa_DATA.common;
using ryowa_DATA.mail;
using System.Collections;
using System.Globalization;
using System.IO;

namespace ryowa_DATA.mail
{
    public partial class frmMailReceive : Form
    {
        public frmMailReceive()
        {
            InitializeComponent();

            timer1.Tick += Timer1_Tick;

            // メール設定読み込み
            adp.Fill(dts.メール設定);

            // メッセージID読み込み
            idAdp.Fill(dts.messageID);
        }

        //処理中ステータス
        int _mJob = global.flgOff;

        ryowaDataSet dts = new ryowaDataSet();
        ryowaDataSetTableAdapters.メール設定TableAdapter adp = new ryowaDataSetTableAdapters.メール設定TableAdapter();
        ryowaDataSetTableAdapters.messageIDTableAdapter idAdp = new ryowaDataSetTableAdapters.messageIDTableAdapter();

        private void Timer1_Tick(object sender, EventArgs e)
        {
            // メールデータ受信
            PopTest();

            // 勤怠データ更新
            data.clsDataUpdate du = new data.clsDataUpdate();
            du.csvToMdb(Properties.Settings.Default.reCsvPath);
        }

        Timer timer1 = new Timer();

        private void frmMailReceive_Load(object sender, EventArgs e)
        {
            //タイマーオフ
            timer1.Enabled = false;
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     メール受信 </summary>
        ///----------------------------------------------------------
        private void PopTest()
        {
            // メール設定データ
            if (!dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                return;
            }

            var s = dts.メール設定.Single(a => a.ID == global.mailKey);

            //処理中
            _mJob = global.flgOn;

            //マウスポインタを待機にする
            this.Cursor = Cursors.WaitCursor;

            //受信メールカウント
            int _mCount = 0;

            //ファイル名連番
            int fNumber = 0;

            //message-ID
            string _msid;

            //差出人がinterFaxのとき1
            int _fromInterFax = 0;

            // POPサーバ、ユーザ名、パスワードを設定
            string hostname = s.POPサーバー;

            ////////////////////////////////////デバッグステータス
            //////if (global.dStatus == 0) hostname = global.pblPopServerName;
            //////else hostname = "vspop.aaa.co,jp"; //デバッグ用

            string username = s.ログイン名;
            string password = s.パスワード;
            int popPort = s.POPポート番号;

            // POP サーバに接続します。
            //addListView(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), "POPサーバに接続", Color.Black);
            PopClient pop = new PopClient(hostname, popPort);

            //POPサーバへの接続障害時は何もしないで待機状態へ戻る　2011/07/20
            if (global.Msglog.StartsWith("ERR"))
            {
                //addListView(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), global.Msglog, Color.Red);

                //非処理中ステータス
                _mJob = global.flgOff;

                //マウスポインタを戻す
                this.Cursor = Cursors.Default;

                ////////////////////////////////////デバッグステータス
                //////global.dStatus = 0;

                return;
            }

            //////////////////////////////////////デバッグステータス
            ////////global.dStatus = 1;

            //addListView(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"), global.Msglog, Color.Black);

            // ログインします。
            pop.Login(username, password);

            // POP サーバに溜まっているメールのリストを取得します。
            ArrayList list = pop.GetList();

            for (int i = 0; i < list.Count; ++i)
            {
                // メール本体を取得する
                string mailtext = pop.GetMail((string)list[i]);

                // Mail クラスを作成
                Mail mail = new Mail(mailtext);

                //受信対象メールの判定値
                int _mTarget = 0;

                //message-IDを取得
                if (mail.Header["Message-ID"].Length > 0)
                {
                    _msid = MailHeader.Decode(mail.Header["Message-ID"][0]).Replace("Message-ID: ", "");

                    // 2015/11/24
                    _msid = _msid.Replace("Message-Id: ", "");
                }
                else
                {
                    _msid = string.Empty;
                }

                // 重複メールは受信しない
                if (getMessageid(_msid))
                {
                    //Content-Typeがあるメール
                    if (mail.Header["Content-Type"].Length > 0)
                    {
                        _mTarget = 1;

                        ////差出人アドレスを調べる
                        //foreach (string add in reAddress)
                        //{
                        //    if (MailHeader.Decode(mail.Header["From"][0]).Replace("From: ", "").IndexOf(add) >= 0)
                        //    {
                        //        _mTarget = 1;
                        //        break;
                        //    }
                        //}
                    }
                }

                // 受信対象メールのとき以下の処理を実行する
                if (_mTarget == 1)
                {
                    // メールデータ
                    mailData md = new mailData();
                    mailDataInitial(md);    //メールデータ初期化

                    string sStr = string.Empty;

                    // 送信日時を取得
                    if (mail.Header["Date"].Length > 0)
                    {
                        sStr = MailHeader.Decode(mail.Header["Date"][0]).Replace("Date: ", "").Trim();

                        //タイムゾーン記号を消去
                        sStr = sStr.Replace("(JST)", "").Trim();
                        sStr = sStr.Replace("(UT)", "").Trim();
                        sStr = sStr.Replace("(EST)", "").Trim();
                        sStr = sStr.Replace("(CST)", "").Trim();
                        sStr = sStr.Replace("(MST)", "").Trim();
                        sStr = sStr.Replace("(PST)", "").Trim();
                        sStr = sStr.Replace("(EDT)", "").Trim();
                        sStr = sStr.Replace("(CDT)", "").Trim();
                        sStr = sStr.Replace("(MDT)", "").Trim();
                        sStr = sStr.Replace("(PDT)", "").Trim();
                        sStr = sStr.Replace("(GMT)", "").Trim();
                        sStr = sStr.Replace("(C)", "").Trim();
                        sStr = sStr.Replace("(UTC)", "").Trim();

                        sStr = sStr.Replace("JST", "").Trim();
                        sStr = sStr.Replace("UT", "").Trim();
                        sStr = sStr.Replace("EST", "").Trim();
                        sStr = sStr.Replace("CST", "").Trim();
                        sStr = sStr.Replace("MST", "").Trim();
                        sStr = sStr.Replace("PST", "").Trim();
                        sStr = sStr.Replace("EDT", "").Trim();
                        sStr = sStr.Replace("CDT", "").Trim();
                        sStr = sStr.Replace("MDT", "").Trim();
                        sStr = sStr.Replace("PDT", "").Trim();
                        sStr = sStr.Replace("GMT", "").Trim();
                        sStr = sStr.Replace("C", "").Trim();
                        sStr = sStr.Replace("UTC", "").Trim();

                        //dg1[0, dg1.Rows.Count - 1].Value = DateTime.Parse(sStr).ToString("yyyy/M/dd HH:mm:ss");

                        // 世界標準時(RFC2822表記)からDetaTime型に変換
                        try
                        {
                            string[] expectedFormats = { "ddd, d MMM yyyy HH':'mm':'ss", "ddd, d MMM yyyy HH':'mm':'ss zzz", "d MMM yyyy HH':'mm':'ss", "d MMM yyyy HH':'mm':'ss zzz", "d", "D", "f", "F", "g", "G", "m", "r", "R", "s", "t", "T", "u", "U", "y" };
                            DateTime myUtcDT3 = System.DateTime.ParseExact(sStr, expectedFormats, DateTimeFormatInfo.InvariantInfo, DateTimeStyles.None);
                            md.sendTime = myUtcDT3.ToString("yyyy/MM/dd HH:mm:ss");
                        }
                        catch (Exception)
                        {
                            md.sendTime = sStr;
                        }
                    }
                    else
                    {
                        sStr = string.Empty;
                    }

                    //差出人
                    md.fromAddress = MailHeader.Decode(mail.Header["From"][0]).Replace("From: ", "");
                    md.fromAddress = md.fromAddress.Replace("'", " ");

                    //if (md.fromAddress.IndexOf(Properties.Settings.Default.interFaxAddress) >= 0)
                    //{
                    //    _fromInterFax = 1;
                    //}
                    //else
                    //{
                    //    _fromInterFax = 0;
                    //}

                    //宛先
                    if (mail.Header["To: "].Length > 0)
                    {
                        md.toAddress = MailHeader.Decode(mail.Header["To: "][0]).Replace("To:", "");
                        md.toAddress = md.toAddress.Replace("'", " ");
                    }
                    else
                    {
                        md.toAddress = string.Empty;
                    }

                    //if (_fromInterFax == 1) //差出人がinterfaxのとき
                    //{
                    //    md.toAddress = MailHeader.Decode(mail.Header["X-Interfax-InterFAXNumber:"][0]).Replace("X-Interfax-InterFAXNumber: ", "");
                    //    md.toAddress = md.toAddress.Replace("'", " ");
                    //}
                    //else
                    //{
                    //    md.toAddress = MailHeader.Decode(mail.Header["To:"][0]).Replace("To: ", "");
                    //    md.toAddress = md.toAddress.Replace("'", " ");
                    //}

                    //件名
                    //件名がないとき 2011/06/27
                    if (mail.Header["Subject"].Length > 0)
                    {
                        md.subject = MailHeader.Decode(mail.Header["Subject"][0]).Replace("Subject: ", "");
                        md.subject = md.subject.Replace("'", " ");
                    }

                    //マルチパートの判定
                    int mp = mail.Body.Multiparts.Length;

                    //マルチパート　または multipart/alternativeの判断
                    if (mp == 0)    //マルチパートではない
                    {
                        //Content-Type
                        sStr = MailHeader.Decode(mail.Header["Content-Type"][0]).Replace("Content-Type:", "").Trim();

                        //charset
                        if (sStr.IndexOf("charset=") > -1)
                        {
                            string sCharset = sStr.Substring(sStr.IndexOf("charset=")).Replace("charset=", "");
                            sCharset = sCharset.Replace(@"""", "");

                            // 2015/11/21
                            int cs = sCharset.IndexOf(";");
                            if (cs > -1)
                            {
                                // 2015/11/21 utf-8; reply-type=originalの「; reply-type=original」部を消去する
                                string cc = sCharset.Substring(cs, sCharset.Length - cs);
                                sCharset = sCharset.Replace(cc, "");
                            }

                            // メール本文を取得する
                            // Content-Type の charset を参照してデコード
                            if (mail.Header["Content-Transfer-Encoding"].Length > 0)
                            {
                                sStr = MailHeader.Decode(mail.Header["Content-Transfer-Encoding"][0]).Replace("Content-Transfer-Encoding:", "").Trim();
                                byte[] bytes;
                                if (sStr == "base64" || sStr == "BASE64")
                                {
                                    bytes = Convert.FromBase64String(mail.Body.Text);
                                }
                                else
                                {
                                    bytes = Encoding.ASCII.GetBytes(mail.Body.Text);
                                }

                                // 2016/02/18
                                if (sCharset == "cp932")
                                {
                                    sCharset = "Shift_JIS";
                                }

                                string mailbody = Encoding.GetEncoding(sCharset).GetString(bytes);
                                md.message = mailbody.Replace("'", " ");
                            }
                            else
                            {
                                md.message = string.Empty;
                            }
                        }
                        else
                        {
                            md.message = string.Empty;
                        }
                    }
                    else   //マルチパートのとき
                    {
                        //本文の確認　2011/06/27
                        for (int ix = 0; ix < mp; ix++)
                        {
                            //Content-Type を検証する
                            MailMultipart part1 = mail.Body.Multiparts[ix];

                            //マルチパートの更に中のマルチパート数を取得する　2011/06/27
                            int mb = part1.Body.Multiparts.Length;

                            //マルチパートの中のマルチパート毎の"Content-Type"を検証する　2011/06/27
                            for (int n = 0; n < mb; n++)
                            {
                                //Content-Type を検証する
                                MailMultipart p = part1.Body.Multiparts[n];
                                sStr = MailHeader.Decode(p.Header["Content-Type"][0]).Replace("Content-Type:", "").Trim();

                                //本文（"text/plain"）か？　2011/06/27
                                if (sStr.IndexOf("text/plain") >= 0)
                                {
                                    //charset
                                    if (sStr.IndexOf("charset=") > -1)
                                    {
                                        string sCharset = sStr.Substring(sStr.IndexOf("charset=")).Replace("charset=", "");
                                        sCharset = sCharset.Replace(@"""", "");

                                        // 2015/11/21
                                        int cs = sCharset.IndexOf(";");
                                        if (cs > -1)
                                        {
                                            // 2015/11/21 utf-8; reply-type=originalの「; reply-type=original」部を消去する
                                            string cc = sCharset.Substring(cs, sCharset.Length - cs);
                                            sCharset = sCharset.Replace(cc, "");
                                        }

                                        //エンコード名以降の文字列を削除する
                                        int m = sCharset.IndexOf(";");
                                        if (m >= 0)
                                        {
                                            sCharset = sCharset.Remove(m);
                                        }

                                        // メール本文を取得する
                                        // Content-Type の charset を参照してデコード
                                        //Content-Transfer-Encodingを取得する
                                        if (mail.Header["Content-Transfer-Encoding"].Length > 0)
                                        {
                                            sStr = MailHeader.Decode(p.Header["Content-Transfer-Encoding"][0]).Replace("Content-Transfer-Encoding:", "").Trim();
                                            byte[] bytes;
                                            if (sStr == "base64" || sStr == "BASE64")
                                            {
                                                bytes = Convert.FromBase64String(p.Body.Text);
                                            }
                                            else
                                            {
                                                bytes = Encoding.ASCII.GetBytes(p.Body.Text);
                                            }

                                            // 2016/02/18
                                            if (sCharset == "cp932")
                                            {
                                                sCharset = "Shift_JIS";
                                            }

                                            string mailbody = Encoding.GetEncoding(sCharset).GetString(bytes);
                                            md.message = mailbody.Replace("'", " ");
                                        }
                                        else
                                        {
                                            md.message = string.Empty;
                                        }
                                    }
                                    break;
                                }
                            }
                        }

                        // 添付ファイルの確認
                        //↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓

                        for (int ix = 1; ix < mp; ix++)
                        {
                            //Content-Type を検証する
                            MailMultipart part2 = mail.Body.Multiparts[ix];
                            sStr = MailHeader.Decode(part2.Header["Content-Type"][0]).Replace("Content-Type:", "").Trim();

                            int fType = global.ANYFILE;

                            //添付ファイルがCSVファイルならば保存する
                            if (sStr.Contains("text/csv"))
                            {
                                fType = global.CSVFILE;

                                //Content-Transfer-Encodingを取得する
                                sStr = MailHeader.Decode(part2.Header["Content-Transfer-Encoding"][0]).Replace("Content-Transfer-Encoding:", "").Trim();

                                if (sStr == "base64")
                                {
                                    byte[] bytes = Convert.FromBase64String(part2.Body.Text);

                                    string fName;
                                    fName = string.Format("{0:0000}", DateTime.Today.Year) +
                                            string.Format("{0:00}", DateTime.Today.Month) +
                                            string.Format("{0:00}", DateTime.Today.Day) +
                                            string.Format("{0:00}", DateTime.Now.Hour) +
                                            string.Format("{0:00}", DateTime.Now.Minute) +
                                            string.Format("{0:00}", DateTime.Now.Second);

                                    fNumber++;

                                    //保存フォルダがあるか？なければ作成する（CSVフォルダ）
                                    if (!System.IO.Directory.Exists(Properties.Settings.Default.reCsvPath))
                                    {
                                        System.IO.Directory.CreateDirectory(Properties.Settings.Default.reCsvPath);
                                    }

                                    fName = Properties.Settings.Default.reCsvPath + fName + string.Format("{0:000}", fNumber) + ".csv";

                                    using (Stream stm = File.Open(fName, FileMode.Create))
                                    using (BinaryWriter bw = new BinaryWriter(stm))
                                    {
                                        bw.Write(bytes);
                                    }

                                    //添付ファイル名
                                    md.addFilename = fName;
                                }
                            }
                        }
                    }

                    //// 確認用に取得したメールをそのままカレントディレクトリに書き出します。
                    //using (StreamWriter sw = new StreamWriter(DateTime.Now.ToString("yyyyMMddHHmmssfff") + i.ToString("0000") + ".txt"))
                    //{
                    //    sw.Write(mailtext);
                    //}

                    // メールを POP サーバから取得します。
                    // ★注意★
                    // 削除したメールを元に戻すことはできません。
                    // 本当に削除していい場合は以下のコメントをはずしてください。
                    //pop.DeleteMail((string)list[i]);

                    // messageid履歴書き込み
                    messageidUpDate(_msid);

                    //// 通信ログ書き込み
                    //mailLogUpDate(md);

                    _mCount++;
                }
            }

            // 切断する
            pop.Close();

            /////以下デバッグ用でコメント化 2011/07/11 //////////////////////////////////////////////////////
            /////↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓↓
            ////Excelファイルから汎用データを作成する
            //ExcelToData(Cn);
            /////↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑↑

            //非処理中ステータス
            _mJob = global.JOBOUT;

            //マウスポインタを戻す
            this.Cursor = Cursors.Default;
        }
        
        ///--------------------------------------------------------
        /// <summary>
        ///     メールデータ初期化 </summary>
        /// <param name="md">
        ///     メールデータインスタンス</param>
        /// <returns>
        ///     メールデータ</returns>
        ///--------------------------------------------------------
        private void mailDataInitial(mailData md)
        {
            md.sendTime = string.Empty;
            md.toAddress = string.Empty;
            md.fromAddress = string.Empty;
            md.subject = string.Empty;
            md.message = string.Empty;
            md.memo = string.Empty;
            md.addFilename = string.Empty;
        }


        ///---------------------------------------------------------
        /// <summary>
        ///     過去のmessageidを参照する </summary>
        /// <param name="stemp">
        ///     検証するmessageid:文字列</param>
        /// <returns>
        ///     過去になし：true,過去にあり：false</returns>
        ///---------------------------------------------------------
        private bool getMessageid(string stemp)
        {
            Boolean tf = false;

            try
            {
                // messageID読み込み
                idAdp.Fill(dts.messageID);

                if (dts.messageID.Count(a => a.message == stemp) > 0)
                {
                    tf = false;
                }
                else
                {
                    tf = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            return tf;
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     MessageID履歴を書き込む </summary>
        ///-------------------------------------------------------------------------
        private void messageidUpDate(string tempid)
        {
            // messageID読み込み
            idAdp.Fill(dts.messageID);

            // messageIDテーブル更新
            ryowaDataSet.messageIDRow r = dts.messageID.NewmessageIDRow();
            r.受信日時 = DateTime.Now;
            r.message = tempid;

            dts.messageID.AddmessageIDRow(r);

            // テーブル更新
            idAdp.Update(dts.messageID);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // データ受信
            PopTest();

            // 勤怠データ更新
            data.clsDataUpdate du = new data.clsDataUpdate();
            du.csvToMdb(Properties.Settings.Default.reCsvPath);

            //インターバルセット
            timer1.Interval = Properties.Settings.Default.receiveSpan * 1000; // 秒単位
            timer1.Enabled = true;


        }
    }
}
