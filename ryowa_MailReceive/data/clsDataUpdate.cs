using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ryowa_MailReceive.common;
using System.Windows.Forms;
using System.Net.Mail;
using Excel = Microsoft.Office.Interop.Excel;

namespace ryowa_MailReceive.data
{
    class clsDataUpdate
    {
        public clsDataUpdate()
        {
            adp.Fill(dts.T_勤怠);
            sAdp.Fill(dts.M_社員);
            mAdp.Fill(dts.メール設定);
        }

        mailReceiveDataSet dts = new mailReceiveDataSet();
        mailReceiveDataSetTableAdapters.T_勤怠TableAdapter adp = new mailReceiveDataSetTableAdapters.T_勤怠TableAdapter();
        mailReceiveDataSetTableAdapters.M_社員TableAdapter sAdp = new mailReceiveDataSetTableAdapters.M_社員TableAdapter();
        mailReceiveDataSetTableAdapters.メール設定TableAdapter mAdp = new mailReceiveDataSetTableAdapters.メール設定TableAdapter();
        
        /// -------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルからT_勤怠の確認チェックを更新する </summary>
        /// <param name="_InPath">
        ///     CSVファイルパス</param>
        /// -------------------------------------------------------------------
        public int checkToMdb(string _InPath, string fTitle)
        {
            int cnt = 0;

            try
            {
                foreach (var file in System.IO.Directory.GetFiles(_InPath))
                {
                    // 出勤簿更新対象ファイル以外はネグる
                    if (!System.IO.Path.GetFileName(file).Contains(fTitle))
                    {
                        continue;
                    }

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(file, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // 出勤簿データフィールド数検証
                        if (stCSV.Length != global.CSVLENGTH_chk)
                        {
                            // 確認送信データではないため、ネグる
                            continue;
                        }

                        DateTime cdt;
                        if (!DateTime.TryParse(stCSV[1], out cdt))
                        {
                            // 2項目が日付ではないため、ネグる
                            continue;
                        }

                        // 個人コード認証
                        if (!shainAuth(Utility.StrtoInt(stCSV[0])))
                        {
                            // 認証不可のときメール返信
                            sendUnAuthMail(stCSV);
                            continue;
                        }

                        // 勤怠データ認証
                        if (dts.T_勤怠.Any(a => a.日付 == DateTime.Parse(stCSV[1]) && a.社員ID == Utility.StrtoInt(stCSV[0])))
                        {
                            // 確認更新
                            sCheckedMaster(stCSV);
                        }

                        cnt++;
                    }

                    // CSVファイル削除
                    System.IO.File.Delete(file);
                }

                // データベースへ反映
                if (cnt > 0)
                {
                    adp.Update(dts.T_勤怠);
                }

                return cnt;
            }
            catch (Exception ex)
            {
                return -1;
                //MessageBox.Show(ex.Message, "勤怠データ確認チェック更新処理", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     出勤簿CSVファイルからT_勤怠を更新する </summary>
        /// <param name="_InPath">
        ///     CSVファイルパス</param>
        /// <param name="fTitle">
        ///     </param>
        /// -------------------------------------------------------------------
        public int csvToMdb(string _InPath, string fTitle)
        {
            int cnt = 0;

            try
            {
                foreach (var file in System.IO.Directory.GetFiles(_InPath))
                {
                    // 出勤簿更新対象ファイル以外はネグる
                    // ※受信時につけた添付ファイル名
                    if (!System.IO.Path.GetFileName(file).Contains(fTitle))
                    {
                        continue;
                    }

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(file, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // 出勤簿データフィールド数検証
                        if (stCSV.Length != global.CSVLENGTH)
                        {
                            // 出勤簿データではないため、ネグる
                            continue;
                        }

                        DateTime cdt;
                        if (!DateTime.TryParse(stCSV[0], out cdt))
                        {
                            // 先頭項目が日付ではないため、ネグる
                            continue;
                        }

                        // 個人コード認証
                        if (!shainAuth(Utility.StrtoInt(stCSV[1])))
                        {
                            // 認証不可のときメール返信
                            sendUnAuthMail(stCSV);
                            continue;
                        }

                        // 勤怠データ認証
                        if (dts.T_勤怠.Any(a => a.日付 == DateTime.Parse(stCSV[0]) && a.社員ID == Utility.StrtoInt(stCSV[1])))
                        {
                            // 登録済みのとき：上書き更新
                            sOverWriteMaster(stCSV);
                        }
                        else
                        {
                            // 勤怠データ追加登録
                            sAddMaster(stCSV);
                        }

                        cnt++;
                    }

                    // CSVファイル削除
                    System.IO.File.Delete(file);
                }

                // データベースへ反映
                if (cnt > 0)
                {
                    adp.Update(dts.T_勤怠);
                }

                return cnt;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "勤怠データ更新処理", MessageBoxButtons.OK);
                return -1; ;
            }
            finally
            {
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>
        ///     出勤簿CSVファイルからT_勤怠を更新する </summary>
        /// <param name="_InPath">
        ///     CSVファイルパス</param>
        /// -------------------------------------------------------------------
        public int reqExcel(string _InPath, string fTitle)
        {
            int cnt = 0;
            string[] xlsArray = null;
            string toAdd = "";

            // 出勤簿・車両走行報告書エクセルシート出力先フォルダが存在するか？
            if (System.IO.Directory.Exists(Properties.Settings.Default.xlsOutPath))
            {
                // 出力先フォルダ内のファイルをすべて削除する
                foreach (var file in System.IO.Directory.GetFiles(Properties.Settings.Default.xlsOutPath))
                {
                    System.IO.File.Delete(file);
                }
            }
            else
            {
                // なければ作成します
                System.IO.Directory.CreateDirectory(Properties.Settings.Default.xlsOutPath);
            }

            try
            {
                // 添付ファイルをよむ
                foreach (var file in System.IO.Directory.GetFiles(_InPath))
                {
                    // 出勤簿要求ファイル以外はネグる
                    // ※受信時につけた添付ファイル名
                    if (!System.IO.Path.GetFileName(file).Contains(fTitle))
                    {
                        continue;
                    }

                    // CSVファイルインポート
                    var s = System.IO.File.ReadAllLines(file, Encoding.Default);
                    foreach (var stBuffer in s)
                    {
                        // カンマ区切りで分割して配列に格納する
                        string[] stCSV = stBuffer.Split(',');

                        // 出勤簿要求データフィールド数検証
                        if (stCSV.Length != global.CSVLENGTH_req)
                        {
                            // 出勤簿要求データではないため、ネグる
                            continue;
                        }

                        // 個人コード認証
                        if (!shainAuth(Utility.StrtoInt(stCSV[2])))
                        {
                            // 認証不可のときメール返信
                            sendUnAuthMailReq(stCSV);
                            continue;
                        }

                        // 出勤簿・車両走行報告書excelシートを作成して返信する
                        string xlsFile = sReport(Utility.StrtoInt(stCSV[0]), Utility.StrtoInt(stCSV[1]), Utility.StrtoInt(stCSV[2]));

                        // 添付ファイル用の配列を作成します
                        if (xlsFile != string.Empty)
                        {
                            Array.Resize(ref xlsArray, cnt + 1);
                            xlsArray[cnt] = xlsFile;
                            toAdd = stCSV[3];
                        }

                        cnt++;
                    }

                    if (xlsArray != null)
                    {
                        // 出勤簿・車両走行報告書エクセルシート返信
                        sendXlsFile(toAdd, xlsArray);
                    }

                    // CSVファイル削除
                    System.IO.File.Delete(file);
                }

                return cnt;
            }
            catch (Exception ex)
            {
                return -1;
                //MessageBox.Show(ex.Message, "出勤簿・車両走行報告書要求処理", MessageBoxButtons.OK);
            }
            finally
            {
            }
        }

        ///--------------------------------------------------------------------
        /// <summary>
        ///     個人コード認証 </summary>
        /// <param name="sNum">
        ///     認証する個人コード</param>
        /// <returns>
        ///     true:認証、false:認証不可</returns>
        ///--------------------------------------------------------------------
        private bool shainAuth(int sNum)
        {
            if (dts.M_社員.Any(a => a.ID == sNum))
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        

        /// ---------------------------------------------------------------
        /// <summary>
        ///     勤怠データ追加 </summary>
        /// <param name="c">
        ///     CSVデータ配列</param>
        /// ---------------------------------------------------------------
        private void sAddMaster(string[] c)
        {
            mailReceiveDataSet.T_勤怠Row r = dts.T_勤怠.NewT_勤怠Row();
            dts.T_勤怠.AddT_勤怠Row(setMasterRow(r, c));
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     勤怠データ確認チェック更新 </summary>
        /// <param name="c">
        ///     CSVデータ配列</param>
        /// ---------------------------------------------------------------
        private void sCheckedMaster(string[] c)
        {
            mailReceiveDataSet.T_勤怠Row r = dts.T_勤怠.Single(a => a.日付 == DateTime.Parse(c[1]) && a.社員ID == Utility.StrtoInt(c[0]));
            setCheckedRow(r, c);
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     勤怠データ上書き </summary>
        /// <param name="c">
        ///     CSVデータ配列</param>
        /// ---------------------------------------------------------------
        private void sOverWriteMaster(string[] c)
        {
            mailReceiveDataSet.T_勤怠Row r = dts.T_勤怠.Single(a => a.日付 == DateTime.Parse(c[0]) && a.社員ID == Utility.StrtoInt(c[1]));
            setMasterRow(r, c);
        }

        /// ----------------------------------------------------------------------------
        /// <summary>
        ///     T_勤怠にデータをセットする </summary>
        /// <param name="r">
        ///     ryowaDataSet.T_勤怠Row </param>
        /// <param name="c">
        ///     CSVデータ配列</param>
        /// <returns>
        ///     ryowaDataSet.T_勤怠Row</returns>
        /// ----------------------------------------------------------------------------
        private mailReceiveDataSet.T_勤怠Row setMasterRow(mailReceiveDataSet.T_勤怠Row r, string[] c)
        {
            r.日付 = DateTime.Parse(c[0]);
            r.社員ID = Utility.StrtoInt(c[1]);
            r.工事ID = Utility.StrtoInt(c[2]);
            r.出勤印 = Utility.StrtoInt(c[3]);
            r.出社時刻時 = c[4];
            r.出社時刻分 = c[5];
            r.開始時刻時 = c[6];
            r.開始時刻分 = c[7];
            r.終了時刻時 = c[8];
            r.終了時刻分 = c[9];
            r.退出時刻時 = c[10];
            r.退出時刻分 = c[11];
            r.休憩 = Utility.StrtoInt(c[12]);
            r.普通残業 = Utility.StrtoInt(c[13]);
            r.深夜残業 = Utility.StrtoInt(c[14]);
            r.休日出勤 = Utility.StrtoInt(c[15]);
            r.代休 = Utility.StrtoInt(c[16]);
            r.休日 = Utility.StrtoInt(c[17]);
            r.欠勤 = Utility.StrtoInt(c[18]);
            r.宿泊 = Utility.StrtoInt(c[19]);
            r.備考 = c[20];

            // ローカルデータから特殊勤務チェックしないので更新対象外
            //r.除雪当番 = Utility.StrtoInt(c[21]);
            //r.特殊出勤 = Utility.StrtoInt(c[22]);
            //r.通し勤務 = Utility.StrtoInt(c[23]);
            //r.夜間手当 = Utility.StrtoInt(c[24]);
            //r.職務手当 = Utility.StrtoInt(c[25]);

            r.全走行 = Utility.StrtoInt(c[26]);
            r.通勤業務走行 = Utility.StrtoInt(c[27]);
            r.私用走行 = Utility.StrtoInt(c[28]);
            r.代休対象日 = c[29];

            //r.確認印 = Utility.StrtoInt(c[30]);  // ローカルデータから確認印チェックしないので更新対象外

            r.登録年月日 = DateTime.Parse(c[31]);
            r.登録ユーザーID = Utility.StrtoInt(c[32]);
            r.更新年月日 = DateTime.Parse(c[33]);
            r.更新ユーザーID = Utility.StrtoInt(c[34]);

            // 早出残業：2018/07/13
            if (c.Length > 36)
            {
                r.早出残業 = Utility.StrtoInt(c[36]);
            }
            else
            {
                r.早出残業 = 0;
            }

            return r;
        }

        /// ----------------------------------------------------------------------------
        /// <summary>
        ///     T_勤怠に確認チェックをセットする </summary>
        /// <param name="r">
        ///     ryowaDataSet.T_勤怠Row </param>
        /// <param name="c">
        ///     CSVデータ配列</param>
        /// <returns>
        ///     ryowaDataSet.T_勤怠Row</returns>
        /// ----------------------------------------------------------------------------
        private mailReceiveDataSet.T_勤怠Row setCheckedRow(mailReceiveDataSet.T_勤怠Row r, string[] c)
        {
            r.確認印 = Utility.StrtoInt(c[2]);
            r.更新年月日 = DateTime.Now;
            return r;
        }

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     社員認証不可時の通知メールを送信する </summary>
        /// <param name="sD">
        ///     メールデータ</param>
        ///-----------------------------------------------------------------------------
        private void sendUnAuthMail(string [] sD)
        {
            // メール設定情報
            if (!dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                MessageBox.Show("メール設定情報が未登録のためメール送信はできません", "メール設定未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            mailReceiveDataSet.メール設定Row r = dts.メール設定.Single(a => a.ID == global.mailKey);

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
            message.Subject = "登録不能な勤怠データです";

            // 送信メールカウント
            int mCnt = 0;

            string toAdd = "";
            string toName = "";

            try
            {
                toAdd = sD[35];
                toName = "";

                //宛先
                message.To.Clear();
                message.To.Add(new MailAddress(toAdd, toName));

                //本文
                message.Body = "";
                message.Body += "送信された出勤簿データの個人コードがマスター登録されていない" + Environment.NewLine;
                message.Body += "コードのため、出勤簿データを更新することができませんでした。" + Environment.NewLine;
                message.Body += "ご使用の出勤簿登録システムの使用者情報を確認してください。" + Environment.NewLine + Environment.NewLine;
                message.Body += "【更新できなかった出勤簿データ】" + Environment.NewLine;
                message.Body += "日付：" + sD[0] + Environment.NewLine;
                message.Body += "個人コード：" + sD[1] + "　※マスター未登録コードです。";

                message.BodyEncoding = System.Text.Encoding.GetEncoding(50220);

                // 送信する
                client.Send(message);

                // ログ書き込み
                putMaillog(toAdd, message.Subject, "送信しました");

                // カウント
                mCnt++;
            }
            catch (SmtpException ex)
            {
                //エラーメッセージ
                string errMsg = ex.Message + Environment.NewLine + Environment.NewLine + "メール情報設定画面で設定内容を確認してください。その後、再実行してください。";
                MessageBox.Show(errMsg, "メール送信失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                // ログ書き込み
                putMaillog(toAdd, message.Subject, ex.Message);
            }
            finally
            {
                // 後片付け
                message.Dispose();
            }
        }

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     社員認証不可時の通知メールを送信する 【出勤簿要求時】</summary>
        /// <param name="sD">
        ///     メールデータ</param>
        ///-----------------------------------------------------------------------------
        private void sendUnAuthMailReq(string[] sD)
        {
            // メール設定情報
            if (!dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                MessageBox.Show("メール設定情報が未登録のためメール送信はできません", "メール設定未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            mailReceiveDataSet.メール設定Row r = dts.メール設定.Single(a => a.ID == global.mailKey);

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
            message.Subject = "マスター未登録社員の出勤簿が要求されました";

            // 送信メールカウント
            int mCnt = 0;

            string toAdd = "";
            string toName = "";

            try
            {
                toAdd = sD[3];
                toName = "";

                //宛先
                message.To.Clear();
                message.To.Add(new MailAddress(toAdd, toName));

                //本文
                message.Body = "";
                message.Body += "要求された出勤簿データの個人コードはマスター未登録です。" + Environment.NewLine;
                message.Body += "個人コードを確認して再度要求処理を行ってください。" + Environment.NewLine + Environment.NewLine;
                message.Body += "【要求された出勤簿データ】" + Environment.NewLine;
                message.Body += "対象年月：" + sD[0] + "年" + sD[1] + "月" + Environment.NewLine;
                message.Body += "個人コード：" + sD[2] + "　※マスター未登録コードです。";

                message.BodyEncoding = System.Text.Encoding.GetEncoding(50220);

                // 送信する
                client.Send(message);

                // ログ書き込み
                putMaillog(toAdd, message.Subject, "送信しました");

                // カウント
                mCnt++;

            }
            catch (SmtpException ex)
            {
                //エラーメッセージ
                string errMsg = ex.Message + Environment.NewLine + Environment.NewLine + "メール情報設定画面で設定内容を確認してください。その後、再実行してください。";
                MessageBox.Show(errMsg, "メール送信失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                // ログ書き込み
                putMaillog(toAdd, message.Subject, ex.Message);
            }
            finally
            {
                // 後片付け
                message.Dispose();
            }
        }

        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     出勤簿・車両走行報告書エクセルシートを送信する 【出勤簿要求時】</summary>
        /// <param name="sD">
        ///     メールデータ</param>
        /// <param name="file">
        ///     添付する出勤簿・車両走行報告書エクセルシートのパスを含むファイル名</param>
        ///-----------------------------------------------------------------------------
        private void sendXlsFile(string pToAdd, string [] file)
        {
            // メール設定情報
            if (!dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                MessageBox.Show("メール設定情報が未登録のためメール送信はできません", "メール設定未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            mailReceiveDataSet.メール設定Row r = dts.メール設定.Single(a => a.ID == global.mailKey);

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
            message.Subject = "出勤簿・車両走行報告書エクセルシート";

            // 送信メールカウント
            int mCnt = 0;

            string toAdd = "";
            string toName = "";

            try
            {
                toAdd = pToAdd;
                toName = "";

                //宛先
                message.To.Clear();
                message.To.Add(new MailAddress(toAdd, toName));

                //本文
                message.Body = "";
                message.Body += "要求された出勤簿・車両走行報告書エクセルシートを送付します。";

                // 添付ファイル
                for (int i = 0; i < file.Length; i++)
                {
                    Attachment att = new Attachment(file[i]);
                    message.Attachments.Add(att);
                }

                message.BodyEncoding = System.Text.Encoding.GetEncoding(50220);

                // 送信する
                client.Send(message);

                // ログ書き込み
                putMaillog(toAdd, message.Subject, "送信しました");

                // カウント
                mCnt++;
            }
            catch (SmtpException ex)
            {
                //エラーメッセージ
                string errMsg = ex.Message + Environment.NewLine + Environment.NewLine + "メール情報設定画面で設定内容を確認してください。その後、再実行してください。";
                MessageBox.Show(errMsg, "メール送信失敗", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                // ログ書き込み
                putMaillog(toAdd, message.Subject, ex.Message);
            }
            finally
            {
                // 後片付け
                message.Dispose();
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     メールログ書き込み </summary>
        /// <param name="md">
        ///     mailDataクラス </param>
        ///--------------------------------------------------------------
        private void putMaillog(string sToAdd, string sSubject, string sMemo)
        {
            string sLog = "送信," + DateTime.Now.ToString() + "," + sToAdd + "," + sSubject + "," + sMemo;
            string[] st = new string[1];
            st[0] = sLog;

            // メール送信ログデータパス名
            string outMailLogFile = Properties.Settings.Default.mailLogFilePath;

            // CSVファイル出力
            var sw = new System.IO.StreamWriter(outMailLogFile, true, Encoding.UTF8);
            sw.WriteLine(st[0]);
            sw.Close();
        }

        ///------------------------------------------------------------------
        /// <summary>
        ///     出勤簿・車両走行報告書作成 </summary>
        ///------------------------------------------------------------------
        private string sReport(int pYY, int pMM, int pNum)
        {
            string xlsName = string.Empty;

            mailReceiveDataSet dts = new mailReceiveDataSet();
            mailReceiveDataSetTableAdapters.T_勤怠TableAdapter adp = new mailReceiveDataSetTableAdapters.T_勤怠TableAdapter();
            mailReceiveDataSetTableAdapters.M_社員TableAdapter sAdp = new mailReceiveDataSetTableAdapters.M_社員TableAdapter();
            mailReceiveDataSetTableAdapters.M_休日TableAdapter hAdp = new mailReceiveDataSetTableAdapters.M_休日TableAdapter();
            mailReceiveDataSetTableAdapters.M_工事TableAdapter pAdp = new mailReceiveDataSetTableAdapters.M_工事TableAdapter();

            // 勤怠テーブル読み込み
            adp.Fill(dts.T_勤怠);
            sAdp.Fill(dts.M_社員);
            hAdp.Fill(dts.M_休日);
            pAdp.Fill(dts.M_工事);

            // 該当勤怠データがないときは終わる
            if (dts.T_勤怠.Count(a => a.日付.Year == pYY && a.日付.Month == pMM && a.社員ID == pNum) == 0)
            {
                return string.Empty;
            }
            
            //エクセルファイル日付明細開始行
            const int S_GYO = 5;
            const int S_GYO2 = 38;

            int eRow = 0;

            try
            {
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
                    //oxlsPrintSheet.Cells[2, 5] = Properties.Settings.Default.gengou;  // 2018/07/18 コメント化
                    //oxlsPrintSheet.Cells[2, 6] = (pYY - Properties.Settings.Default.rekiHosei).ToString();
                    oxlsPrintSheet.Cells[2, 6] = pYY.ToString(); // 和暦から西暦へ
                    oxlsPrintSheet.Cells[2, 7] = "年";
                    oxlsPrintSheet.Cells[2, 8] = pMM.ToString();
                    oxlsPrintSheet.Cells[2, 9] = "月分";

                    // 個人コード
                    oxlsPrintSheet.Cells[3, 4] = pNum.ToString();

                    // 該当月を取得
                    DateTime dt = DateTime.Parse(pYY.ToString() + "/" + pMM.ToString() + "/01").AddMonths(1).AddDays(-1);
                   
                    for (int i = 0; i < dt.Day; i++)
                    {
                        DateTime sDt = DateTime.Parse(pYY.ToString() + "/" + pMM.ToString() + "/" + (i + 1).ToString());

                        // 勤怠登録されている日付を対象とする
                        if (!dts.T_勤怠.Any(a => a.日付 == sDt && a.社員ID == pNum))
                        {
                            continue;
                        }

                        // 勤怠データを取得
                        var t = dts.T_勤怠.Single(a => a.日付 == sDt && a.社員ID == pNum);

                        // 氏名
                        if (t.M_社員Row != null)
                        {
                            oxlsPrintSheet.Cells[2, 2] = t.M_社員Row.氏名;
                        }
                        else
                        {
                            oxlsPrintSheet.Cells[2, 2] = "";
                        }

                        // 印字行
                        eRow = S_GYO + i;

                        oxlsPrintSheet.Cells[eRow, 1] = (i + 1).ToString();     // 日付

                        // 休日か？
                        if (t.M_休日Row != null)
                        {
                            // 網掛け 
                            rng[0] = (Excel.Range)oxlsPrintSheet.Cells[eRow, 1];
                            rng[1] = (Excel.Range)oxlsPrintSheet.Cells[eRow, 1];
                            oxlsPrintSheet.get_Range(rng[0], rng[1]).Interior.Color = System.Drawing.Color.LightGray;
                        }

                        // 工事名
                        string pName = "";
                        if (t.M_工事Row != null)
                        {
                            pName = t.M_工事Row.名称;
                        }
                        oxlsPrintSheet.Cells[eRow, 2] = pName;

                        if (t.出勤印 == global.flgOn)
                        {
                            oxlsPrintSheet.Cells[eRow, 3] = "出勤";
                        }
                        else if (t.休日出勤 == global.flgOn)
                        {
                            oxlsPrintSheet.Cells[eRow, 3] = "休出";
                        }
                        else if (t.代休 == global.flgOn)
                        {
                            oxlsPrintSheet.Cells[eRow, 3] = "代休";
                        }
                        else if (!t.Is休日Null() && t.休日 == global.flgOn)
                        {
                            oxlsPrintSheet.Cells[eRow, 3] = "休日";
                        }
                        else if (!t.Is欠勤Null() && t.欠勤 == global.flgOn)
                        {
                            oxlsPrintSheet.Cells[eRow, 3] = pName + "休み";
                        }


                        if (t.出勤印 == global.flgOn || t.休日出勤 == global.flgOn)
                        {
                            oxlsPrintSheet.Cells[eRow, 4] = t.出社時刻時 + ":" + t.出社時刻分.PadLeft(2, '0');
                            oxlsPrintSheet.Cells[eRow, 5] = t.開始時刻時 + ":" + t.開始時刻分.PadLeft(2, '0');
                            oxlsPrintSheet.Cells[eRow, 6] = t.終了時刻時 + ":" + t.終了時刻分.PadLeft(2, '0');
                            oxlsPrintSheet.Cells[eRow, 7] = t.退出時刻時 + ":" + t.退出時刻分.PadLeft(2, '0');

                            if (t.休憩 > 0)
                            {
                                oxlsPrintSheet.Cells[eRow, 8] = Utility.intToHhMM(t.休憩);
                            }
                            else
                            {
                                oxlsPrintSheet.Cells[eRow, 8] = string.Empty;
                            }

                            // 2018/07/18 コメント化
                            //if (t.普通残業 > 0)
                            //{
                            //    oxlsPrintSheet.Cells[eRow, 9] = Utility.intToHhMM(t.普通残業);
                            //}
                            //else
                            //{
                            //    oxlsPrintSheet.Cells[eRow, 9] = string.Empty;
                            //}

                            // 早出残業＋普通残業 2018/07/18
                            int zanTotal = hayadeZanTotal(Utility.nulltoStr(t.早出残業), Utility.nulltoStr(t.普通残業));

                            if (zanTotal > 0)
                            {
                                oxlsPrintSheet.Cells[eRow, 9] = Utility.intToHhMM(zanTotal);
                            }
                            else
                            {
                                oxlsPrintSheet.Cells[eRow, 9] = string.Empty;
                            }

                            if (t.深夜残業 > 0)
                            {
                                oxlsPrintSheet.Cells[eRow, 10] = Utility.intToHhMM(t.深夜残業);
                            }
                            else
                            {
                                oxlsPrintSheet.Cells[eRow, 10] = string.Empty;
                            }
                        }

                        if (t.宿泊 == global.flgOn)
                        {
                            oxlsPrintSheet.Cells[eRow, 11] = "◯";
                        }
                        else
                        {
                            oxlsPrintSheet.Cells[eRow, 11] = "";
                        }

                        if (!t.Is備考Null())
                        {
                            oxlsPrintSheet.Cells[eRow, 12] = t.備考;
                        }
                        else
                        {
                            oxlsPrintSheet.Cells[eRow, 12] = string.Empty;
                        }

                        oxlsPrintSheet.Cells[eRow, 14] = getKmAll(pNum, sDt);
                        oxlsPrintSheet.Cells[eRow, 15] = t.通勤業務走行.ToString();
                        oxlsPrintSheet.Cells[eRow, 16] = t.私用走行.ToString();

                        // 確認印
                        if (t.確認印 != global.flgOff)
                        {
                            //oxlsPrintSheet.Cells[eRow, 17] = "✔";

                            // 確認欄に社員名を表示 : 2016/04/08
                            if (dts.M_社員.Any(a => a.ID == t.確認印))
                            {
                                var s = dts.M_社員.Single(a => a.ID == t.確認印);
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
                    
                    // 工事別集計欄
                    var ss = dts.T_勤怠
                        .Where(a => a.社員ID == pNum && a.日付.Year == pYY && a.日付.Month == pMM)
                        .GroupBy(a => a.工事ID)
                        .Select(gr => new
                        {
                            pID = gr.Key,
                            pHayade = gr.Sum(a => a.早出残業),
                            pZan = gr.Sum(a => a.普通残業),
                            pSinya = gr.Sum(a => a.深夜残業),
                            pkm = gr.Sum(a => a.通勤業務走行),
                            pkmShiyou = gr.Sum(a => a.私用走行)
                        });

                    int iX = 0;

                    foreach (var t in ss)
                    {
                        var m = dts.M_工事.Single(a => a.ID == t.pID);
                        oxlsPrintSheet.Cells[S_GYO2 + iX, 1] = m.名称;
                        oxlsPrintSheet.Cells[S_GYO2 + iX, 4] = m.勤務地名;

                        // 出勤日数
                        int s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == pYY &&
                                    a.日付.Month == pMM && a.社員ID == pNum &&
                                    (a.出勤印 == global.flgOn || a.休日出勤 == global.flgOn));

                        oxlsPrintSheet.Cells[S_GYO2 + iX, 5] = s.ToString();

                        // 代休日数
                        s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == pYY &&
                                    a.日付.Month == pMM && a.社員ID == pNum &&
                                    a.代休 == global.flgOn);

                        oxlsPrintSheet.Cells[S_GYO2 + iX, 6] = s.ToString();

                        // 工事部署ごとの休日勤務時間・法定休日勤務時間を求める
                        int hol = 0;
                        int hotei = 0;
                        Utility.getHolTime(dts.T_勤怠, out hol, out hotei, t.pID, pYY, pMM, pNum);

                        // 工事部署ごとの休日代休時間・法定休日時間取得を求める
                        int holD = 0;
                        int hoteiD = 0;
                        Utility.getdaikyuTime(dts, out holD, out hoteiD, pYY, pMM, pNum, t.pID);

                        // 代休取得した時間を差し引く
                        hol -= holD;
                        hotei -= hoteiD;

                        oxlsPrintSheet.Cells[S_GYO2 + iX, 7] = Utility.intToHhMM(hol);
                        oxlsPrintSheet.Cells[S_GYO2 + iX, 8] = Utility.intToHhMM(hotei);

                        // 宿泊日数
                        s = dts.T_勤怠.Count(a => a.工事ID == t.pID && a.日付.Year == pYY &&
                                    a.日付.Month == pMM && a.社員ID == pNum &&
                                    a.宿泊 == global.flgOn);

                        oxlsPrintSheet.Cells[S_GYO2 + iX, 11] = s.ToString();

                        // 早出残業＋普通残業 2018/07/18
                        int zanTotal = hayadeZanTotal(Utility.nulltoStr(t.pHayade), Utility.nulltoStr(t.pZan));

                        if (zanTotal > 0)
                        {
                            oxlsPrintSheet.Cells[S_GYO2 + iX, 9] = Utility.intToHhMM(zanTotal);
                        }
                        else
                        {
                            oxlsPrintSheet.Cells[S_GYO2 + iX, 9] = string.Empty;
                        }

                        //oxlsPrintSheet.Cells[S_GYO2 + iX, 9] = Utility.intToHhMM(t.pZan); // 2018/07/18 コメント化
                        oxlsPrintSheet.Cells[S_GYO2 + iX, 10] = Utility.intToHhMM(t.pSinya);
                        //oxlsPrintSheet.Cells[eRow, 15] = t.pkm.ToString();
                        //oxlsPrintSheet.Cells[eRow, 16] = t.pkmShiyou.ToString();

                        iX++;

                        // 印字は5行まで
                        if (iX > 5)
                        {
                            break;
                        }
                    }
                    
                    // 今月末走行距離
                    //int sYY = pYY + Properties.Settings.Default.rekiHosei;
                    int sYY = pYY;  // 和暦から西暦へ 2018/07/13
                    DateTime dt2 = DateTime.Parse(sYY.ToString() + "/" + pMM.ToString() + "/01");
                    dt2 = dt.AddMonths(1).AddDays(-1);
                    oxlsPrintSheet.Cells[39, 15] = getKmAll(pNum, dt);

                    // 前月末走行距離
                    dt = dt.AddMonths(-1);
                    oxlsPrintSheet.Cells[40, 15] = getKmAll(pNum, dt);
                    
                    // 確認のためのウィンドウを表示する
                    //oXls.Visible = true;

                    //印刷
                    //oXlsBook.PrintPreview(true);
                    //oXlsBook.PrintOut();

                    //保存処理
                    oXls.DisplayAlerts = false;

                    xlsName = Properties.Settings.Default.xlsOutPath + pNum.ToString() + " " + pYY.ToString() + "年" + pMM.ToString() + "月 出勤簿・車両走行報告書.xlsx";
                    oXlsBook.SaveAs(xlsName);

                    //Bookをクローズ
                    oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                    //Excelを終了
                    oXls.Quit();
                }

                catch (Exception e)
                {
                    //MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

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
                }
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message, "印刷処理", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            return xlsName;
        }
        
        private int hayadeZanTotal(string gHayade, string gZan)
        {
            // 早出残業＋普通残業 2018/07/11
            //string[] aa = gHayade.Split(':');
            //string[] bb = gZan.Split(':');
            //int hayade = 0;
            //int zan = 0;

            //if (aa.Length > 1)
            //{
            //    hayade = Utility.StrtoInt(gHayade);
            //}

            //if (bb.Length > 1)
            //{
            //    zan = Utility.StrtoInt(gZan);
            //}

            return Utility.StrtoInt(gHayade) + Utility.StrtoInt(gZan);
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
        
    }
}
