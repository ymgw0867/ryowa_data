using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ryowa_DATA.common;
using System.Windows.Forms;
using System.Net.Mail;

namespace ryowa_DATA.data
{
    class clsDataUpdate
    {
        public clsDataUpdate()
        {
            adp.Fill(dts.T_勤怠);
            sAdp.Fill(dts.M_社員);
            mAdp.Fill(dts.メール設定);
        }

        ryowaDataSet dts = new ryowaDataSet();
        ryowaDataSetTableAdapters.T_勤怠TableAdapter adp = new ryowaDataSetTableAdapters.T_勤怠TableAdapter();
        ryowaDataSetTableAdapters.M_社員TableAdapter sAdp = new ryowaDataSetTableAdapters.M_社員TableAdapter();
        ryowaDataSetTableAdapters.メール設定TableAdapter mAdp = new ryowaDataSetTableAdapters.メール設定TableAdapter();

        /// -------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルインポート </summary>
        /// <param name="_InPath">
        ///     CSVファイルパス</param>
        /// -------------------------------------------------------------------
        public void csvToMdb(string _InPath)
        {
            int cnt = 0;
            int ng = 0;

            try
            {
                foreach (var file in System.IO.Directory.GetFiles(_InPath))
                {
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
                            ng++;
                            continue;
                        }

                        if (dts.T_勤怠.Any(a => a.日付 == DateTime.Parse(stCSV[0]) && a.社員ID == Utility.StrtoInt(stCSV[1])))
                        {
                            // 登録済みのとき：上書き更新
                            sOverWriteMaster(stCSV);
                        }
                        else
                        {
                            // 勤怠データ更新
                            sAddMaster(stCSV);
                        }

                        cnt++;
                    }

                    // CSVファイル削除
                    System.IO.File.Delete(file);
                }

                // データベースへ反映
                adp.Update(dts.T_勤怠);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "勤怠データ更新処理", MessageBoxButtons.OK);
            }
            finally
            {
                string msg = "";

                if (ng > 0)
                {
                    msg = "個人コードがマスター未登録の出勤簿データが" + ng.ToString() + "件あり" + Environment.NewLine;
                    msg += "送付元へ更新不可通知メールを送付しました。通信ログを確認してください。" + Environment.NewLine;
                    msg += "通信ログファイル：" + Properties.Settings.Default.mailLogFilePath;

                    MessageBox.Show(msg,"更新不可データ",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                }

                if (cnt > 0)
                {
                    msg = cnt.ToString() + "件の出勤簿を更新しました";
                    MessageBox.Show(msg, "処理完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    msg = "処理対象となる出勤簿データはありませんでした";
                    MessageBox.Show(msg, "処理完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
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
            ryowaDataSet.T_勤怠Row r = dts.T_勤怠.NewT_勤怠Row();
            dts.T_勤怠.AddT_勤怠Row(setMasterRow(r, c));
        }

        /// ---------------------------------------------------------------
        /// <summary>
        ///     勤怠データ上書き </summary>
        /// <param name="c">
        ///     CSVデータ配列</param>
        /// ---------------------------------------------------------------
        private void sOverWriteMaster(string[] c)
        {
            ryowaDataSet.T_勤怠Row r = dts.T_勤怠.Single(a => a.日付 == DateTime.Parse(c[0]) && a.社員ID == Utility.StrtoInt(c[1]));
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
        private ryowaDataSet.T_勤怠Row setMasterRow(ryowaDataSet.T_勤怠Row r, string[] c)
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

            // ローカルでは特殊勤務チェックしないので更新対象外
            //r.除雪当番 = Utility.StrtoInt(c[21]);
            //r.特殊出勤 = Utility.StrtoInt(c[22]);
            //r.通し勤務 = Utility.StrtoInt(c[23]);
            //r.夜間手当 = Utility.StrtoInt(c[24]);
            //r.職務手当 = Utility.StrtoInt(c[25]);

            r.全走行 = Utility.StrtoInt(c[26]);
            r.通勤業務走行 = Utility.StrtoInt(c[27]);
            r.私用走行 = Utility.StrtoInt(c[28]);
            r.代休対象日 = c[29];

            //r.確認印 = Utility.StrtoInt(c[30]);// ローカルでは確認印チェックしないので更新対象外

            r.登録年月日 = DateTime.Parse(c[31]);
            r.登録ユーザーID = Utility.StrtoInt(c[32]);
            r.更新年月日 = DateTime.Parse(c[33]);
            r.更新ユーザーID = Utility.StrtoInt(c[34]);

            return r;
        }



        ///-----------------------------------------------------------------------------
        /// <summary>
        ///     社員認証不可時の通知メールを送信する </summary>
        /// <param name="attachFile">
        ///     添付ファイル名</param>
        /// <param name="sSubject">
        ///     件名</param>
        /// <param name="sBody">
        ///     メール本文</param>
        ///-----------------------------------------------------------------------------
        private void sendUnAuthMail(string[] sD)
        {
            // メール設定情報
            if (!dts.メール設定.Any(a => a.ID == global.mailKey))
            {
                MessageBox.Show("メール設定情報が未登録のためメール送信はできません", "メール設定未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            ryowaDataSet.メール設定Row r = dts.メール設定.Single(a => a.ID == global.mailKey);

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
                toAdd = sD[35];
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
    }
}
