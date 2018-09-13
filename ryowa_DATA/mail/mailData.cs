using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ryowa_DATA.mail
{
    class mailData
    {
        // メールデータ
        public int status;          // 通信区分   
        public string sendTime;     // 送信時刻
        public string fromAddress;  // 差出人
        public string toAddress;    // 宛先
        public string subject;      // 件名
        public string message;      // 本文
        public string addFilename;  // 添付ファイル名
        public string memo;         // 備考
    }
}
