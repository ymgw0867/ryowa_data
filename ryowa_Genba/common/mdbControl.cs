using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace ryowa_Genba.common
{
    class mdbControl
    {
        protected StringBuilder sb = new StringBuilder();

        public void dbConnect(OleDbCommand cm)
        {
            // データベース接続文字列
            OleDbConnection Cn = new OleDbConnection();

            sb.Clear();
            sb.Append(Properties.Settings.Default.ryowagbConnectionString);
            Cn.ConnectionString = sb.ToString();
            Cn.Open();

            cm.Connection = Cn;
        }
    }
}
