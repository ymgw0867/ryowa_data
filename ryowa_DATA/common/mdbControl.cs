using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;

namespace ryowa_DATA.common
{
    class mdbControl
    {
        //public OleDbCommand sCom = new OleDbCommand();
        //public OleDbCommand sCom2 = new OleDbCommand();
        protected StringBuilder sb = new StringBuilder();

        public void dbConnect(OleDbCommand cm)
        {
            // データベース接続文字列
            OleDbConnection Cn = new OleDbConnection();
            sb.Clear();
            //sb.Append("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=");
            //sb.Append(Properties.Settings.Default.mdbOlePath);
            sb.Append(Properties.Settings.Default.ryowaConnectionString);
            Cn.ConnectionString = sb.ToString();
            Cn.Open();

            cm.Connection = Cn;
            //sCom.Connection = Cn;
        }
    }
}
