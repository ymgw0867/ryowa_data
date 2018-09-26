using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LINQtoCSV;

namespace ryowa_DATA.common
{
    class clsLinqCsv
    {
        [CsvColumn(FieldIndex = 1)]
        public int sID { get; set; }

        [CsvColumn(FieldIndex = 2)]
        public string sName { get; set; }

        [CsvColumn(FieldIndex = 3)]
        public int pID { get; set; }

        [CsvColumn(FieldIndex = 4)]
        public string pName { get; set; }

        [CsvColumn(FieldIndex = 5)]
        public int sJinkanhi { get; set; }

        [CsvColumn(FieldIndex = 6)]
        public decimal sHaichiDays { get; set; }

        [CsvColumn(FieldIndex = 7)]
        public string sGanbaDays { get; set; }

        [CsvColumn(FieldIndex = 8)]
        public string sKinmuchiDays { get; set; }

        [CsvColumn(FieldIndex = 9)]
        public string sStayDays { get; set; }

        [CsvColumn(FieldIndex = 10)]
        public decimal sHolTM { get; set; }

        [CsvColumn(FieldIndex = 11)]
        public decimal sHouteiTM { get; set; }

        [CsvColumn(FieldIndex = 12)]
        public decimal sZanTM { get; set; }

        [CsvColumn(FieldIndex = 13)]
        public decimal sSiTM { get; set; }

        [CsvColumn(FieldIndex = 14)]
        public string sJyosetsu { get; set; }

        [CsvColumn(FieldIndex = 15)]
        public string sTokushu { get; set; }

        [CsvColumn(FieldIndex = 16)]
        public string sTooshi { get; set; }

        [CsvColumn(FieldIndex = 17)]
        public string sYakan { get; set; }

        [CsvColumn(FieldIndex = 18)]
        public string sShokumu { get; set; }
    }
}
