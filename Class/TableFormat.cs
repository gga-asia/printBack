using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Class
{
    #region"20240919表格修改"
    class TableFormat
    {
        public int TableNum { get; set; }

        public int RowNum { get; set; }

        public int IsInsert { get; set; }

        public double RowHeight { get; set; }

        public List<CellFormat> Cells { get; set; }
    }
    #endregion
}
