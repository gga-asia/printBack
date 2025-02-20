using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1.Class
{
    #region"20240919表格修改"
    class CellFormat
    {
        public int CellNum { get; set; }
        public string SplitType { get; set; } // 分割方向：Vertical, Horizontal, None
        public bool SplitShareLine { get; set; } // 分割後是否顯示中格線
        public string MergeDirection { get; set; } // 合並方向：Vertical, Horizontal, None
        public int MergeCount { get; set; } // 合並格的数量
    }
    #endregion
}
