using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsEngine
{
    /// <summary>
    /// Xls 設定選項
    /// </summary>
    public class xlsEngineOption
    {
        /// <summary>範本檔案路徑</summary>
        public string TemplateFilePath { get; set; }

        /// <summary>詳細資料索引</summary>
        public int DetailRowIndex { get; set; }

        /// <summary>頁尾起始索引</summary>
        public int FooterRowIndexStart { get; set; }

        /// <summary>頁尾結束索引</summary>
        public int FooterRowIndexEnd { get; set; }
    }
}
