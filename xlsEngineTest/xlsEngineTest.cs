using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Data;
using System.IO;

namespace xlsEngineTest
{
    [TestClass]
    public class xlsEngineTest
    {
        private readonly OrderReportManager _reportManager;
        private readonly OrderReportXlsEngine _reportXlsEngine;

        public xlsEngineTest()
        {
            this._reportManager = new OrderReportManager();
            this._reportXlsEngine = new OrderReportXlsEngine();
        }


        //此方法僅有使用基底類別產生 EXCEL 檔案，會儲存在 files 資料夾
        [TestMethod]
        public void GenerateReport()
        {
            string _path = ("../../files/row-test.xls");

            DataTable _table;
            byte[] _content;

            this._reportXlsEngine.SetFormula("@sum", "1000");
            this._reportXlsEngine.SetFormula("@date", "2017/01/18");

            _table = this._reportManager.GetOrderReport();
            _content = this._reportXlsEngine.GetReportXls(_table);
            
            File.WriteAllBytes(_path, _content);
        }
    }
    
}
