using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using xlsEngine;

namespace xlsEngineTest
{
    public sealed class OrderReportXlsEngine : xlsEngineProvider
    {
        protected override void SetOption(xlsEngineOption option)
        {
            option.TemplateFilePath = @"../../files/row-source.xls";
            option.DetailRowIndex = 3;
            option.FooterRowIndexStart = 4;
            option.FooterRowIndexEnd = 5;
        }

        protected override void SetRowMap(HSSFRow row, DataRow record, DataRow recordBefore)
        {
            this.SetCell(row, record, recordBefore, 0, "Schema");
            this.SetCell(row, record, recordBefore, 1, "Name");
            this.SetCell(row, record, recordBefore, 2, "Quantity");
            this.SetCell(row, record, recordBefore, 3, "UnitPrice");
            this.SetCell(row, record, recordBefore, 4, "Amount");
        }
    }

    public sealed class OrderReportManager
    {
        public DataTable GetOrderReport()
        {
            DataTable _table;
            DataRow _row;
            
            _table = new DataTable();
            _table.Columns.Add("Schema", typeof(string));
            _table.Columns.Add("Name", typeof(string));
            _table.Columns.Add("Quantity", typeof(decimal));
            _table.Columns.Add("UnitPrice", typeof(decimal));
            _table.Columns.Add("Amount", typeof(decimal));
            
            _row = _table.NewRow();
            _row["Schema"] = "70-461";
            _row["Name"] = "Querying Microsoft SQL Server 2012/2014";
            _row["Quantity"] = 1m;
            _row["UnitPrice"] = 165m;
            _row["Amount"] = 165m;
            _table.Rows.Add(_row);

            _row = _table.NewRow();
            _row["Schema"] = "70-462";
            _row["Name"] = "Administering Microsoft SQL Server 2012/2014 Databases";
            _row["Quantity"] = 1m;
            _row["UnitPrice"] = 165m;
            _row["Amount"] = 165m;
            _table.Rows.Add(_row);

            _row = _table.NewRow();
            _row["Schema"] = "70-463";
            _row["Name"] = "Implementing a Data Wared:house with Microsoft SQL Server 2012/2014";
            _row["Quantity"] = 1m;
            _row["UnitPrice"] = 165m;
            _row["Amount"] = 165m;
            _table.Rows.Add(_row);

            return _table;
        }
    }

}
