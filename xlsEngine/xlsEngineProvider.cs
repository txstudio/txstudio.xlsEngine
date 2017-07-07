using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsEngine
{
    /// <summary>
    /// 提供 Xls 格式報表實作類別
    /// </summary>
    public abstract class xlsEngineProvider
    {
        const string formulaPrefix = "@";

        private xlsEngineOption _option;

        protected Dictionary<string, string> _formulas;


        public xlsEngineProvider()
        {
            this._option = new xlsEngineOption();
            this._formulas = new Dictionary<string, string>();

            this.SetOption(this._option);
        }


        #region 需要複寫方法

        /// <summary>設定 Xls 選項</summary>
        /// <param name="option"></param>
        protected abstract void SetOption(xlsEngineOption option);

        /// <summary>
        /// DataRow 物件與 Excel 資料列對應
        /// </summary>
        /// <param name="row">資料列</param>
        /// <param name="record">要新增的資料物件</param>
        /// <param name="recordBefore">前一筆資料物件</param>
        protected abstract void SetRowMap(HSSFRow row, DataRow record, DataRow recordBefore);

        #endregion


        #region 公開方法

        /// <summary>
        /// 取得 Excel 格式的 byte 陣列
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public byte[] GetReportXls(DataTable table)
        {
            byte[] _inBuffer;
            byte[] _outBuffer;

            _inBuffer = File.ReadAllBytes(this._option.TemplateFilePath);


            HSSFWorkbook _workbook;
            HSSFSheet _sheet;

            using (MemoryStream _inStream = new MemoryStream(_inBuffer))
            {
                using (MemoryStream _outStream = new MemoryStream())
                {
                    _workbook = new HSSFWorkbook(_inStream);

                    _sheet = _workbook.GetSheetAt(0) as HSSFSheet;

                    //設定變數
                    this.SetFormula(_sheet);

                    //進行資料列內容初始化
                    this.SetDataRowDefault(_sheet);

                    //新增資料
                    this.SetDataRow(_sheet, table);

                    _workbook.Write(_outStream);

                    _outBuffer = _outStream.ToArray();
                }
            }

            return _outBuffer;
        }

        /// <summary>
        /// 新增公式或參數資料
        /// </summary>
        /// <param name="key">公式或參數名稱</param>
        /// <param name="value">數值</param>
        public virtual void SetFormula(string key, string value)
        {
            if (key.StartsWith(formulaPrefix) == false)
                key = ("@" + key);

            this._formulas.Add(key, value);
        }

        #endregion


        protected void SetCell(HSSFRow row
                                , DataRow record
                                , DataRow recordBefore
                                , int index
                                , string columnName
                                , ICustomFormula formula)
        {
            this.SetCell(row
                        , record
                        , recordBefore
                        , index
                        , columnName
                        , false
                        , formula);
        }

        protected void SetCell(HSSFRow row
                                , DataRow record
                                , DataRow recordBefore
                                , int index
                                , string columnName
                                , bool isDuplicateHide = false
                                , ICustomFormula formula = null)
        {
            HSSFCell _cell;
            object _recordValue;
            object _recordBeforeValue;

            _cell = row.CreateCell(index) as HSSFCell;
            _recordValue = record[columnName];
            _recordBeforeValue = recordBefore[columnName];

            if(Convert.IsDBNull(_recordValue) == true)
                return;

            if (formula != null)
                _recordValue = formula.GetValue(_recordValue);

            _cell.SetCellValue(_recordValue.ToString());


            //確認是否有隱藏重複資料
            if(isDuplicateHide == true)
                if(Convert.IsDBNull(_recordValue) == false)
                    if (_recordValue.Equals(_recordBeforeValue) == true)
                        _cell.SetCellValue(string.Empty);

        }



        /// <summary>設定變數物件</summary>
        /// <param name="sheet">要設定的工作表 (Sheet)</param>
        private void SetFormula(HSSFSheet sheet)
        {
            HSSFRow _row;
            HSSFCell _cell;

            //設定變數
            for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                _row = sheet.GetRow(rowIndex) as HSSFRow;

                if (_row == null)
                    continue;

                for (int columnIndex = 0; columnIndex < _row.LastCellNum; columnIndex++)
                {
                    _cell = _row.GetCell(columnIndex) as HSSFCell;

                    if (_cell == null)
                    {
                        continue;
                    }

                    //依序設定變數數值
                    if (CellType.String.Equals(_cell.CellType) == true)
                    {
                        var _formulaKey = _cell.StringCellValue;

                        if (_formulaKey.StartsWith(formulaPrefix) == true)
                        {
                            if (this._formulas.ContainsKey(_formulaKey) == true)
                            {
                                _cell.SetCellValue(this._formulas[_formulaKey]);
                            }
                            else
                            {
                                _cell.SetCellValue(string.Empty);
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }

                }
            }

        }

        /// <summary>逐列進行資料列的初始化</summary>
        /// <param name="sheet">要設定的工作表 (Sheet)</param>
        private void SetDataRowDefault(HSSFSheet sheet)
        {
            HSSFRow _row;
            float _defaultRowHeightInPoints;

            _defaultRowHeightInPoints = sheet.DefaultRowHeightInPoints;

            for (int rowIndex = 0; rowIndex < sheet.LastRowNum; rowIndex++)
            {
                _row = sheet.GetRow(rowIndex) as HSSFRow;

                if (_row == null)
                {
                    _row = sheet.CreateRow(rowIndex) as HSSFRow;
                    _row.HeightInPoints = _defaultRowHeightInPoints;
                    _row.CreateCell(0).SetCellValue(string.Empty);
                }
            }
        }

        /// <summary>設定資料列內容</summary>
        /// <param name="sheet"></param>
        /// <param name="table"></param>
        private void SetDataRow(HSSFSheet sheet, DataTable table)
        {
            int _rowIndex;
            int _num;
            float _defaultRowHeightInPoints;

            HSSFRow _row;
            DataRow _before;
            xlsRowPorter _porter;

            _rowIndex = this._option.DetailRowIndex;
            _num = 0;
            _defaultRowHeightInPoints = sheet.DefaultRowHeightInPoints;

            _before = table.NewRow();
            _porter = new xlsRowPorter(sheet);

            foreach (DataRow _record in table.Rows)
            {
                _row = sheet.CreateRow(_rowIndex) as HSSFRow;
                _row.HeightInPoints = _defaultRowHeightInPoints;

                //複寫資料列內容
                this.SetRowMap(_row, _record, _before);

                //搬移資料列
                _porter.Move(this._option.FooterRowIndexStart + _num
                            , this._option.FooterRowIndexEnd + _num);

                _rowIndex = _rowIndex + 1;
                _num = _num + 1;

                _before = _record;
            }
            
        }

        



        /// <summary>取得 Xls 設定物件 - 唯讀</summary>
        public xlsEngineOption EngineOption
        {
            get
            {
                return this._option;
            }
        }

    }
}
