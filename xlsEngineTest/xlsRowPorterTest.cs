using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using xlsEngine;
using System.IO;
using NPOI.HSSF.UserModel;

namespace xlsEngineTest
{
    public class CellRecord
    {
        public string Value { get; set; }
        public int RowIndex { get; set; }
        public int CellIndex { get; set; }
    }

    [TestClass]
    public class xlsRowPorterTest
    {
        const string range_footer_path = "../../files/row-porter-source.xls";
        const string single_footer_path = "../../files/row-porter-source-single-footer.xls";

        private string[] _range_footer_values = new string[] { "總計", "@sum", "日期", "@date" };
        private string[] _single_footer_values = new string[] { "總資料筆數", "@number", "總計", "@sum" };

        private xlsRowPorter _rowPorter;
        private HSSFSheet _sheet;


        [TestInitialize]
        public void Setup()
        {
        }


        [TestMethod]
        public void MoveRangeRowToNextRow()
        {
            bool _isMove;
            byte[] _content;

            IEnumerable<CellRecord> _befores;
            IEnumerable<CellRecord> _afters;

            _content = this.GetXlsFile(range_footer_path);
            _sheet = this.GetSheet(_content, 0);

            _rowPorter = new xlsRowPorter(_sheet);

            
            _befores = this.GetSheetIndex(_sheet, _range_footer_values);
            _rowPorter.Move(5, 6);
            _afters = this.GetSheetIndex(_sheet, _range_footer_values);

            _isMove = true;

            //比對儲存格內容是否被正確地進行搬移
            foreach (var _value in _range_footer_values)
            {
                var _match_before = _befores.Where(x => x.Value == _value).FirstOrDefault();
                var _match_after = _afters.Where(x => x.Value == _value).FirstOrDefault();

                if(_match_after == null)
                {
                    Console.WriteLine("下列儲存格內容並沒有被搬移");
                    Console.WriteLine("row={0},cell={1}\t{2}", _match_before.RowIndex, _match_before.CellIndex, _match_before.Value);
                    Console.WriteLine();

                    _isMove = false;
                }
                else
                { 
                    if (_match_before.RowIndex == (_match_after.RowIndex - 1))
                    {

                    }
                    else
                    {
                        //儲存格資料列並沒有正確被搬移
                        Console.WriteLine("下列儲存格內容並沒有被搬移到正確位置");
                        Console.WriteLine("row={0},cell={1}\t{2}", _match_before.RowIndex, _match_before.CellIndex, _match_before.Value);
                        Console.WriteLine("row={0},cell={1}\t{2}", _match_after.RowIndex, _match_after.CellIndex, _match_after.Value);
                        Console.WriteLine();

                        _isMove = false;
                    }
                }

            }

            Assert.IsTrue(_isMove);
        }

        [TestMethod]
        public void MoveRangeRowToSpecifyInterval()
        {
            int _interval;

            bool _isMove;
            byte[] _content;

            IEnumerable<CellRecord> _befores;
            IEnumerable<CellRecord> _afters;


            _interval = 10;

            _content = this.GetXlsFile(range_footer_path);
            _sheet = this.GetSheet(_content, 0);

            _rowPorter = new xlsRowPorter(_sheet);


            _befores = this.GetSheetIndex(_sheet, _range_footer_values);
            _rowPorter.Move(5, 6, _interval);
            _afters = this.GetSheetIndex(_sheet, _range_footer_values);

            _isMove = true;

            //比對儲存格內容是否被正確地進行搬移
            foreach (var _value in _range_footer_values)
            {
                var _match_before = _befores.Where(x => x.Value == _value).FirstOrDefault();
                var _match_after = _afters.Where(x => x.Value == _value).FirstOrDefault();

                if (_match_after == null)
                {
                    Console.WriteLine("下列儲存格內容並沒有被搬移");
                    Console.WriteLine("row={0},cell={1}\t{2}", _match_before.RowIndex, _match_before.CellIndex, _match_before.Value);
                    Console.WriteLine();

                    _isMove = false;
                }
                else
                {
                    if (_match_before.RowIndex == (_match_after.RowIndex - _interval))
                    {

                    }
                    else
                    {
                        //儲存格資料列並沒有正確被搬移
                        Console.WriteLine("下列儲存格內容並沒有被搬移到正確位置");
                        Console.WriteLine("row={0},cell={1}\t{2}", _match_before.RowIndex, _match_before.CellIndex, _match_before.Value);
                        Console.WriteLine("row={0},cell={1}\t{2}", _match_after.RowIndex, _match_after.CellIndex, _match_after.Value);
                        Console.WriteLine();

                        _isMove = false;
                    }
                }

            }

            Assert.IsTrue(_isMove);
        }


        [TestMethod]
        public void MoveSingleRowToNextRow()
        {
            bool _isMove;
            byte[] _content;

            IEnumerable<CellRecord> _befores;
            IEnumerable<CellRecord> _afters;

            _content = this.GetXlsFile(single_footer_path);
            _sheet = this.GetSheet(_content, 0);

            _rowPorter = new xlsRowPorter(_sheet);


            _befores = this.GetSheetIndex(_sheet, _single_footer_values);
            _rowPorter.Move(5, 5);
            _afters = this.GetSheetIndex(_sheet, _single_footer_values);

            _isMove = true;

            //比對儲存格內容是否被正確地進行搬移
            foreach (var _value in _single_footer_values)
            {
                var _match_before = _befores.Where(x => x.Value == _value).FirstOrDefault();
                var _match_after = _afters.Where(x => x.Value == _value).FirstOrDefault();

                if (_match_after == null)
                {
                    Console.WriteLine("下列儲存格內容並沒有被搬移");
                    Console.WriteLine("row={0},cell={1}\t{2}", _match_before.RowIndex, _match_before.CellIndex, _match_before.Value);
                    Console.WriteLine();

                    _isMove = false;
                }
                else
                {
                    if (_match_before.RowIndex == (_match_after.RowIndex - 1))
                    {

                    }
                    else
                    {
                        //儲存格資料列並沒有正確被搬移
                        Console.WriteLine("下列儲存格內容並沒有被搬移到正確位置");
                        Console.WriteLine("row={0},cell={1}\t{2}", _match_before.RowIndex, _match_before.CellIndex, _match_before.Value);
                        Console.WriteLine("row={0},cell={1}\t{2}", _match_after.RowIndex, _match_after.CellIndex, _match_after.Value);
                        Console.WriteLine();

                        _isMove = false;
                    }
                }

            }

            Assert.IsTrue(_isMove);
        }



        private byte[] GetXlsFile(string path)
        {
            byte[] _content;

            _content = File.ReadAllBytes(path);

            return (_content);
        }

        private HSSFSheet GetSheet(byte[] content, int index)
        {
            HSSFWorkbook _workbook;
            HSSFSheet _sheet;

            using (MemoryStream stream = new MemoryStream(content))
            {
                _workbook = new HSSFWorkbook(stream);
                _sheet = _workbook.GetSheetAt(index) as HSSFSheet;
            }

            return _sheet;
        }

        /// <summary>取得指定儲存格內容的索引位置= [列索引、行索引]</summary>
        /// <param name="sheet">指定工作表</param>
        /// <returns></returns>
        private IEnumerable<CellRecord> GetSheetIndex(HSSFSheet sheet, IEnumerable<string> values)
        {
            var _records = new List<CellRecord>();

            HSSFRow _row;
            string _value;

            for (int rowIndex = 0; rowIndex <= _sheet.LastRowNum; rowIndex++)
            {
                _row = _sheet.GetRow(rowIndex) as HSSFRow;

                if (_row == null)
                    continue;

                for (int cellIndex = 0; cellIndex <= _row.LastCellNum; cellIndex++)
                {
                    var _cell = _sheet.GetRow(rowIndex).GetCell(cellIndex);

                    if (_cell == null)
                        continue;

                    _value = _cell.StringCellValue;

                    if (values.Contains(_value) == true)
                    {
                        _records.Add(new CellRecord() {
                            Value = _value,
                            RowIndex = rowIndex,
                            CellIndex = cellIndex
                        });
                    }
                    
                }
            }

            return _records;
        }

    }
}
