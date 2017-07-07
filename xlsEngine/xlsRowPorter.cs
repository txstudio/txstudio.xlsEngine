using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsEngine
{
    /// <summary>
    /// 提供 Excel 資料列搬移物件
    /// </summary>
    public class xlsRowPorter
    {
        private HSSFSheet _sheet;

        public xlsRowPorter(HSSFSheet sheet)
        {
            this._sheet = sheet;
        }


        /// <summary>
        /// 向下位移指定區間的資料列
        /// </summary>
        /// <param name="indexStart">區間起始索引</param>
        /// <param name="indexEnd">區間結束索引</param>
        public void Move(int indexStart, int indexEnd)
        {
            int _index;

            _index = indexEnd;

            while (indexStart <= _index)
            {
                this.CopyRow(_index, (_index + 1));

                _index = (_index - 1);
            }
        }

        /// <summary>
        /// 將指定區間資料列做指定數量的位移
        /// </summary>
        /// <param name="indexStart">區間起始索引</param>
        /// <param name="indexEnd">區間結束索引</param>
        /// <param name="interval">指定位移量</param>
        public void Move(int indexStart, int indexEnd, int interval)
        {
            if (interval == 0)
                return;

            for (int i = 0; i < interval; i++)
            {
                this.Move(indexStart + i, indexEnd + i);
            }
        }


        private void CopyRow(int sourceRowIndex, int destinationRowIndex)
        {
            HSSFRow _newRow;
            HSSFRow _sourceRow;

            HSSFCell _oldCell;
            HSSFCell _newCell;
            

            _newRow = this._sheet.GetRow(destinationRowIndex) as HSSFRow;
            _sourceRow = this._sheet.GetRow(sourceRowIndex) as HSSFRow;

            if (_sourceRow == null)
                return;

            if (_newRow != null)
                this._sheet.ShiftRows(destinationRowIndex, destinationRowIndex, 1);
            else
                _newRow = this._sheet.CreateRow(destinationRowIndex) as HSSFRow;


            //將來源資料列內容複製到目標資料列
            for (int index = 0; index < _sourceRow.LastCellNum; index++)
            {
                _oldCell = _sourceRow.GetCell(index) as HSSFCell;
                _newCell = _newRow.CreateCell(index) as HSSFCell;

                if (_oldCell == null)
                {
                    _newCell = null;
                    continue;
                }

                //新儲存格的樣式需與舊儲存格樣式相符
                _newCell.CellStyle = _oldCell.CellStyle;

                switch (_oldCell.CellType)
                {
                    case CellType.String:
                        _newCell.SetCellValue(_oldCell.StringCellValue);
                        break;
                    default:
                        break;
                }

            }


            //複製來源資料列的儲存格合併格式
            CellRangeAddress _rangeAddr;
            CellRangeAddress _newCellRangeAddr;

            for (int index = 0; index < this._sheet.NumMergedRegions - 1; index++)
            {
                _rangeAddr = this._sheet.GetMergedRegion(index);

                if (_rangeAddr.FirstRow == _sourceRow.RowNum)
                {
                    _newCellRangeAddr = new CellRangeAddress(_newRow.RowNum
                                                            , _newRow.RowNum + (_rangeAddr.LastRow - _rangeAddr.FirstRow)
                                                            , _rangeAddr.FirstColumn
                                                            , _rangeAddr.LastColumn);

                    this._sheet.AddMergedRegion(_newCellRangeAddr);
                }
            }


            //移除來源資料列的儲存格合併格式
            bool _reCheck;

            _reCheck = true;

            while (_reCheck == true)
            {
                _reCheck = false;

                for (int index = 0; index < this._sheet.NumMergedRegions - 1; index++)
                {
                    var _cellRangeAddress = this._sheet.GetMergedRegion(index);

                    if (_cellRangeAddress.FirstRow == _sourceRow.RowNum)
                    {
                        this._sheet.RemoveMergedRegion(index);
                        _reCheck = true;
                        break;
                    }
                }
            }


            //設定新增資料列的高度
            _newRow.Height = _sourceRow.Height;

            //移除來源資料列
            this._sheet.RemoveRow(_sourceRow);
        }
    }
}
