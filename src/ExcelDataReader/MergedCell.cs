using ExcelDataReader.Core;
using System;

namespace ExcelDataReader
{
    /// <summary>
    /// Header and footer text. 
    /// </summary>
    internal sealed class MergedCell
    {
        private string FromCell { get;  set; }
        private string ToCell { get;  set; }

        public int FromColumn { get; private set; }
        public int FromRow { get; private set; }

        public int ToColumn { get; private set; }
        public int ToRow { get; private set; }

        private Cell SourceCell { get; set; }

        internal MergedCell(string from, string to)
        {
            FromCell = from;
            ToCell = to;

            int fromColumn, fromRow, toColumn, toRow;
            ReferenceHelper.ParseReference(FromCell, out  fromColumn, out  fromRow);
            //0 indexed vs 1 indexed
            FromColumn = fromColumn - 1;
            FromRow = fromRow - 1;

            ReferenceHelper.ParseReference(ToCell, out  toColumn, out toRow);
            //0 indexed vs 1 indexed
            ToColumn = toColumn - 1;
            ToRow = toRow - 1;
        }

        internal MergedCell(int fromColumn,int fromRow, int toColumn, int toRow)
        {
            FromColumn = fromColumn ;
            FromRow = fromRow ;
            ToColumn = toColumn ;
            ToRow = toRow ;
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="col"></param>
        /// <param name="row"></param>
        /// <param name="cell"></param>
        /// <param name="merged"></param>
        /// <returns></returns>
        public bool GetSourceValue(int col, int row, Cell cell, out Cell merged)
        {
            merged = null;
            if (col == FromColumn && row == FromRow)
            {
                SourceCell = cell;
                return false;
            }

            if(SourceCell != null && col >= FromColumn && col <= ToColumn 
                && row >= FromRow && row <= ToRow)
            {
                merged = new Cell()
                {
                    ColumnIndex = col,
                    NumberFormatIndex = SourceCell.NumberFormatIndex,
                    Value = SourceCell.Value

                };
                
                return true;
            }

            return false;
        }
        
    }
}