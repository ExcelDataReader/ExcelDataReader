using System.Collections;
using System.Collections.Generic;
using ExcelDataReader.Silverlight.Data;

namespace ExcelDataReader.Silverlight.Data
{
    public class ExcelWorkBookFactory : IWorkBookFactory
    {
        public IWorkBook CreateWorkBook()
        {
            return new ExcelWorkBook();
        }
    }

    public class ExcelWorkBook : IWorkBook
    {
        internal ExcelWorkBook()
        {
            WorkSheets = new ExcelWorkSheetCollection();
        }

        public IWorkSheetCollection WorkSheets { get; private set; }

        public IWorkSheet CreateWorkSheet()
        {
            return new ExcelWorkSheet();
        }

        public string DataSetName { get; set; }
    }

    public class ExcelWorkSheetCollection : List<IWorkSheet>, IWorkSheetCollection
    {
        internal ExcelWorkSheetCollection() { }
    }

    public class ExcelWorkSheet : IWorkSheet
    {
        private readonly ExcelDataColumnCollection _Columns;
        private readonly ExcelDataRowCollection _Rows;

        internal ExcelWorkSheet()
        {
            _Columns = new ExcelDataColumnCollection();
            _Rows = new ExcelDataRowCollection();
        }

        public string Name { get; set; }

        public IDataColumnCollection Columns
        {
            get { return _Columns; }
        }

        public IDataColumn CreateDataColumn()
        {
            return new ExcelDataColumn();
        }

        public IDataRowCollection Rows
        {
            get { return _Rows; }
        }

        public IDataRow CreateDataRow()
        {
            return new ExcelDataRow();
        }
    }

    public class ExcelDataColumn : IDataColumn
    {
        internal ExcelDataColumn() { }

        public string ColumnName { get; set; }
    }

    public class ExcelDataRow : IDataRow
    {
        internal ExcelDataRow() { }

        public IList Values { get; set; }
    }

    public class ExcelDataColumnCollection : List<IDataColumn>, IDataColumnCollection
    {
        internal ExcelDataColumnCollection() { }
    }

    public class ExcelDataRowCollection : List<IDataRow>, IDataRowCollection
    {
        internal ExcelDataRowCollection() { }
    }
}