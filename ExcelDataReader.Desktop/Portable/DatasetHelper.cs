using System;
using System.Data;
using ExcelDataReader.Portable.Data;

namespace ExcelDataReader.Desktop.Portable
{
    public class DatasetHelper : IDatasetHelper
    {
        private DataSet workbookData;
        private DataTable currentTable;
        public bool IsValid { get; set; }

        public object Dataset
        {
            get { return workbookData; }
        }

        public void CreateNew()
        {
            workbookData = new DataSet();
        }

        public void CreateNewTable(string name)
        {
            currentTable = new DataTable(name);
        }

        public void EndLoadTable()
        {
            workbookData.Tables.Add(currentTable);
        }

        public void AddColumn(string columnName)
        {
            if (columnName == null)
            {
                currentTable.Columns.Add(null, typeof (Object));
                return;
            }

            //if a colum  already exists with the name append _i to the duplicates
            var adjustedColumnName = columnName;
            var column = currentTable.Columns[columnName];
            var i = 1;
            while (column != null)
            {
                adjustedColumnName = string.Format("{0}_{1}", columnName, i);
                column = currentTable.Columns[adjustedColumnName];
                i++;
            }

            currentTable.Columns.Add(adjustedColumnName, typeof(Object));
        }

        public void BeginLoadData()
        {
            currentTable.BeginLoadData();
        }

        public void AddRow(params object[] values)
        {
            currentTable.Rows.Add(values);
        }

        public void DatasetLoadComplete()
        {
            workbookData.AcceptChanges();
            DatasetHelpers.FixDataTypes(workbookData);
        }
    }

    //public class TableHelper : ITableHelper
    //{
    //    private TableHelper table;

    //    public void CreateNew(string sheetName)
    //    {
    //        throw new NotImplementedException();
    //    }
    //}

    //public class DatasetFactory : IDatasetHelperFactory
    //{
    //    public IDatasetHelper CreateDataset()
    //    {
    //        return new DatasetHelper();
    //    }

    //    public ITableHelper CreateTable()
    //    {
    //        return new TableHelper();
    //    }
    //}
}
