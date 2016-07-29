using System;
using System.Data;
using System.Collections.Generic;
using ExcelDataReader.Data;

namespace ExcelDataReader.Dataset
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
            FixDataTypes(workbookData);
        }

        public void AddExtendedPropertyToTable(string propertyName, string propertyValue)
        {
            currentTable.ExtendedProperties.Add(propertyName, propertyValue);
        }
		
        internal static void FixDataTypes(DataSet dataset)
        {
            var tables = new List<DataTable>(dataset.Tables.Count);
            bool convert = false;
            foreach (DataTable table in dataset.Tables)
            {
               
                if ( table.Rows.Count == 0)
                {
                    tables.Add(table);
                    continue;
                }
                DataTable newTable = null;
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    Type type = null;
                    foreach (DataRow row  in table.Rows)
                    {
                        if (row.IsNull(i))
                            continue;
                        var curType = row[i].GetType();
                        if (curType != type)
                        {
                            if (type == null)
                                type = curType;
                            else
                            {
                                type = null;
                                break;
                            }
                        }
                    }
                    if (type != null)
                    {
                        convert = true;
                        if (newTable == null)
                            newTable = table.Clone();
                        newTable.Columns[i].DataType = type;

                    }
                }
                if (newTable != null)
                {
                    newTable.BeginLoadData();
                    foreach (DataRow row in table.Rows)
                    {
                        newTable.ImportRow(row);
                    }

                    newTable.EndLoadData();
                    tables.Add(newTable);

                }
                else tables.Add(table);
            }
            if (convert)
            {
                dataset.Tables.Clear();
                dataset.Tables.AddRange(tables.ToArray());
            }
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
