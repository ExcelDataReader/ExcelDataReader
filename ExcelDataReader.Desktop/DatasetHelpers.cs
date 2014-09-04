using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelDataReader.Desktop
{
	/// <summary>
	/// Helpers class
	/// </summary>
	internal static class DatasetHelpers
	{
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

        //public static void AddColumn(DataTable table, string columnName)
        //{
        //    //if a colum  already exists with the name append _i to the duplicates
        //    var adjustedColumnName = columnName;
        //    var column = table.Columns[columnName];
        //    var i = 1;
        //    while (column != null)
        //    {
        //        adjustedColumnName = string.Format("{0}_{1}", columnName, i);
        //        column = table.Columns[adjustedColumnName];
        //        i++;
        //    }

        //    table.Columns.Add(adjustedColumnName, typeof(Object));
        //}


    }
}