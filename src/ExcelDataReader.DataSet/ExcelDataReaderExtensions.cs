using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelDataReader
{
    /// <summary>
    /// ExcelDataReader DataSet extensions
    /// </summary>
    public static class ExcelDataReaderExtensions
    {
        /// <summary>
        /// Converts all sheets to a DataSet
        /// </summary>
        /// <param name="self">The IExcelDataReader instance</param>
        /// <param name="configuration">An optional configuration object to modify the behavior of the conversion</param>
        /// <returns>A dataset with all workbook contents</returns>
        public static DataSet AsDataSet(this IExcelDataReader self, ExcelDataSetConfiguration configuration = null)
        {
            if (configuration == null)
            {
                configuration = new ExcelDataSetConfiguration();
            }

            self.Reset();

            var tableIndex = -1;
            var result = new DataSet();
            do
            {
                tableIndex++;
                if (configuration.FilterSheet != null && !configuration.FilterSheet(self, tableIndex))
                {
                    continue;
                }

                var tableConfiguration = configuration.ConfigureDataTable != null
                    ? configuration.ConfigureDataTable(self)
                    : null;

                if (tableConfiguration == null)
                {
                    tableConfiguration = new ExcelDataTableConfiguration();
                }

                var table = AsDataTable(self, tableConfiguration);
                result.Tables.Add(table);
            }
            while (self.NextResult());

            result.AcceptChanges();

            if (configuration.UseColumnDataType)
            {
                FixDataTypes(result);
            }

            self.Reset();

            return result;
        }

        private static string GetUniqueColumnName(DataTable table, string name)
        {
            var columnName = name;
            var i = 1;
            while (table.Columns[columnName] != null)
            {
                columnName = string.Format("{0}_{1}", name, i);
                i++;
            }

            return columnName;
        }

        private static DataTable AsDataTable(IExcelDataReader self, ExcelDataTableConfiguration configuration)
        {
            var result = new DataTable { TableName = self.Name };
            result.ExtendedProperties.Add("visiblestate", self.VisibleState);
            var first = true;
            var emptyRows = 0;
            var columnIndices = new List<int>();
            while (self.Read())
            {
                if (first)
                {
                    if (configuration.UseHeaderRow && configuration.ReadHeaderRow != null)
                    {
                        configuration.ReadHeaderRow(self);
                    }

                    for (var i = 0; i < self.FieldCount; i++)
                    {
                        if (configuration.FilterColumn != null && !configuration.FilterColumn(self, i))
                        {
                            continue;
                        }

                        var name = configuration.UseHeaderRow
                            ? Convert.ToString(self.GetValue(i))
                            : null;

                        if (string.IsNullOrEmpty(name))
                        {
                            name = configuration.EmptyColumnNamePrefix + i;
                        }

                        // if a column already exists with the name append _i to the duplicates
                        var columnName = GetUniqueColumnName(result, name);
                        var column = new DataColumn(columnName, typeof(object)) { Caption = name };
                        result.Columns.Add(column);
                        columnIndices.Add(i);
                    }

                    result.BeginLoadData();
                    first = false;

                    if (configuration.UseHeaderRow)
                    {
                        continue;
                    }
                }

                if (configuration.FilterRow != null && !configuration.FilterRow(self))
                {
                    continue;
                }

                if (IsEmptyRow(self))
                {
                    emptyRows++;
                    continue;
                }

                for (var i = 0; i < emptyRows; i++)
                {
                    result.Rows.Add(result.NewRow());
                }

                emptyRows = 0;

                var row = result.NewRow();

                for (var i = 0; i < columnIndices.Count; i++)
                {
                    var columnIndex = columnIndices[i];
                    var value = self.GetValue(columnIndex);
                    row[i] = value;
                }
                
                result.Rows.Add(row);
            }

            result.EndLoadData();
            return result;
        }

        private static bool IsEmptyRow(IExcelDataReader reader)
        {
            for (var i = 0; i < reader.FieldCount; i++)
            {
                if (reader.GetValue(i) != null)
                    return false;
            }

            return true;
        }

        private static void FixDataTypes(DataSet dataset)
        {
            var tables = new List<DataTable>(dataset.Tables.Count);
            bool convert = false;
            foreach (DataTable table in dataset.Tables)
            {
                if (table.Rows.Count == 0)
                {
                    tables.Add(table);
                    continue;
                }

                DataTable newTable = null;
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    Type type = null;
                    foreach (DataRow row in table.Rows)
                    {
                        if (row.IsNull(i))
                            continue;
                        var curType = row[i].GetType();
                        if (curType != type)
                        {
                            if (type == null)
                            {
                                type = curType;
                            }
                            else
                            {
                                type = null;
                                break;
                            }
                        }
                    }

                    if (type == null)
                        continue;
                    convert = true;
                    if (newTable == null)
                        newTable = table.Clone();
                    newTable.Columns[i].DataType = type;
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
                else
                {
                    tables.Add(table);
                }
            }

            if (convert)
            {
                dataset.Tables.Clear();
                dataset.Tables.AddRange(tables.ToArray());
            }
        }
    }
}
