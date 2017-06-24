﻿using System;
using System.Collections.Generic;
using System.Data;

namespace ExcelDataReader
{
    public static class ExcelDataReaderExtensions
    {
        public static DataSet AsDataSet(this IExcelDataReader self, ExcelDataSetConfiguration configuration = null)
        {
            if (configuration == null)
            {
                configuration = new ExcelDataSetConfiguration();
            }

            self.Reset();

            var result = new DataSet();
            do
            {
                var tableConfiguration = configuration.ConfigureDataTable != null
                    ? configuration.ConfigureDataTable(self)
                    : new ExcelDataTableConfiguration();
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
            bool first = true;
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
                        var emptyColumnName = configuration.EmptyColumnNamePrefix + i;
                        var name = configuration.UseHeaderRow
                            ? self.GetString(i) ?? emptyColumnName
                            : emptyColumnName;

                        // if a column already exists with the name append _i to the duplicates
                        var columnName = GetUniqueColumnName(result, name);
                        var column = new DataColumn(columnName, typeof(object)) { Caption = name };
                        result.Columns.Add(column);
                    }

                    result.BeginLoadData();
                    first = false;

                    if (configuration.UseHeaderRow)
                    {
                        continue;
                    }
                }

                var row = result.NewRow();

                for (var i = 0; i < self.FieldCount; i++)
                {
                    // var name = self.GetName(i);
                    var value = self.GetValue(i);
                    row[i] = value;
                }
                
                result.Rows.Add(row);
            }

            result.EndLoadData();
            return result;
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
