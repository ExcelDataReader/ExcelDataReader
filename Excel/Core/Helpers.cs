using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Text;

namespace ExcelDataReader.Portable.Core
{
	/// <summary>
	/// Helpers class
	/// </summary>
	internal static class Helpers
	{
#if CF_DEBUG || CF_RELEASE

		/// <summary>
		/// Determines whether [is single byte] [the specified encoding].
		/// </summary>
		/// <param name="encoding">The encoding.</param>
		/// <returns>
		/// 	<c>true</c> if [is single byte] [the specified encoding]; otherwise, <c>false</c>.
		/// </returns>
		public static bool IsSingleByteEncoding(Encoding encoding)
		{
			return encoding.GetChars(new byte[] { 0xc2, 0xb5 }).Length == 1;
		}
#else

		/// <summary>
		/// Determines whether [is single byte] [the specified encoding].
		/// </summary>
		/// <param name="encoding">The encoding.</param>
		/// <returns>
		/// 	<c>true</c> if [is single byte] [the specified encoding]; otherwise, <c>false</c>.
		/// </returns>
		public static bool IsSingleByteEncoding(Encoding encoding)
		{
			return encoding.IsSingleByte;
		}
#endif

		public static double Int64BitsToDouble(long value)
		{
			return BitConverter.ToDouble(BitConverter.GetBytes(value), 0);
		}

	    private static Regex re = new Regex("_x([0-9A-F]{4,4})_");

        public static string ConvertEscapeChars(string input)
        {
            return re.Replace(input, m => (((char)UInt32.Parse(m.Groups[1].Value, NumberStyles.HexNumber))).ToString());
        }

	    public static object ConvertFromOATime(double value)
	    {
	        if ((value >= 0.0) && (value < 60.0))
	        {
	            value++;
	        }
	        //if (date1904)
	        //{
	        //    Value += 1462.0;
	        //}
	        return DateTime.FromOADate(value);
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

		public static void AddColumnHandleDuplicate(DataTable table, string columnName)
		{
			//if a colum  already exists with the name append _i to the duplicates
			var adjustedColumnName = columnName;
			var column = table.Columns[columnName];
			var i = 1;
			while (column != null)
			{
				adjustedColumnName = string.Format("{0}_{1}", columnName, i);
				column = table.Columns[adjustedColumnName];
				i++;
			}

			table.Columns.Add(new DataColumn(adjustedColumnName, typeof(Object)) { Caption = columnName });
		}
    }
}