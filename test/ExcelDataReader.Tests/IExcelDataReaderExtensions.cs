using System;
using System.Data;
using System.Collections.Generic;
using Excel;
using ExcelDataReader.Dataset;
using System.Collections;

#if NET20
namespace System.Runtime.CompilerServices {
	/// <summary>
	/// 
	/// </summary>
	[AttributeUsage(AttributeTargets.Assembly | AttributeTargets.Class | AttributeTargets.Method, AllowMultiple = false, Inherited = false)]
	public class ExtensionAttribute : Attribute {
	}
}
#endif

namespace Excel
{
    public static class IExcelDataReaderExtensions {
		public static DataSet AsDataSet(this IExcelDataReader self) {
			return self.AsDataSet(false);
		}

		public static DataSet AsDataSet(this IExcelDataReader self, bool convertOADateTime) {
			var h = new DatasetHelper();
			self.LoadDataSet(h, convertOADateTime);
			return (DataSet)h.Dataset;
		}

		static DataTable AsDataTable(IExcelDataReader self, bool yeah) {
			var result = new DataTable();
			result.TableName = self.Name;
			bool first = true;
			while (self.Read()) {
				if (first) {
					for (var i = 0; i < self.FieldCount; i++) {
						var name = self.GetName(i);
						var type = self.GetFieldType(i);
						result.Columns.Add(name, type);
					}
				}

				var row = result.NewRow();

				for (var i = 0; i < self.FieldCount; i++) {
					var name = self.GetName(i);
					row.ItemArray[i] = self.GetValue(i);
				}

				result.Rows.Add(row);
				
			}
			return result;
		}
	}
}
