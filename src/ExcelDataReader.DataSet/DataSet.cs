#if NETSTANDARD1_2

using System;
using System.Collections;
using System.Collections.Generic;

namespace Excel {

	public class DataRowCollection : IEnumerable<DataRow> {
		List<DataRow> rows = new List<DataRow>();

		public void Add(params object[] values) {
			rows.Add(new DataRow(values));
		}

		public void Add(DataRow row) {
			rows.Add(row);
		}

		public void Clear() {
			rows.Clear();
		}

		public IEnumerator<DataRow> GetEnumerator() {
			return rows.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator() {
			return rows.GetEnumerator();
		}

		public DataRow this[int index]
		{
			get { return rows[index]; }
		}

		public int Count
		{
			get
			{
				return rows.Count;
			}
		}
	}

	public class DataRow {
		public object[] ItemArray { get; }
		public object this[int index]
		{
			get {
				var result = ItemArray[index];
				if (result == null)
					return DBNull.Value;
				return result;
			}
			set {
				ItemArray[index] = value;
			}
		}

		internal DataRow(object[] itemArray) {
			ItemArray = itemArray;
		}

		public bool IsNull(int i) {
			return ItemArray[i] == null;
		}
	}

	public class DataTableCollection : IEnumerable<DataTable> {
		List<DataTable> tables = new List<DataTable>();

		public void Add(DataTable table) {
			tables.Add(table);
		}

		public DataTable this[int index] {
			get { return tables[index]; }
		}

		public DataTable this[string name] {
			get {
				foreach (var table in tables)
					if (table.TableName == name)
						return table;
				return null;
			}
		}

		public int Count {
			get {
				return tables.Count;
			}
		}

		public void AddRange(DataTable[] range) {
			tables.AddRange(range);
		}

		public void Clear() {
			tables.Clear();
		}

		public IEnumerator<DataTable> GetEnumerator() {
			return tables.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator() {
			return tables.GetEnumerator();
		}
	}

	public class DataColumn {
		public string ColumnName { get; set; }
		public Type DataType { get; set; }
		public string Caption { get; set; }
		public DataColumn(string key, Type type) {
			ColumnName = key;
			DataType = type;
		}
	}

	public class DataColumnCollection : IEnumerable<DataColumn> {
		List<DataColumn> columns = new List<DataColumn>();

		public void Add(string key, Type type) {
			Add(new DataColumn(key, type));
		}

		public void Add(DataColumn column) {
			if (column.ColumnName != null && this[column.ColumnName] != null) {
				throw new ArgumentException("Duplicate column name: " + column.ColumnName);
			}
			columns.Add(column);
		}

		public IEnumerator<DataColumn> GetEnumerator() {
			return columns.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator() {
			return columns.GetEnumerator();
		}

		public DataColumn this[int index]
		{
			get { return columns[index]; }
		}

		public DataColumn this[string name]
		{
			get {
				foreach (var column in columns)
					if (column.ColumnName == name)
						return column;
				return null;
			}
		}

		public int Count
		{
			get
			{
				return columns.Count;
			}
		}
	}

	public class PropertyCollection : IEnumerable<KeyValuePair<string, string> > {
		Dictionary<string, string> dict = new Dictionary<string, string>();

		public string this[string key] {
			get { return dict[key]; }
		}

		public void Add(string key, string value) {
			dict.Add(key, value);
		}

		public IEnumerator<KeyValuePair<string, string>> GetEnumerator() {
			return dict.GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator() {
			return dict.GetEnumerator();
		}
	}

	public partial class DataTable {
		public string TableName { get; set; }
		public DataRowCollection Rows { get; set; }
		public DataColumnCollection Columns { get; set; }
		public PropertyCollection ExtendedProperties { get; set; }

		public DataTable() {
			Rows = new DataRowCollection();
			Columns = new DataColumnCollection();
			ExtendedProperties = new PropertyCollection();
		}

		public DataTable(string name) {
			TableName = name;
			Rows = new DataRowCollection();
			Columns = new DataColumnCollection();
			ExtendedProperties = new PropertyCollection();
		}

		public DataRow NewRow() {
			var itemArray = new object[Columns.Count];
			return new DataRow(itemArray);
		}

		public void ImportRow(DataRow row) {
			var result = NewRow();
			for (var i = 0; i < row.ItemArray.Length; i++) {
				result.ItemArray[i] = row.ItemArray[i];
			}
			Rows.Add(row);
		}

		public void BeginLoadData() { }
		public void EndLoadData() { }
		public DataTable Clone() {
			var result = new DataTable(TableName);
			foreach (var property in ExtendedProperties) {
				result.ExtendedProperties.Add(property.Key, property.Value);
			}

			foreach (var column in Columns) {
				result.Columns.Add(new DataColumn(column.ColumnName, column.DataType));
			}
			return result;
		}
	}

	public class DataSet {
		public DataTableCollection Tables { get; set; }

		public DataSet() {
			Tables = new DataTableCollection();
		}

		public void AcceptChanges() { }
	}
}

#endif
