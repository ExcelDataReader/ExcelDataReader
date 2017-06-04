namespace ExcelDataReader
{
    public class DataTable
    {
        public DataTable()
        {
            Rows = new DataRowCollection();
            Columns = new DataColumnCollection();
            ExtendedProperties = new PropertyCollection();
        }

        public DataTable(string name)
        {
            TableName = name;
            Rows = new DataRowCollection();
            Columns = new DataColumnCollection();
            ExtendedProperties = new PropertyCollection();
        }

        public string TableName { get; set; }

        public DataRowCollection Rows { get; set; }

        public DataColumnCollection Columns { get; set; }

        public PropertyCollection ExtendedProperties { get; set; }

        public DataRow NewRow()
        {
            var itemArray = new object[Columns.Count];
            return new DataRow(itemArray);
        }

        public void ImportRow(DataRow row)
        {
            var result = NewRow();
            for (int i = 0; i < row.ItemArray.Length; i++)
            {
                result.ItemArray[i] = row.ItemArray[i];
            }

            Rows.Add(row);
        }

        public void BeginLoadData()
        {
        }

        public void EndLoadData()
        {
        }

        public DataTable Clone()
        {
            var result = new DataTable(TableName);
            foreach (var property in ExtendedProperties)
            {
                result.ExtendedProperties.Add(property.Key, property.Value);
            }

            foreach (var column in Columns)
            {
                result.Columns.Add(new DataColumn(column.ColumnName, column.DataType));
            }

            return result;
        }
    }
}
