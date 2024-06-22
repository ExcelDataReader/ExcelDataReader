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
            return new DataRow(Columns.Count);
        }

        public void ImportRow(DataRow row)
        {
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
