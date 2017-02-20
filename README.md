ExcelDataReader
===============

Lightweight and fast library written in C# for reading Microsoft Excel files (added methods for batch reading, row skipping, get schema, get top rows, get sheet names)

This project has just migrated from CodePlex - as is.
Please feel free to fork and submit pull requests.

**Note**
Please try the latest source from the repo before reporting issues as there have been recent changes.
Also, if you are reporting an issue it is really useful if you can supply an example excel file as this makes debugging much easier and without it we may not be able to resolve any problems.

## Finding the binaries
It is recommended to use Nuget 
```
Install-Package ExcelDataReader
```
The current binaries are still on the codeplex site, but these will not be updated going forward. If there are enough requests for separate binary hosting other than nuget then we'll come up with some other solution.

## How to use
### C# code :
```c#
FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

//1. Reading from a binary Excel file ('97-2003 format; *.xls)
IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

//2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

//3. DataSet - The result of each spreadsheet will be created in the result.Tables
DataSet result = excelReader.AsDataSet();

//4. DataSet - Create column names from first row
excelReader.IsFirstRowAsColumnNames = true;
DataSet result = excelReader.AsDataSet();

//5. Data Reader methods
while (excelReader.Read())
{
	//excelReader.GetInt32(0);
}

//6. Free resources (IExcelDataReader is IDisposable)
excelReader.Close();
```

### VB.NET Code:

```vb.net

Dim stream As FileStream = File.Open(filePath, FileMode.Open, FileAccess.Read)

1. Reading from a binary Excel file ('97-2003 format; *.xls)
Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateBinaryReader(stream)
...
2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
...
3. DataSet - The result of each spreadsheet will be created in the result.Tables
Dim result As DataSet = excelReader.AsDataSet()
...
4. DataSet - Create column names from first row
excelReader.IsFirstRowAsColumnNames = True
Dim result As DataSet = excelReader.AsDataSet()

5. Data Reader methods
While excelReader.Read()
	'excelReader.GetInt32(0);
End While

6. Free resources (IExcelDataReader is IDisposable)
excelReader.Close()
```

### Tips
* SQL reporting services. Set ReadOption.Loose in the CreateBinaryReader factory method to skip some bounds checking which was causing SSRS generated xls to fail. (Only on changeset >= 82970)

===============

# ExcelDataReader - BatchRead
Added excel file batch reading and row skipping support. Implemented GetSchema method. Added utility methods such as GetTopRows and GetSheetNames.

## How to use
### C# code :
```c#

1. Reading sheet names
	List<string> sheetNames = null;
	using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestMultiSheet")))
	{
		if (!excelReader.IsValid) { throw new Exception(excelReader.ExceptionMessage); }
		sheetNames = excelReader.GetSheetNames();
	}
2. Reading top rows
	DataTable dataTable = null;
    using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestMultiSheet")))
    {
        if (!excelReader.IsValid) { throw new Exception(excelReader.ExceptionMessage); }
        dataTable = excelReader.GetTopRows(5, new SheetParameters("Sheet2", false));
    }
3. Reading schema (sheetNames, columnNames and dataTypes)
   Note: This forms schema based on first 1000 rows, to increase sample size for inferring schema provide excelReader.BatchSize of desired rows. 
	DataSet dataset = null;
	using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestMultiSheet")))
	{
		if (!excelReader.IsValid) { throw new Exception(excelReader.ExceptionMessage); }
		dataset = excelReader.GetSchema();
	}
4. Read single sheet in batch
	DataTable dataTable = null;
	using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestMultiSheet")))
	{
		if (!excelReader.IsValid) { throw new Exception(excelReader.ExceptionMessage); }
		// Read one sheet of excel in batch
		excelReader.SheetName = "Sheet2";
		excelReader.IsFirstRowAsColumnNames = false; // default is true
		excelReader.SkipRows = 0; // default is 0
		excelReader.BatchSize = 10000; // modify as per need, default is 1000
		while (excelReader.ReadBatch())
		{
			dataTable = excelReader.GetCurrentBatch();
			// process batch rows
		}
	}
5. Read all sheets of excel in batch
	DataSet dataSet = null;
	DataTable dataTable = null;
	using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(Helper.GetTestWorkbook("xTestMultiSheet")))
	{
		if (!excelReader.IsValid) { throw new Exception(excelReader.ExceptionMessage); }
		dataSet = excelReader.GetSchema();
		excelReader.BatchSize = 10000; // modify as per need, default is 1000
		foreach (DataTable dt in dataSet.Tables)
		{
			excelReader.SheetName = dt.TableName;
			excelReader.IsFirstRowAsColumnNames = Convert.ToBoolean(dt.ExtendedProperties["IsFirstRowAsColumnNames"]);
			excelReader.SkipRows = Convert.ToInt32(dt.ExtendedProperties["SkipRows"]);
			while (excelReader.ReadBatch())
			{
				dataTable = excelReader.GetCurrentBatch();
				// process batch rows
			}
		}
	}