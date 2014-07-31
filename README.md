ExcelDataReader
===============

Lightweight and fast library written in C# for reading Microsoft Excel files

This project has just migrated from CodePlex - as is.
Please feel free to fork and submit pull requests.

**Note**
Please try the latest source from the repo before reporting issues as there have been recent changes.
Also, if you are reporting an issue it is really useful if you can supply an example excel file as this makes debugging much easier and without it we may not be able to resolve any problems.

## How to use
### C# code :
```c#
FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

//1. Reading from a binary Excel file ('97-2003 format; *.xls)
IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
//...
//2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
//...
//3. DataSet - The result of each spreadsheet will be created in the result.Tables
DataSet result = excelReader.AsDataSet();
//...
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

'1. Reading from a binary Excel file ('97-2003 format; *.xls)
Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateBinaryReader(stream)
'...
'2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)
'...
'3. DataSet - The result of each spreadsheet will be created in the result.Tables
Dim result As DataSet = excelReader.AsDataSet()
'...
'4. DataSet - Create column names from first row
excelReader.IsFirstRowAsColumnNames = True
Dim result As DataSet = excelReader.AsDataSet()

'5. Data Reader methods
While excelReader.Read()
    'excelReader.GetInt32(0);
End While

'6. Free resources (IExcelDataReader is IDisposable)
excelReader.Close()
```

### Tips
* SQL reporting services. Set ReadOption.Loose in the CreateBinaryReader factory method to skip some bounds checking which was causing SSRS generated xls to fail. (Only on changeset >= 82970)
