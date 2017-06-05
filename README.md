ExcelDataReader
===============

Lightweight and fast library written in C# for reading Microsoft Excel files.

Please feel free to fork and submit pull requests.

If you are reporting an issue it is really useful if you can supply an example Excel file as this makes debugging much easier and without it we may not be able to resolve any problems.

This project is using a git-flow style workflow so please submit pull requests to the develop branch if possible.

[![Build status](https://ci.appveyor.com/api/projects/status/ii6hbs9otpbg1nqh/branch/master?svg=true)](https://ci.appveyor.com/project/andersnm/exceldatareader/branch/master) [![Build status](https://ci.appveyor.com/api/projects/status/ii6hbs9otpbg1nqh/branch/develop?svg=true)](https://ci.appveyor.com/project/andersnm/exceldatareader/branch/develop)

## Finding the binaries
It is recommended to use Nuget. F.ex through the VS Package Manager Console `Install-Package <package>` or using the VS "Manage NuGet Packages..." extension. 

As of ExcelDataReader version 3.0, the project was split into multiple packages:

Install the `ExcelDataReader` base package to use the "low level" reader interface. Compatible with net20, net45 and netstandard 1.3.

Install the `ExcelDataReader.DataSet` extension package to use the AsDataSet() method and load spreadsheets into System.Data.DataSet. This will also pull in the base package. Compatible with net20 and net45.


## How to use
### C# code :
```c#
FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

// Choose one of either 1 or 2
// 1. Reading from a binary Excel file ('97-2003 format; *.xls)
IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

// 2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

// Choose one of either 3, 4, or 5
// 3. DataSet - The result of each spreadsheet will be created in the result.Tables
DataSet result = excelReader.AsDataSet();

// 4. DataSet - Create column names from first row
excelReader.IsFirstRowAsColumnNames = true;
DataSet result = excelReader.AsDataSet();

// 5. Data Reader methods
while (excelReader.Read()) {
	// excelReader.GetInt32(0);
}

// 6. Free resources (IExcelDataReader is IDisposable)
excelReader.Close();
```

### VB.NET Code:

```vb.net

Dim stream As FileStream = File.Open(filePath, FileMode.Open, FileAccess.Read)

' 1. Reading from a binary Excel file ('97-2003 format; *.xls)
Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateBinaryReader(stream)

' 2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
Dim excelReader As IExcelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream)

' 3. DataSet - The result of each spreadsheet will be created in the result.Tables
Dim result As DataSet = excelReader.AsDataSet()

' 4. DataSet - Create column names from first row
excelReader.IsFirstRowAsColumnNames = True
Dim result As DataSet = excelReader.AsDataSet()

' 5. Data Reader methods
While excelReader.Read()
	' excelReader.GetInt32(0);
End While

' 6. Free resources (IExcelDataReader is IDisposable)
excelReader.Close()
```

### Tips
* SQL reporting services. Set ReadOption.Loose in the CreateBinaryReader factory method to skip some bounds checking which was causing SSRS generated xls to fail. (Only on changeset >= 82970)
