namespace ExcelDataReader.Tests;

internal static class Configuration
{
    public static ExcelDataSetConfiguration NoColumnNamesConfiguration { get; } = new()
    {
        ConfigureDataTable = reader => new()
        {
            UseHeaderRow = false
        }
    };

    public static ExcelDataSetConfiguration FirstRowColumnNamesConfiguration { get; } = new()
    {
        ConfigureDataTable = reader => new()
        {
            UseHeaderRow = true
        }
    };

    public static ExcelDataSetConfiguration FirstRowColumnNamesPrefixConfiguration { get; } = new()
    {
        ConfigureDataTable = reader => new()
        {
            UseHeaderRow = true,
            EmptyColumnNamePrefix = "Prefix"
        }
    };

    public static Stream GetTestWorkbook(string key) => TestWorkbook.Open(key);

    public static string GetTestWorkbookPath(string key) => TestWorkbook.GetPath(key);
}
