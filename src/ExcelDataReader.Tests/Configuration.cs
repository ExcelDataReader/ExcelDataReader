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

    public static Stream GetTestWorkbook(string key)
    {
        var fileName = GetTestWorkbookPath(key);
        return new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    }

    public static string GetTestWorkbookPath(string key)
    {
        var directory = TestContext.CurrentContext.TestDirectory;
        while (directory != null && !File.Exists(Path.Combine(directory, "ExcelDataReader.sln")))
            directory = Path.GetDirectoryName(directory);

        var resources = Path.Combine(directory, "src/TestData");
        var path = Path.Combine(resources, key);
        path = Path.GetFullPath(path);
        Assert.That(path, Does.Exist, $"File not found: '{path}'.");
        return path;
    }
}
