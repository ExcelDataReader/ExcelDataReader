namespace ExcelDataReader.Tests;

internal static class TestWorkbook
{
    public static Stream Open(string key)
    {
        var fileName = GetPath(key);
        return new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
    }

    public static string GetPath(string key)
    {
        var directory = TestContext.CurrentContext.TestDirectory;
        while (directory != null && !File.Exists(Path.Combine(directory, "ExcelDataReader.slnx")))
            directory = Path.GetDirectoryName(directory);

        var resources = Path.Combine(directory, "src", "TestData");
        var path = Path.GetFullPath(Path.Combine(resources, key));
        Assert.That(path, Does.Exist, $"File not found: '{path}'.");
        return path;
    }
}
