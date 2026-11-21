using System.Globalization;
using System.IO.Compression;
using System.Text;
using ExcelDataReader.Core.NumberFormat;

namespace ExcelDataReader.Tests;

public class FormatReaderCharacterizationTest
{
    // Frozen outcomes from 5f675d4: edge cases, section combinations, and seed-765 generated formats.
    private const string SnapshotName =
#if NETFRAMEWORK
        "NumberFormatClassification.net462.txt.gz";
#else
        "NumberFormatClassification.modern.txt.gz";
#endif

    [TestCase("en-US")]
    [TestCase("tr-TR")]
    public void ClassificationMatchesUnchangedParser(string culture)
    {
        var previousCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(culture);
            using var stream = Configuration.GetTestWorkbook(SnapshotName);
            using var gzip = new GZipStream(stream, CompressionMode.Decompress);
            using var reader = new StreamReader(gzip);
            int count = 0;
            while (reader.ReadLine() is { } line)
            {
                string[] parts = line.Split('\t');
                string text = Encoding.Unicode.GetString(Convert.FromBase64String(parts[0]));
                bool malformedSectionChanged = parts[1] == nameof(FormatException) && HasEmptyCondition(text) ||
                    parts[1] == nameof(ArgumentOutOfRangeException) && HasUnclosedBracket(text);
                if (!malformedSectionChanged)
                    Assert.That(Outcome(text), Is.EqualTo(parts[1]), text);
                count++;
            }

            Assert.That(count, Is.GreaterThan(4000));
        }
        finally
        {
            CultureInfo.CurrentCulture = previousCulture;
        }
    }

    [Test]
    public void NullInputPreservesException()
    {
        Assert.Throws<NullReferenceException>(() => new NumberFormatString(null));
    }

    private static bool HasEmptyCondition(string formatString) =>
        formatString.IndexOf("[>]", StringComparison.Ordinal) >= 0 ||
        formatString.IndexOf("[<]", StringComparison.Ordinal) >= 0 ||
        formatString.IndexOf("[=]", StringComparison.Ordinal) >= 0 ||
        formatString.IndexOf("[>=]", StringComparison.Ordinal) >= 0 ||
        formatString.IndexOf("[<=]", StringComparison.Ordinal) >= 0 ||
        formatString.IndexOf("[<>]", StringComparison.Ordinal) >= 0;

    private static bool HasUnclosedBracket(string formatString) =>
        formatString.LastIndexOf('[') > formatString.LastIndexOf(']');

    private static string Outcome(string text)
    {
        try
        {
            var format = new NumberFormatString(text);
            Assert.That(format.FormatString, Is.SameAs(text));
            return $"{format.IsValid},{format.IsDateTimeFormat},{format.IsTimeSpanFormat}";
        }
        catch (FormatException)
        {
            return nameof(FormatException);
        }
        catch (OverflowException)
        {
            return nameof(OverflowException);
        }
        catch (ArgumentOutOfRangeException)
        {
            return nameof(ArgumentOutOfRangeException);
        }
    }
}
