namespace ExcelDataReader.Tests;

public class DecodedCellMaterializationTests
{
    [TestCase(false)]
    [TestCase(true)]
    public void ValueTagsDistinguishIntegersReferencesAndFormattedStrings(bool date1904)
    {
        Assert.That(Materialize(42, "Literal", 0, date1904), Is.TypeOf<int>().And.EqualTo(42));
        Assert.That(Materialize(42.5, "Literal", 0, date1904), Is.TypeOf<double>().And.EqualTo(42.5));
        Assert.That(Materialize(0U, "SharedString", 14, date1904), Is.Null);
        Assert.That(Materialize(uint.MaxValue, "SharedString", 0, date1904), Is.Null);
        Assert.That(Materialize(0, "Literal", 14, date1904), Is.EqualTo(date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 31)));
        Assert.That(Materialize(2, "Literal", 46, date1904), Is.EqualTo(TimeSpan.FromDays(2)));

        const string time = "02:30:00";
        Assert.That(Materialize(time, "Literal", 14, date1904), Is.SameAs(time));
        Assert.That(Materialize(time, "FormattedString", 14, date1904), Is.EqualTo(
            (date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 31)).AddHours(2.5)));
        Assert.That(Materialize("P2DT3H", "FormattedString", 46, date1904), Is.EqualTo(TimeSpan.FromHours(51)));
        Assert.That(Materialize("not a date", "FormattedString", 14, date1904), Is.EqualTo("not a date"));
        Assert.That(Materialize("not a duration", "FormattedString", 46, date1904), Is.EqualTo("not a duration"));
        Assert.That(Materialize("_x0041_", "Literal", 0, date1904), Is.EqualTo("_x0041_"));
        Assert.That(Materialize(null, "Literal", 14, date1904), Is.Null);
        Assert.That(Materialize(true, "Literal", 14, date1904), Is.True);
    }

    [TestCase("Issue368_Header.xls", false)]
    [TestCase("Issue368_Ixfe.xls", false)]
    public void LegacyIntegersAndInlineFormatsSurviveReset(string fixture, bool singlePass)
    {
        using var reader = Open(fixture, singlePass);
        Assert.That(reader.Read(), Is.True);
        reader.Reset();
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetValue(2), Is.TypeOf<int>().And.EqualTo(1234));
            Assert.That(reader.GetNumberFormatString(2), Is.EqualTo("00.0"));
            Assert.That(reader.Read(), Is.True);
            Assert.That(reader.GetValue(2), Is.EqualTo(4321));
            Assert.That(reader.GetFieldType(2), Is.EqualTo(fixture == "Issue368_Header.xls" ? typeof(double) : typeof(int)));
            Assert.That(reader.GetNumberFormatString(2), Is.EqualTo("0000.00"));
            while (reader.Read())
            {
            }

            reader.Reset();
        }
    }

    [TestCase(false)]
    [TestCase(true)]
    public void SequentialBiff2IntegersKeepInlineNumberFormats(bool singlePass)
    {
        using var reader = Open("Issue368_Formats.xls", singlePass);
        Assert.That(reader.Read(), Is.True);
        reader.Reset();
        for (int pass = 0; pass < 2; pass++)
        {
            for (int row = 0; row < 42; row++)
            {
                Assert.That(reader.Read(), Is.True);
                Assert.That(reader.GetValue(0), Is.TypeOf<int>().And.EqualTo(row % 10));
                Assert.That(reader.GetNumberFormatString(0), Is.EqualTo("\"" + row + "\" 0.00"));
            }

            while (reader.Read())
            {
            }

            reader.Reset();
        }
    }

    [TestCase("Issue368_Header.xls")]
    [TestCase("Issue368_Ixfe.xls")]
    public void NonSequentialLegacyCellsStillRequireIndexedReading(string fixture)
    {
        using var reader = Open(fixture, true);
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            Assert.Throws<InvalidOperationException>(() =>
            {
                while (reader.Read())
                {
                }
            });
            reader.Reset();
        }
    }

    [TestCase("xls", false)]
    [TestCase("xls", true)]
    [TestCase("xlsb", false)]
    [TestCase("xlsb", true)]
    [TestCase("xlsx", false)]
    [TestCase("xlsx", true)]
    public void CachedFormulaErrorsRemainSeparateFromNullValues(string extension, bool singlePass)
    {
        using var reader = Open("Issue329_Error." + extension, singlePass);
        Assert.That(reader.Read(), Is.True);
        reader.Reset();
        for (int pass = 0; pass < 2; pass++)
        {
            Assert.That(reader.Read(), Is.True);
            for (int column = 0; column < 3; column++)
                Assert.That(reader.IsDBNull(column), Is.True);
            Assert.That(reader.GetCellError(0), Is.EqualTo(CellError.DIV0));
            Assert.That(reader.GetCellError(1), Is.EqualTo(CellError.NA));
            Assert.That(reader.GetCellError(2), Is.EqualTo(CellError.VALUE));
            while (reader.Read())
            {
            }

            reader.Reset();
        }
    }

    [TestCase("xls", false)]
    [TestCase("xls", true)]
    [TestCase("xlsb", false)]
    [TestCase("xlsb", true)]
    [TestCase("xlsx", false)]
    [TestCase("xlsx", true)]
    public void CachedFormulaStringsAndDatesSurviveReset(string extension, bool singlePass)
    {
        using var reader = Open("EncodingFormulaDate1520." + extension, singlePass);
        Assert.That(reader.Read(), Is.True);
        reader.Reset();
        for (int pass = 0; pass < 2; pass++)
        {
            for (int row = 0; row <= 8; row++)
            {
                Assert.That(reader.Read(), Is.True);
                if (row == 1)
                {
                    Assert.That(reader.GetString(0), Is.EqualTo("John test"));
                    Assert.That(reader.GetDateTime(1).Date, Is.EqualTo(new DateTime(2009, 5, 1)));
                }
                else if (row == 2)
                {
                    Assert.That(reader.GetString(0), Is.EqualTo("Simon Hodgetts"));
                    Assert.That(reader.GetDateTime(4).TimeOfDay, Is.EqualTo(TimeSpan.FromHours(11)));
                }
                else if (row is 7 or 8)
                {
                    Assert.That(reader.GetString(0), Is.EqualTo("librement réutilisable"));
                }
            }

            while (reader.Read())
            {
            }

            reader.Reset();
        }
    }

    private static IExcelDataReader Open(string fixture, bool singlePass) =>
        ExcelReaderFactory.CreateReader(Configuration.GetTestWorkbook(fixture), new ExcelReaderConfiguration { SinglePassMode = singlePass });

    private static object Materialize(object value, string kind, int formatIndex, bool date1904)
    {
        var assembly = typeof(ExcelReaderFactory).Assembly;
        var workbookType = assembly.GetType("ExcelDataReader.Core.CommonWorkbook")!;
        var valueType = assembly.GetType("ExcelDataReader.Core.DecodedCellValue")!;
        var kindType = assembly.GetType("ExcelDataReader.Core.CellValueKind")!;
        var styleType = assembly.GetType("ExcelDataReader.Core.ExtendedFormat")!;
        var style = Activator.CreateInstance(styleType, [formatIndex])!;
        var decoded = kind == "SharedString"
            ? valueType.GetMethod("SharedString")!.Invoke(null, [value])!
            : Activator.CreateInstance(valueType, [value, Enum.Parse(kindType, kind), 0U])!;
        var cell = workbookType.GetMethod("CreateCell")!.Invoke(
            Activator.CreateInstance(workbookType), [7, decoded, style, CellError.DIV0, date1904])!;
        var cellType = cell.GetType();
        Assert.That(cellType.GetProperty("ColumnIndex")!.GetValue(cell), Is.EqualTo(7));
        Assert.That(cellType.GetProperty("EffectiveStyle")!.GetValue(cell), Is.SameAs(style));
        Assert.That(cellType.GetProperty("Error")!.GetValue(cell), Is.EqualTo(CellError.DIV0));
        return cellType.GetProperty("Value")!.GetValue(cell);
    }
}
