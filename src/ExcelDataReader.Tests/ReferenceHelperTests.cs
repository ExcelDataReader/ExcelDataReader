using ExcelDataReader.Core;

namespace ExcelDataReader.Tests;

public class ReferenceHelperTests
{
    [TestCase("A1", true, 1, 1)]
    [TestCase("a1", true, 1, 1)]
    [TestCase("\u017F1", false, 0, 0)]
    [TestCase("xFd1048576", true, 16384, 1048576)]
    [TestCase("A0", true, 1, 0)]
    [TestCase("A2147483647", true, 1, int.MaxValue)]
    [TestCase("A2147483648", false, 1, 0)]
    [TestCase("A1!", false, 1, 0)]
    [TestCase("A\u0661", false, 0, 0)]
    [TestCase("$A1", false, 0, 0)]
    [TestCase("A", false, 1, 0)]
    [TestCase("", false, 0, 0)]
    [TestCase(null, false, 0, 0)]
    public void ReferencesPreserveValidation(string reference, bool valid, int column, int row)
    {
        Assert.That(ReferenceHelper.ParseReference(reference, out int actualColumn, out int actualRow), Is.EqualTo(valid));
        Assert.That(actualColumn, Is.EqualTo(column));
        Assert.That(actualRow, Is.EqualTo(row));
    }

    [Test]
    public void ColumnLettersAcceptOnlyAsciiLetters()
    {
        for (int i = 0; i <= char.MaxValue; i++)
        {
            char character = (char)i;
            bool letter = character is (>= 'A' and <= 'Z') or (>= 'a' and <= 'z');
            bool actual = ReferenceHelper.ParseReference(character + "1", out int column, out int row);
            Assert.That(actual, Is.EqualTo(letter), $"U+{i:X4}");
            if (letter)
            {
                int expectedColumn = character is >= 'a' and <= 'z' ? character - 'a' + 1 : character - 'A' + 1;
                Assert.That(column, Is.EqualTo(expectedColumn));
                Assert.That(row, Is.EqualTo(1));
            }
        }
    }
}