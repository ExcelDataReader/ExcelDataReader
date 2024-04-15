using System;

using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    public static class AssertUtilities
    {
        public static void DoOpenOfficeTest(IExcelDataReader excelReader)
        {
            excelReader.Read();
            Assert.That(excelReader.FieldCount, Is.EqualTo(6));
            Assert.That(excelReader.GetString(0), Is.EqualTo("column a"));
            Assert.That(excelReader.GetString(1), Is.EqualTo(" column b"));
            Assert.That(excelReader.GetString(2), Is.EqualTo(" column b"));
            Assert.That(excelReader.GetString(3), Is.Null);
            Assert.That(excelReader.GetString(4), Is.EqualTo("column e"));
            Assert.That(excelReader.GetString(5), Is.EqualTo(" column b"));

            excelReader.Read();
            Assert.That(excelReader.GetDouble(0), Is.EqualTo(2));
            Assert.That(excelReader.GetString(1), Is.EqualTo("b"));
            Assert.That(excelReader.GetString(2), Is.EqualTo("c"));
            Assert.That(excelReader.GetString(3), Is.EqualTo("d"));
            Assert.That(excelReader.GetString(4), Is.EqualTo(" e "));

            excelReader.Read();
            Assert.That(excelReader.FieldCount, Is.EqualTo(6));
            Assert.That(excelReader.GetDouble(0), Is.EqualTo(3));
            Assert.That(excelReader.GetDouble(1), Is.EqualTo(2));
            Assert.That(excelReader.GetDouble(2), Is.EqualTo(3));
            Assert.That(excelReader.GetDouble(3), Is.EqualTo(4));
            Assert.That(excelReader.GetDouble(4), Is.EqualTo(5));

            excelReader.Read();
            Assert.That(excelReader.FieldCount, Is.EqualTo(6));
            Assert.That(excelReader.GetDouble(0), Is.EqualTo(4));
            Assert.That(excelReader.GetDateTime(1), Is.EqualTo(new DateTime(2012, 10, 13)));
            Assert.That(excelReader.GetDateTime(2), Is.EqualTo(new DateTime(2012, 10, 14)));
            Assert.That(excelReader.GetDateTime(3), Is.EqualTo(new DateTime(2012, 10, 15)));
            Assert.That(excelReader.GetDateTime(4), Is.EqualTo(new DateTime(2012, 10, 16)));

            for (int i = 4; i < 34; i++)
            {
                excelReader.Read();
                Assert.That(excelReader.GetDouble(0), Is.EqualTo(i + 1));
                Assert.That(excelReader.GetDouble(1), Is.EqualTo(i + 2));
                Assert.That(excelReader.GetDouble(2), Is.EqualTo(i + 3));
                Assert.That(excelReader.GetDouble(3), Is.EqualTo(i + 4));
                Assert.That(excelReader.GetDouble(4), Is.EqualTo(i + 5));
            }

            excelReader.NextResult();
            excelReader.Read();
            Assert.That(excelReader.FieldCount, Is.EqualTo(0));

            excelReader.NextResult();
            excelReader.Read();
            Assert.That(excelReader.FieldCount, Is.EqualTo(0));

            // test dataset
            DataSet result = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);
            Assert.That(result.Tables.Count, Is.EqualTo(3));
            Assert.That(result.Tables[0].Columns.Count, Is.EqualTo(6));
            Assert.That(result.Tables[0].Rows.Count, Is.EqualTo(33));

            Assert.That(result.Tables[0].Columns[0].ColumnName, Is.EqualTo("column a"));
            Assert.That(result.Tables[0].Columns[1].ColumnName, Is.EqualTo(" column b"));
            Assert.That(result.Tables[0].Columns[2].ColumnName, Is.EqualTo(" column b_1"));
            Assert.That(result.Tables[0].Columns[3].ColumnName, Is.EqualTo("Column3"));
            Assert.That(result.Tables[0].Columns[4].ColumnName, Is.EqualTo("column e"));
            Assert.That(result.Tables[0].Columns[5].ColumnName, Is.EqualTo(" column b_2"));

            Assert.That(Convert.ToInt32(result.Tables[0].Rows[0][0]), Is.EqualTo(2));
            Assert.That(result.Tables[0].Rows[0][1], Is.EqualTo("b"));
            Assert.That(result.Tables[0].Rows[0][2], Is.EqualTo("c"));
            Assert.That(result.Tables[0].Rows[0][3], Is.EqualTo("d"));
            Assert.That(result.Tables[0].Rows[0][4], Is.EqualTo(" e "));

            Assert.That(Convert.ToInt32(result.Tables[0].Rows[1][0]), Is.EqualTo(3));
            Assert.That(Convert.ToInt32(result.Tables[0].Rows[1][1]), Is.EqualTo(2));
            Assert.That(Convert.ToInt32(result.Tables[0].Rows[1][2]), Is.EqualTo(3));
            Assert.That(Convert.ToInt32(result.Tables[0].Rows[1][3]), Is.EqualTo(4));
            Assert.That(Convert.ToInt32(result.Tables[0].Rows[1][4]), Is.EqualTo(5));

            Assert.That(Convert.ToInt32(result.Tables[0].Rows[2][0]), Is.EqualTo(4));
            Assert.That(result.Tables[0].Rows[2][1], Is.EqualTo(new DateTime(2012, 10, 13)));
            Assert.That(result.Tables[0].Rows[2][2], Is.EqualTo(new DateTime(2012, 10, 14)));
            Assert.That(result.Tables[0].Rows[2][3], Is.EqualTo(new DateTime(2012, 10, 15)));
            Assert.That(result.Tables[0].Rows[2][4], Is.EqualTo(new DateTime(2012, 10, 16)));

            for (int i = 4; i < 33; i++)
            {
                Assert.That(Convert.ToInt32(result.Tables[0].Rows[i][0]), Is.EqualTo(i + 2));
                Assert.That(Convert.ToInt32(result.Tables[0].Rows[i][1]), Is.EqualTo(i + 3));
                Assert.That(Convert.ToInt32(result.Tables[0].Rows[i][2]), Is.EqualTo(i + 4));
                Assert.That(Convert.ToInt32(result.Tables[0].Rows[i][3]), Is.EqualTo(i + 5));
                Assert.That(Convert.ToInt32(result.Tables[0].Rows[i][4]), Is.EqualTo(i + 6));
            }

            // Test default and overridden column name prefix
            Assert.That(result.Tables[0].Columns[0].ColumnName, Is.EqualTo("column a"));
            Assert.That(result.Tables[0].Columns[3].ColumnName, Is.EqualTo("Column3"));

            DataSet prefixedResult = excelReader.AsDataSet(Configuration.FirstRowColumnNamesPrefixConfiguration);
            Assert.That(prefixedResult.Tables[0].Columns[0].ColumnName, Is.EqualTo("column a"));
            Assert.That(prefixedResult.Tables[0].Columns[3].ColumnName, Is.EqualTo("Prefix3"));
        }
    }
}
