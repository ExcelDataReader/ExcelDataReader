using System;
using System.Data;

using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    public static class AssertUtilities
    {
        public static void DoOpenOfficeTest(IExcelDataReader excelReader)
        {
            excelReader.Read();
            Assert.AreEqual(6, excelReader.FieldCount);
            Assert.AreEqual("column a", excelReader.GetString(0));
            Assert.AreEqual(" column b", excelReader.GetString(1));
            Assert.AreEqual(" column b", excelReader.GetString(2));
            Assert.IsNull(excelReader.GetString(3));
            Assert.AreEqual("column e", excelReader.GetString(4));
            Assert.AreEqual(" column b", excelReader.GetString(5));

            excelReader.Read();
            Assert.AreEqual(2, excelReader.GetDouble(0));
            Assert.AreEqual("b", excelReader.GetString(1));
            Assert.AreEqual("c", excelReader.GetString(2));
            Assert.AreEqual("d", excelReader.GetString(3));
            Assert.AreEqual(" e ", excelReader.GetString(4));

            excelReader.Read();
            Assert.AreEqual(6, excelReader.FieldCount);
            Assert.AreEqual(3, excelReader.GetDouble(0));
            Assert.AreEqual(2, excelReader.GetDouble(1));
            Assert.AreEqual(3, excelReader.GetDouble(2));
            Assert.AreEqual(4, excelReader.GetDouble(3));
            Assert.AreEqual(5, excelReader.GetDouble(4));

            excelReader.Read();
            Assert.AreEqual(6, excelReader.FieldCount);
            Assert.AreEqual(4, excelReader.GetDouble(0));
            Assert.AreEqual(new DateTime(2012, 10, 13), excelReader.GetDateTime(1));
            Assert.AreEqual(new DateTime(2012, 10, 14), excelReader.GetDateTime(2));
            Assert.AreEqual(new DateTime(2012, 10, 15), excelReader.GetDateTime(3));
            Assert.AreEqual(new DateTime(2012, 10, 16), excelReader.GetDateTime(4));

            for (int i = 4; i < 34; i++)
            {
                excelReader.Read();
                Assert.AreEqual(i + 1, excelReader.GetDouble(0));
                Assert.AreEqual(i + 2, excelReader.GetDouble(1));
                Assert.AreEqual(i + 3, excelReader.GetDouble(2));
                Assert.AreEqual(i + 4, excelReader.GetDouble(3));
                Assert.AreEqual(i + 5, excelReader.GetDouble(4));
            }

            excelReader.NextResult();
            excelReader.Read();
            Assert.AreEqual(0, excelReader.FieldCount);

            excelReader.NextResult();
            excelReader.Read();
            Assert.AreEqual(0, excelReader.FieldCount);

            // test dataset
            DataSet result = excelReader.AsDataSet(Configuration.FirstRowColumnNamesConfiguration);
            Assert.AreEqual(3, result.Tables.Count);
            Assert.AreEqual(6, result.Tables[0].Columns.Count);
            Assert.AreEqual(33, result.Tables[0].Rows.Count);

            Assert.AreEqual("column a", result.Tables[0].Columns[0].ColumnName);
            Assert.AreEqual(" column b", result.Tables[0].Columns[1].ColumnName);
            Assert.AreEqual(" column b_1", result.Tables[0].Columns[2].ColumnName);
            Assert.AreEqual("Column3", result.Tables[0].Columns[3].ColumnName);
            Assert.AreEqual("column e", result.Tables[0].Columns[4].ColumnName);
            Assert.AreEqual(" column b_2", result.Tables[0].Columns[5].ColumnName);

            Assert.AreEqual(2, Convert.ToInt32(result.Tables[0].Rows[0][0]));
            Assert.AreEqual("b", result.Tables[0].Rows[0][1]);
            Assert.AreEqual("c", result.Tables[0].Rows[0][2]);
            Assert.AreEqual("d", result.Tables[0].Rows[0][3]);
            Assert.AreEqual(" e ", result.Tables[0].Rows[0][4]);

            Assert.AreEqual(3, Convert.ToInt32(result.Tables[0].Rows[1][0]));
            Assert.AreEqual(2, Convert.ToInt32(result.Tables[0].Rows[1][1]));
            Assert.AreEqual(3, Convert.ToInt32(result.Tables[0].Rows[1][2]));
            Assert.AreEqual(4, Convert.ToInt32(result.Tables[0].Rows[1][3]));
            Assert.AreEqual(5, Convert.ToInt32(result.Tables[0].Rows[1][4]));

            Assert.AreEqual(4, Convert.ToInt32(result.Tables[0].Rows[2][0]));
            Assert.AreEqual(new DateTime(2012, 10, 13), result.Tables[0].Rows[2][1]);
            Assert.AreEqual(new DateTime(2012, 10, 14), result.Tables[0].Rows[2][2]);
            Assert.AreEqual(new DateTime(2012, 10, 15), result.Tables[0].Rows[2][3]);
            Assert.AreEqual(new DateTime(2012, 10, 16), result.Tables[0].Rows[2][4]);

            for (int i = 4; i < 33; i++)
            {
                Assert.AreEqual(i + 2, Convert.ToInt32(result.Tables[0].Rows[i][0]));
                Assert.AreEqual(i + 3, Convert.ToInt32(result.Tables[0].Rows[i][1]));
                Assert.AreEqual(i + 4, Convert.ToInt32(result.Tables[0].Rows[i][2]));
                Assert.AreEqual(i + 5, Convert.ToInt32(result.Tables[0].Rows[i][3]));
                Assert.AreEqual(i + 6, Convert.ToInt32(result.Tables[0].Rows[i][4]));
            }

            // Test default and overridden column name prefix
            Assert.AreEqual("column a", result.Tables[0].Columns[0].ColumnName);
            Assert.AreEqual("Column3", result.Tables[0].Columns[3].ColumnName);

            DataSet prefixedResult = excelReader.AsDataSet(Configuration.FirstRowColumnNamesPrefixConfiguration);
            Assert.AreEqual("column a", prefixedResult.Tables[0].Columns[0].ColumnName);
            Assert.AreEqual("Prefix3", prefixedResult.Tables[0].Columns[3].ColumnName);

        }
    }
}
