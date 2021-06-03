// Copyright (c) 2021 gojimmypi
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient; // NuGet package, version 4.8.2 https://github.com/dotnet/SqlClient 
using System.IO;
using System.Text;
using ExcelDataReader;

namespace BulkCopyStreamReader
{


    class Program
    {
        const int SQL_TIMEOUT = 60;
        const int MAX_COLUMNS = 1000;

        /// ***********************************************************************************************************************************
        /// <summary>
        /// integrated authentication SQL connection string
        /// </summary>
        /// <param name="server"></param>
        /// <param name="database"></param>
        /// <returns></returns>
        public static string TrustedConnectionString(string server, string database)
        //***********************************************************************************************************************************
        {
            // see http://msdn.microsoft.com/library/default.asp?url=/library/en-us/cpref/html/frlrfSystemDataSqlClientSqlConnectionClassConnectionStringTopic.asp
            //
            // there is some debate as to whether the Oledb provider is indeed faster than the native client!
            //  
            string appname = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            string computername = Environment.MachineName.ToString();
            return "Workstation ID=" + computername + "_" + appname + ";" +
                   "packet size=8192;" +
                   "Persist Security Info=false;" +
                   "Server=" + server + ";" +
                   "Database=" + database + ";" +
                   "Trusted_Connection=true; " +
                   // "Network Library=dbmssocn;" +
                   "Pooling=True; " +
                   "Enlist=True; " +
                   "Connection Lifetime=14400; " +
                   "Max Pool Size=20; Min Pool Size=0";
        }

        private static void ReadFile(string FilePath, ref DataSet ds, long StartRow = 0, long EndRow = 0)
        {
            // sample dataset reader code from: https://github.com/ExcelDataReader/ExcelDataReader#asdataset-configuration-options
            ds = new DataSet();
            long ThisRowNumber = 0;
            bool UseRange = (StartRow > 0) && (EndRow > 0);
            using (var stream = File.Open(FilePath, FileMode.Open, FileAccess.Read))
            {
                var reader = ExcelReaderFactory.CreateOpenXmlReader(stream, new ExcelReaderConfiguration()
                {
                    // Gets or sets the encoding to use when the input XLS lacks a CodePage
                    // record, or when the input CSV lacks a BOM and does not parse as UTF8. 
                    // Default: cp1252 (XLS BIFF2-5 and CSV only)
                    FallbackEncoding = Encoding.UTF8,

                    // Gets or sets the password used to open password protected workbooks.
                    // Password = "password",

                    // Gets or sets an array of CSV separator candidates. The reader 
                    // autodetects which best fits the input data. Default: , ; TAB | # 
                    // (CSV only)
                    // AutodetectSeparators = new char[] { ',', ';', '\t', '|', '#' },

                    // Gets or sets a value indicating whether to leave the stream open after
                    // the IExcelDataReader object is disposed. Default: false
                    LeaveOpen = false,

                    // Gets or sets a value indicating the number of rows to analyze for
                    // encoding, separator and field count in a CSV. When set, this option
                    // causes the IExcelDataReader.RowCount property to throw an exception.
                    // Default: 0 - analyzes the entire file (CSV only, has no effect on other
                    // formats)
                    AnalyzeInitialCsvRows = 0,
                });

                ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    // Gets or sets a value indicating whether to set the DataColumn.DataType 
                    // property in a second pass.
                    UseColumnDataType = false,

                    // Gets or sets a callback to determine whether to include the current sheet
                    // in the DataSet. Called once per sheet before ConfigureDataTable.
                    FilterSheet = (tableReader, sheetIndex) => true,

                    // Gets or sets a callback to obtain configuration options for a DataTable. 
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        // Gets or sets a value indicating the prefix of generated column names.
                        EmptyColumnNamePrefix = "Column",

                        // Gets or sets a value indicating whether to use a row from the 
                        // data as column names.
                        UseHeaderRow = true,

                        // Gets or sets a callback to determine which row is the header row. 
                        // Only called when UseHeaderRow = true.
                        ReadHeaderRow = (rowReader) => {
                            // F.ex skip the first row and use the 2nd row as column headers:
                            // rowReader.Read();
                        },

                        // Gets or sets a callback to determine whether to include the 
                        // current row in the DataTable.
                        FilterRow = (rowReader) => {
                            if (UseRange)
                            {
                                ThisRowNumber++;
                                return ((ThisRowNumber >= StartRow) && (ThisRowNumber <= EndRow));
                            }
                            else
                            {
                                return true;
                            }

                        },

                        // Gets or sets a callback to determine whether to include the specific
                        // column in the DataTable. Called once per column after reading the 
                        // headers.
                        FilterColumn = (rowReader, columnIndex) => {
                            // currently including all rows
                            return true;
                        }
                    }
                });
            }



        }


        private static void BulkInsertXLS(string fromFileXLSX, string toServer, string toDatabase, string toTable)
        {
            // we need a non-blank thisDS to determine the dataset_definition_code that is needed to get the IncrementalLoad value
            // so loadup just the header from our file
            DataSet xlsHeader = new ();
            ReadFile(fromFileXLSX, ref xlsHeader,0,0);


            using (var stream = File.Open(fromFileXLSX, FileMode.Open, FileAccess.Read))
            {
                // in this example, we do not have the upload_session_id and create_app_user_id fields in the XLS, but we want them included at bulk insert time
                Dictionary<string, object> myInjectedFields = new();
                myInjectedFields.Add("upload_session_id", 1234); // your process would pick a session id, typically via a SQL stored proc and/or identity field.
                myInjectedFields.Add("create_app_user_id", Environment.UserName);

                using (var r = ExcelReaderFactory.CreateReader(stream, myInjectedFields))
                {
                    // var options = SqlBulkCopyOptions.TableLock | SqlBulkCopyOptions.KeepNulls; // | SqlBulkCopyOptions.FireTriggers; // SqlBulkCopyOptions.CheckConstraints | SqlBulkCopyOptions.KeepIdentity;
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(TrustedConnectionString(toServer, toDatabase)))
                    {
                        bulkCopy.DestinationTableName = "[" + toDatabase + "].dbo.[" + toTable + "]";
                        bulkCopy.BulkCopyTimeout = SQL_TIMEOUT;

                        int foundColumnCount = xlsHeader.Tables[0].Columns.Count;

                        // column mapping
                        foreach (DataColumn column in xlsHeader.Tables[0].Columns)
                        {
                            // Console.WriteLine(column.ColumnName);
                            bulkCopy.ColumnMappings.Add(column.ColumnName, "[" + column.ColumnName + "]");
                        }

                        // recall the Header is from XLS, missing injected fields, so we need to add them:
                        bulkCopy.ColumnMappings.Add("upload_session_id", "upload_session_id");
                        bulkCopy.ColumnMappings.Add("create_app_user_id", "create_app_user_id");

                        // bulkCopy.EnableStreaming = true;

                        // If you start from an empty table, the operation will be minimally logged. However, if there is already data in the table, the operation will be logged unless you use trace flag 610.
                        // see https://docs.microsoft.com/en-us/previous-versions/sql/sql-server-2008/dd425070(v=sql.100)?redirectedfrom=MSDN
                        /// SqlHelper.ExecuteNonQuery(UploadHelper.TrustedConnectionString(toServer, toDatabase), CommandType.Text, "DBCC TRACEON (310);");
                        bulkCopy.WriteToServer(r);
                    }
                    // SqlHelper.ExecuteNonQuery(UploadHelper.TrustedConnectionString(toServer, toDatabase), CommandType.Text, "DBCC TRACEOFF (310);");
                }

            }
        } // close using SqlDbDataReader r

        static void Main(string[] args)
        {
            // see https://github.com/ExcelDataReader/ExcelDataReader#important-note-on-net-core
            // By default, ExcelDataReader throws a NotSupportedException "No data is available for encoding 1252." on .NET Core.
            //
            // To fix, add a dependency to the package System.Text.Encoding.CodePages and then add code to register the code page provider
            // during application initialization(f.ex in Startup.cs):
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // we'll typically start in .\ExcelDataReader\Example\BulkCopyStreamReader\bin\Debug\net5.0
            // sample is located in     .\ExcelDataReader\Example\XLS
            string theXlsFile = @"..\\..\\..\\..\\..\\Example\\XLS\\test_file.xlsx";
            string BulkInsertServer = "yourserver";
            string BulkInsertDatabase = "yourdatabase";
            string BulkInsertTable = "bulk_test";
            BulkInsertXLS(fromFileXLSX: theXlsFile,
                     toServer: BulkInsertServer,
                     toDatabase: BulkInsertDatabase,
                     toTable: BulkInsertTable);
        }
    }
}