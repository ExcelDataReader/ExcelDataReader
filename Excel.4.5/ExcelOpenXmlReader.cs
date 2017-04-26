//#define DEBUGREADERS

using ExcelDataReader.Desktop.Portable;
using ExcelDataReader.Portable.Async;
using System;
using System.Data;
using System.Text;

namespace Excel
{

    public class ExcelOpenXmlReader : IExcelDataReader
    {
        private readonly ExcelDataReader.Portable.IExcelDataReader portable;
        private bool disposed;

        #region Members


        #endregion

        internal ExcelOpenXmlReader(ExcelDataReader.Portable.IExcelDataReader portable)
        {
            this.portable = portable;
            //_isValid = true;
            //_isFirstRead = true;

            //_defaultDateTimeStyles = new List<int>(new int[] 
            //{
            //    14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47
            //});

        }



        #region IExcelDataReader Members

        public void Initialize(System.IO.Stream fileStream)
        {
            AsyncHelper.RunSync(() => portable.InitializeAsync(fileStream));
        }

        public DataSet AsDataSet()
        {
            return AsDataSet(true);
        }

        //todo: this is identical in ExcelBinaryReader
        public DataSet AsDataSet(bool convertOADateTime)
        {
            var datasetHelper = new DatasetHelper();
            AsyncHelper.RunSync(() => portable.LoadDataSetAsync(datasetHelper, convertOADateTime));

            return (DataSet)datasetHelper.Dataset;
        }

        public bool IsFirstRowAsColumnNames
        {
            get
            {
                return portable.IsFirstRowAsColumnNames;
            }
            set
            {
                portable.IsFirstRowAsColumnNames = value;
            }
        }

        public bool DoAllowEmptyTables
        {
            get { return portable.DoAllowEmptyTables; }
            set { portable.DoAllowEmptyTables = value; }
        }

        public Encoding Encoding
        {
            get { return portable.Encoding; }
        }

        public Encoding DefaultEncoding
        {
            get { return portable.DefaultEncoding; }
        }

        public bool IsValid
        {
            get { return portable.IsValid; }
        }

        public string ExceptionMessage
        {
            get { return portable.ExceptionMessage; }
        }

        public string Name
        {
            get
            {
                return portable.Name;
            }
        }

        public string VisibleState
        {
            get { return portable.VisibleState; }
        }

        public void Close()
        {
            portable.Close();
        }

        public int Depth
        {
            get { return portable.Depth; }
        }

        public int ResultsCount
        {
            get { return portable.ResultsCount; }
        }

        public bool IsClosed
        {
            get { return portable.IsClosed; }
        }

        public bool NextResult()
        {
            return portable.NextResult();
        }

        public bool Read()
        {
            return portable.Read();
        }

        public int FieldCount
        {
            get { return portable.FieldCount; }
        }

        public bool GetBoolean(int i)
        {
            return portable.GetBoolean(i);
        }

        public DateTime GetDateTime(int i)
        {
            return portable.GetDateTime(i);

        }

        public decimal GetDecimal(int i)
        {
            return portable.GetDecimal(i);
        }

        public double GetDouble(int i)
        {
            return portable.GetDouble(i);
        }

        public float GetFloat(int i)
        {
            return portable.GetFloat(i);
        }

        public short GetInt16(int i)
        {
            return portable.GetInt16(i);
        }

        public int GetInt32(int i)
        {
            return portable.GetInt32(i);
        }

        public long GetInt64(int i)
        {
            return portable.GetInt64(i);
        }

        public string GetString(int i)
        {
            return portable.GetString(i);
        }

        public object GetValue(int i)
        {
            return portable.GetValue(i);
        }

        public bool IsDBNull(int i)
        {
            return portable.IsDBNull(i);
        }

        public object this[int i]
        {
            get { return portable[i]; }
        }

        #endregion

        #region IDisposable Members

        public void Dispose()
        {
            Dispose(true);

            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called.

            if (!this.disposed)
            {
                if (disposing)
                {
                    portable.Dispose();
                }


                disposed = true;
            }
        }

        ~ExcelOpenXmlReader()
        {
            Dispose(false);
        }

        #endregion

        #region  Not Supported IDataReader Members



        public DataTable GetSchemaTable()
        {
            throw new NotSupportedException();
        }

        public int RecordsAffected
        {
            get { throw new NotSupportedException(); }
        }

        #endregion

        #region Not Supported IDataRecord Members


        public byte GetByte(int i)
        {
            throw new NotSupportedException();
        }

        public long GetBytes(int i, long fieldOffset, byte[] buffer, int bufferoffset, int length)
        {
            throw new NotSupportedException();
        }

        public char GetChar(int i)
        {
            throw new NotSupportedException();
        }

        public long GetChars(int i, long fieldoffset, char[] buffer, int bufferoffset, int length)
        {
            throw new NotSupportedException();
        }

        public IDataReader GetData(int i)
        {
            throw new NotSupportedException();
        }

        public string GetDataTypeName(int i)
        {
            throw new NotSupportedException();
        }

        public Type GetFieldType(int i)
        {
            throw new NotSupportedException();
        }

        public Guid GetGuid(int i)
        {
            throw new NotSupportedException();
        }

        public string GetName(int i)
        {
            throw new NotSupportedException();
        }

        public int GetOrdinal(string name)
        {
            throw new NotSupportedException();
        }

        public int GetValues(object[] values)
        {
            throw new NotSupportedException();
        }

        public object this[string name]
        {
            get { throw new NotSupportedException(); }
        }

        #endregion


    }
}
