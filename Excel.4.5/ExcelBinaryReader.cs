using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using Excel.Portable;

namespace Excel
{
	/// <summary>
	/// ExcelDataReader Class
	/// </summary>
	public class ExcelBinaryReader : IExcelDataReader
	{
		#region Members

		private bool disposed;
        private ExcelDataReader.Portable.IExcelDataReader portable;

	    #endregion

		internal ExcelBinaryReader(ExcelDataReader.Portable.IExcelDataReader portableReader)
		{
            portable = portableReader;
		}

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

		~ExcelBinaryReader()
		{
			Dispose(false);
		}

		#endregion

		#region IExcelDataReader Members

		public void Initialize(Stream fileStream)
		{
            portable.Initialize(fileStream);
		}

		public DataSet AsDataSet()
		{
			return AsDataSet(false);
		}

		public DataSet AsDataSet(bool convertOADateTime)
		{
		    var datasetHelper = new DatasetHelper();
            portable.LoadDataSet(datasetHelper, convertOADateTime);

		    return (DataSet)datasetHelper.Dataset;

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

		public bool IsValid
		{
            get { return portable.IsValid; }
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
            get { return portable[i]; ; }
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

		#region IExcelDataReader Members


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

	    public ExcelDataReader.Portable.ReadOption SheetReadOption { get; private set; }

	    public bool ConvertOaDate
		{
            get { return portable.ConvertOaDate; }
            set { portable.ConvertOaDate = value; }
		}

		public ReadOption ReadOption
		{
            get { return (ReadOption)portable.SheetReadOption; }
		}

		#endregion
	}

	/// <summary>
	/// Strict is as normal, Loose is more forgiving and will not cause an exception if a record size takes it beyond the end of the file. It will be trunacted in this case (SQl Reporting Services)
	/// </summary>
	public enum ReadOption
	{
		Strict,
		Loose
	}
}