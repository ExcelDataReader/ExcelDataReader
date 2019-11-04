using System;
using System.Collections.Generic;
using System.Text;

using ExcelDataReader.Core.OpenXmlFormat.Records;

#nullable enable

namespace ExcelDataReader.Core.OpenXmlFormat
{
    internal abstract class RecordReader : IDisposable
    {
        ~RecordReader()
        {
            Dispose(false);
        }

        /// <inheritdoc />
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public abstract Record? Read();

        protected virtual void Dispose(bool disposing) 
        {
        }
    }
}
