using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using ExcelDataReader.Core.OpenXmlFormat.Records;

namespace ExcelDataReader.Core.OpenXmlFormat.XmlFormat
{
    internal abstract class XmlRecordReader : RecordReader
    {
        private IEnumerator<Record> _enumerator;

        public XmlRecordReader(XmlReader reader)
        {
            Reader = reader;
        }

        protected XmlReader Reader { get; }

        public override Record Read()
        {
            if (_enumerator == null)
                _enumerator = ReadOverride().GetEnumerator();
            if (_enumerator.MoveNext())
                return _enumerator.Current;
            return null;
        }

        protected abstract IEnumerable<Record> ReadOverride();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            _enumerator?.Dispose();
#if NET20
            if (disposing)
                Reader.Close();
#else
            if (disposing)
                Reader.Dispose();
#endif
        }
    }
}
