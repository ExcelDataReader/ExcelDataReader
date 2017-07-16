using System;
#if NET20 || NET45
using System.Runtime.Serialization;
#endif

namespace ExcelDataReader.Exceptions
{
#if NET20 || NET45
    [Serializable]
#endif
    public class BiffRecordException : ExcelReaderException
    {
        public BiffRecordException()
        {
        }

        public BiffRecordException(string message)
            : base(message)
        {
        }

        public BiffRecordException(string message, Exception inner)
            : base(message, inner)
        {
        }

#if NET20 || NET45
        protected BiffRecordException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
#endif
    }
}
