using System;
#if NET20 || NET45
using System.Runtime.Serialization;
#endif

namespace ExcelDataReader.Exceptions
{
#if NET20 || NET45
    [Serializable]
#endif
    public class ExcelReaderException : Exception
    {
        public ExcelReaderException()
        {
        }

        public ExcelReaderException(string message)
            : base(message)
        {
        }

        public ExcelReaderException(string message, Exception inner)
            : base(message, inner)
        {
        }

#if NET20 || NET45
        protected ExcelReaderException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
#endif
    }
}
