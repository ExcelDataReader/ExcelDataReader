using System;
#if NET20 || NET45
using System.Runtime.Serialization;
#endif

namespace ExcelDataReader.Exceptions
{
#if NET20 || NET45
    [Serializable]
#endif
    public class HeaderException : ExcelReaderException
    {
        public HeaderException()
        {
        }

        public HeaderException(string message)
            : base(message)
        {
        }

        public HeaderException(string message, Exception inner)
            : base(message, inner)
        {
        }

#if NET20 || NET45
        protected HeaderException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
#endif
    }
}
