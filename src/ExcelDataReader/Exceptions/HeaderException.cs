using System;
#if NET20 || NET45 || NETSTANDARD2_0
using System.Runtime.Serialization;
#endif

namespace ExcelDataReader.Exceptions
{
    /// <summary>
    /// Thrown when ExcelDataReader cannot parse the header
    /// </summary>
#if NET20 || NET45 || NETSTANDARD2_0
    [Serializable]
#endif
    public class HeaderException : ExcelReaderException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        public HeaderException()
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        public HeaderException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        /// <param name="inner">The inner exception</param>
        public HeaderException(string message, Exception inner)
            : base(message, inner)
        {
        }

#if NET20 || NET45 || NETSTANDARD2_0
        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        /// <param name="info">The serialization info</param>
        /// <param name="context">The streaming context</param>
        protected HeaderException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
#endif
    }
}
