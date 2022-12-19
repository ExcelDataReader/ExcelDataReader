using System;
using System.Runtime.Serialization;

namespace ExcelDataReader.Exceptions
{
    /// <summary>
    /// Thrown when ExcelDataReader cannot parse the header.
    /// </summary>
    [Serializable]
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
        /// <param name="message">The error message.</param>
        public HeaderException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        /// <param name="message">The error message.</param>
        /// <param name="inner">The inner exception.</param>
        public HeaderException(string message, Exception inner)
            : base(message, inner)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="HeaderException"/> class.
        /// </summary>
        /// <param name="info">The serialization info.</param>
        /// <param name="context">The streaming context.</param>
        protected HeaderException(SerializationInfo info, StreamingContext context)
            : base(info, context)
        {
        }
    }
}
