using System;

namespace ExcelDataReader.Exceptions
{
    /// <summary>
    /// Thrown when there is a problem parsing the Compound Document container format used by XLS and password-protected XLSX.
    /// </summary>
    public class CompoundDocumentException : ExcelReaderException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CompoundDocumentException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        public CompoundDocumentException(string message)
            : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="CompoundDocumentException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        /// <param name="inner">The inner exception</param>
        public CompoundDocumentException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
