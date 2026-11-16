#if !NET8_0_OR_GREATER
using System.Runtime.Serialization;
#endif

namespace ExcelDataReader.Exceptions;

/// <summary>
/// Thrown when there is a problem parsing the Compound Document container format used by XLS and password-protected XLSX.
/// </summary>
#if !NET8_0_OR_GREATER
[Serializable]
#endif
public class CompoundDocumentException : ExcelReaderException
{
    /// <summary>
    /// Initializes a new instance of the <see cref="CompoundDocumentException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    public CompoundDocumentException(string message)
        : base(message)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="CompoundDocumentException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="inner">The inner exception.</param>
    public CompoundDocumentException(string message, Exception inner)
        : base(message, inner)
    {
    }

#if !NET8_0_OR_GREATER
    /// <summary>
    /// Initializes a new instance of the <see cref="CompoundDocumentException"/> class.
    /// </summary>
    /// <param name="info">The serialization info.</param>
    /// <param name="context">The streaming context.</param>
    protected CompoundDocumentException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
#endif
}
