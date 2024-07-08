using System.Runtime.Serialization;

namespace ExcelDataReader.Exceptions;

/// <summary>
/// Base class for exceptions thrown by ExcelDataReader.
/// </summary>
#if !NET8_0_OR_GREATER
[Serializable]
#endif
public class ExcelReaderException : Exception
{
    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelReaderException"/> class.
    /// </summary>
    public ExcelReaderException()
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelReaderException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    public ExcelReaderException(string message)
        : base(message)
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelReaderException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="inner">The inner exception.</param>
    public ExcelReaderException(string message, Exception inner)
        : base(message, inner)
    {
    }

#if !NET8_0_OR_GREATER
    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelReaderException"/> class.
    /// </summary>
    /// <param name="info">The serialization info.</param>
    /// <param name="context">The streaming context.</param>
    protected ExcelReaderException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
#endif
}
