#if !NET8_0_OR_GREATER
using System.Runtime.Serialization;
#endif

namespace ExcelDataReader.Exceptions;

/// <summary>
/// Thrown when ExcelDataReader cannot parse the header.
/// </summary>
#if !NET8_0_OR_GREATER
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

#if !NET8_0_OR_GREATER
    /// <summary>
    /// Initializes a new instance of the <see cref="HeaderException"/> class.
    /// </summary>
    /// <param name="info">The serialization info.</param>
    /// <param name="context">The streaming context.</param>
    protected HeaderException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
#endif
}
