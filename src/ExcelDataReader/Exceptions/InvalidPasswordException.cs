#if !NET8_0_OR_GREATER
using System.Runtime.Serialization;
#endif

namespace ExcelDataReader.Exceptions;

/// <summary>
/// Thrown when ExcelDataReader cannot open a password-protected document because the password is missing or incorrect.
/// </summary>
#if !NET8_0_OR_GREATER
[Serializable]
#endif
public class InvalidPasswordException : ExcelReaderException
{
    /// <summary>
    /// Initializes a new instance of the <see cref="InvalidPasswordException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    public InvalidPasswordException(string message)
        : base(message)
    {
    }

#if !NET8_0_OR_GREATER
    /// <summary>
    /// Initializes a new instance of the <see cref="InvalidPasswordException"/> class.
    /// </summary>
    /// <param name="info">The serialization info.</param>
    /// <param name="context">The streaming context.</param>
    protected InvalidPasswordException(SerializationInfo info, StreamingContext context)
        : base(info, context)
    {
    }
#endif
}
