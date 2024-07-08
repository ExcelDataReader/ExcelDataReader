namespace ExcelDataReader.Exceptions;

/// <summary>
/// Thrown when ExcelDataReader cannot open a password protected document because the password.
/// </summary>
/// <remarks>
/// Initializes a new instance of the <see cref="InvalidPasswordException"/> class.
/// </remarks>
/// <param name="message">The error message.</param>
public class InvalidPasswordException(string message) : ExcelReaderException(message)
{
}
