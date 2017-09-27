namespace ExcelDataReader.Exceptions
{
    /// <summary>
    /// Thrown when ExcelDataReader cannot open a password protected document because the password
    /// </summary>
    public class InvalidPasswordException : ExcelReaderException
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="InvalidPasswordException"/> class.
        /// </summary>
        /// <param name="message">The error message</param>
        public InvalidPasswordException(string message)
            : base(message)
        {
        }
    }
}
