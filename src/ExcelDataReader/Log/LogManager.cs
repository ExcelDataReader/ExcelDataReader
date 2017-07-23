namespace ExcelDataReader.Log
{
    /// <summary>
    /// 2.0 version of LogExtensions, not as awesome as Extension methods
    /// </summary>
    public static class LogManager
    {
        /// <summary>
        /// Gets the logger for a type.
        /// </summary>
        /// <typeparam name="T">The type to fetch a logger for.</typeparam>
        /// <param name="type">The type to get the logger for.</param>
        /// <returns>Instance of a logger for the object.</returns>
        /// <remarks>This method is thread safe.</remarks>
        public static ILog Log<T>(T type)
        {
            return ExcelDataReader.Log.Log.GetLoggerFor(typeof(T));
        }
    }
}
