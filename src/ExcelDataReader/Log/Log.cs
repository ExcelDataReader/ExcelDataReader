using System;
using ExcelDataReader.Log.Logger;

namespace ExcelDataReader.Log
{
    /// <summary>
    /// logger type initialization
    /// </summary>
    public static class Log
    {
        private static Type logType = typeof(NullLog);
        private static ILog logger;

        /// <summary>
        /// Sets up logging to be with a certain type
        /// </summary>
        /// <typeparam name="T">The type of ILog for the application to use</typeparam>
        public static void InitializeWith<T>() 
            where T : ILog, new()
        {
            logType = typeof(T);
        }

        /// <summary>
        /// Sets up logging to be with a certain instance. The other method is preferred.
        /// </summary>
        /// <param name="loggerType">Type of the logger.</param>
        /// <remarks>This is mostly geared towards testing</remarks>
        public static void InitializeWith(ILog loggerType)
        {
            logType = loggerType.GetType();
            logger = loggerType;
        }

        /// <summary>
        /// Initializes a new instance of a logger for an object.
        /// This should be done only once per object name.
        /// </summary>
        /// <param name="objectName">Name of the object.</param>
        /// <returns>ILog instance for an object if log type has been intialized; otherwise null</returns>
        public static ILog GetLoggerFor(string objectName)
        {
            var logger = Log.logger;

            if (Log.logger == null)
            {
                logger = Activator.CreateInstance(logType) as ILog;
                logger?.InitializeFor(objectName);
            }

            return logger;
        }
    }
}
