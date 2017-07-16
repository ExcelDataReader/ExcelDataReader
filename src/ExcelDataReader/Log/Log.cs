using System;
using ExcelDataReader.Log.Logger;

namespace ExcelDataReader.Log
{
    /// <summary>
    /// logger type initialization
    /// </summary>
    public static class Log
    {
        private static readonly object LockObject = new object();

        private static Type logType = typeof(NullLogFactory);
        private static ILogFactory factoryInstance;

        /// <summary>
        /// Sets up logging to be with a certain type
        /// </summary>
        /// <typeparam name="T">The type of ILog for the application to use</typeparam>
        public static void InitializeWith<T>() 
            where T : ILogFactory, new()
        {
            lock (LockObject)
            {
                logType = typeof(T);
                factoryInstance = null;
            }
        }

        /// <summary>
        /// Initializes a new instance of a logger for an object.
        /// This should be done only once per object name.
        /// </summary>
        /// <param name="loggingType">The type to get a logger for.</param>
        /// <returns>ILog instance for an object if log type has been intialized; otherwise a null logger.</returns>
        public static ILog GetLoggerFor(Type loggingType)
        {
            var factory = factoryInstance;
            if (factory == null)
            {
                lock (LockObject)
                {
                    if (factory == null)
                    {
                        factory = factoryInstance = (ILogFactory)Activator.CreateInstance(logType);
                    }
                }
            }

            return factory.Create(loggingType);
        }
    }
}
