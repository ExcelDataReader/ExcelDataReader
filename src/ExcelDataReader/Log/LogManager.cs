using System;
using System.Collections.Generic;

namespace ExcelDataReader.Log
{
    /// <summary>
    /// 2.0 version of LogExtensions, not as awesome as Extension methods
    /// </summary>
    public static class LogManager
    {
        /// <summary>
        /// Dictionary that ensures only one instance of a logger for a type.
        /// </summary>
        private static readonly Dictionary<string, ILog> _dictionary = new Dictionary<string, ILog>();

        private static readonly object _sync = new object();

        /// <summary>
        /// Gets the logger for <see cref="T"/>.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="type">The type to get the logger for.</param>
        /// <returns>Instance of a logger for the object.</returns>
        /// <remarks>This method is thread safe.</remarks>
        public static ILog Log<T>(T type)
        {
            string objectName = typeof(T).FullName;
            return Log(objectName);
        }

        /// <summary>
        /// Gets the logger for the specified object name.
        /// </summary>
        /// <param name="objectName">Either use the fully qualified object name or the short. If used with Log&lt;T&gt;() you must use the fully qualified object name"/></param>
        /// <returns>Instance of a logger for the object.</returns>
        /// <remarks>This method is thread safe.</remarks>
        public static ILog Log(string objectName)
        {
            lock (_sync)
            {
                ILog result;
                if (_dictionary.TryGetValue(objectName, out result))
                    return result;

                // The logger does not exist. Create it and add it to the dictionary.
                result = ExcelDataReader.Log.Log.GetLoggerFor(objectName);
                _dictionary.Add(objectName, result);

                return result;
            }
        }
    }
}

