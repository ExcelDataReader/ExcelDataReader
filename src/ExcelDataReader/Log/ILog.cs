using System;

namespace ExcelDataReader.Log
{
    /// <summary>
    /// Custom interface for logging messages
    /// </summary>
    public interface ILog
    {
        /// <summary>
        /// Debug level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Debug(string message, params object[] formatting);

        /// <summary>
        /// Info level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Info(string message, params object[] formatting);

        /// <summary>
        /// Warn level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Warn(string message, params object[] formatting);

        /// <summary>
        /// Error level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Error(string message, params object[] formatting);

        /// <summary>
        /// Fatal level of the specified message. The other method is preferred since the execution is deferred.
        /// </summary>
        /// <param name="message">The message.</param>
        /// <param name="formatting">The formatting.</param>
        void Fatal(string message, params object[] formatting);
    }

    /// <summary>
    /// Factory interface for loggers.
    /// </summary>
    public interface ILogFactory
    {
        /// <summary>
        /// Create a logger for the specified type.
        /// </summary>
        /// <param name="loggingType">The type to create a logger for.</param>
        /// <returns>The logger instance.</returns>
        ILog Create(Type loggingType);
    }
}