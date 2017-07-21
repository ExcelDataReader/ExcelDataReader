using System;

namespace ExcelDataReader.Log.Logger
{
    /// <summary>
    /// The default logger until one is set.
    /// </summary>
    public struct NullLogFactory : ILogFactory, ILog
    {
        /// <inheritdoc />
        public void Debug(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public void Info(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public void Warn(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public void Error(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public void Fatal(string message, params object[] formatting)
        {
        }

        /// <inheritdoc />
        public ILog Create(Type loggingType)
        {
            return this;
        }
    }
}
