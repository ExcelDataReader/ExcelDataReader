using System;

namespace ExcelDataReader.Log.Logger
{
    /// <summary>
    /// The default logger until one is set.
    /// </summary>
    public struct NullLogFactory : ILogFactory, ILog
    {
        public void Debug(string message, params object[] formatting)
        {
        }

        public void Info(string message, params object[] formatting)
        {
        }

        public void Warn(string message, params object[] formatting)
        {
        }

        public void Error(string message, params object[] formatting)
        {
        }

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
