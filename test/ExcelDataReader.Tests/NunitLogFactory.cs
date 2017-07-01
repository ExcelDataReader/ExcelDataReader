using System;

using ExcelDataReader.Log;

using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    public class NunitLogFactory : ILogFactory
    {
        /// <inheritdoc />
        public ILog Create(Type loggingType)
        {
            return new NunitLog(loggingType);
        }
    }

    /// <summary>
    /// ILog implementation that outputs to the nunit test context. 
    /// </summary>
    public class NunitLog : ILog
    {
        private readonly Type _loggingType;

        public NunitLog(Type loggingType)
        {
            _loggingType = loggingType;
        }

        public void Debug(string message, params object[] formatting)
        {
            // Too much is logged at the debug level. 
            // WriteMessage(_loggingType, "DEBUG", message, formatting);
        }

        public void Info(string message, params object[] formatting)
        {
            WriteMessage(_loggingType, "INFO", message, formatting);
        }

        public void Warn(string message, params object[] formatting)
        {
            WriteMessage(_loggingType, "WARN", message, formatting);
        }

        public void Error(string message, params object[] formatting)
        {
            WriteMessage(_loggingType, "ERROR", message, formatting);
        }

        public void Fatal(string message, params object[] formatting)
        {
            WriteMessage(_loggingType, "FATAL", message, formatting);
        }

        private static void WriteMessage(Type type, string level, string format, object[] args)
        {
            TestContext.Out.Write("{0} {1}:", type.FullName, level);
            TestContext.Out.WriteLine(format, args);
        }
    }
}
