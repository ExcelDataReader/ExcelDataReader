#if NET20 || NET45
using ExcelDataReader.Log;
using log4net.Config;

[assembly: XmlConfigurator(Watch = true)]
namespace ExcelDataReader.Tests
{
    /// <summary>
    /// Log4net logger implementing special ILog class
    /// </summary>
    public class Log4NetLog : ILog, ILog<Log4NetLog>
    {
        private log4net.ILog logger;

        public void InitializeFor(string loggerName)
        {
            XmlConfigurator.Configure();
            logger = log4net.LogManager.GetLogger(loggerName);
        }

        public void Debug(string message, params object[] formatting)
        {
            logger.DebugFormat(message, formatting);
        }

        public void Info(string message, params object[] formatting)
        {
            logger.InfoFormat(message, formatting);
        }

        public void Warn(string message, params object[] formatting)
        {
            logger.WarnFormat(message, formatting);
        }

        public void Error(string message, params object[] formatting)
        {
            logger.ErrorFormat(message, formatting);
        }

        public void Fatal(string message, params object[] formatting)
        {
            logger.FatalFormat(message, formatting);
        }
    }
}
#endif
