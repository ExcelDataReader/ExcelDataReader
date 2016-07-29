#if NET20 || NET45
using ExcelDataReader.Log;
using log4net.Config;

[assembly: XmlConfigurator(Watch = true)]
namespace ExcelDataReader.Tests.Log.Logger
{
	/// <summary>
	/// Log4net logger implementing special ILog class
	/// </summary>
	public partial class Log4NetLog : ILog, ILog<Log4NetLog>
	{
		private global::log4net.ILog _logger;

		public void InitializeFor(string loggerName)
		{
			log4net.Config.XmlConfigurator.Configure();
			_logger = log4net.LogManager.GetLogger(loggerName);
		}

		public void Debug(string message, params object[] formatting)
		{
			if (_logger.IsDebugEnabled) _logger.DebugFormat(message, formatting);
		}

		public void Info(string message, params object[] formatting)
		{
			if (_logger.IsInfoEnabled) _logger.InfoFormat(message, formatting);
		}

		public void Warn(string message, params object[] formatting)
		{
			if (_logger.IsWarnEnabled) _logger.WarnFormat(message, formatting);
		}

		public void Error(string message, params object[] formatting)
		{
			// don't need to check for enabled at this level
			_logger.ErrorFormat(message, formatting);
		}

		public void Fatal(string message, params object[] formatting)
		{
			// don't need to check for enabled at this level
			_logger.FatalFormat(message, formatting);
		}

	}
}
#endif
