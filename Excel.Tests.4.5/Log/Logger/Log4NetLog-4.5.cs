using System;
using ExcelDataReader.Portable.Log;

namespace Excel.Tests.Log.Logger
{
	/// <summary>
	/// Log4net logger implementing special ILog class
	/// </summary>
	public partial class Log4NetLog : ILog, ILog<Log4NetLog>
	{
		public void Debug(Func<string> message)
		{
			if (_logger.IsDebugEnabled) _logger.Debug(message.Invoke());
		}

		public void Info(Func<string> message)
		{
			if (_logger.IsInfoEnabled) _logger.Info(message.Invoke());
		}

		public void Warn(Func<string> message)
		{
			if (_logger.IsWarnEnabled) _logger.Warn(message.Invoke());
		}

		public void Error(Func<string> message)
		{
			// don't need to check for enabled at this level
			_logger.Error(message.Invoke());
		}

		public void Fatal(Func<string> message)
		{
			// don't need to check for enabled at this level
			_logger.Fatal(message.Invoke());
		}

	}
}
