using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel.Log
{
	/// <summary>
	/// Custom interface for logging messages
	/// </summary>
	public partial interface ILog
	{
		/// <summary>
		/// Debug level of the specified message.
		/// </summary>
		/// <param name="message">The message.</param>
		void Debug(Func<string> message);

		/// <summary>
		/// Info level of the specified message.
		/// </summary>
		/// <param name="message">The message.</param>
		void Info(Func<string> message);

		/// <summary>
		/// Warn level of the specified message.
		/// </summary>
		/// <param name="message">The message.</param>
		void Warn(Func<string> message);

		/// <summary>
		/// Error level of the specified message.
		/// </summary>
		/// <param name="message">The message.</param>
		void Error(Func<string> message);

		/// <summary>
		/// Fatal level of the specified message.
		/// </summary>
		/// <param name="message">The message.</param>
		void Fatal(Func<string> message);
	}


}
