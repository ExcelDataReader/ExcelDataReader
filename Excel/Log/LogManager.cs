using System;
using System.Collections.Generic;
using System.Text;

namespace Excel.Log
{
	/// <summary>
	/// 2.0 version of LogExtensions, not as awesome as Extension methods
	/// </summary>
	public static class LogManager
	{
		/// <summary>
		/// Concurrent dictionary that ensures only one instance of a logger for a type.
		/// </summary>
		private static readonly Dictionary<string, ILog> _dictionary = new Dictionary<string, ILog>();

		private static object _sync = new Object();

		/// <summary>
		/// Gets the logger for <see cref="T"/>.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="type">The type to get the logger for.</param>
		/// <returns>Instance of a logger for the object.</returns>
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
		public static ILog Log(string objectName)
		{
			ILog result = null;

			if (_dictionary.ContainsKey(objectName))
				result = _dictionary[objectName];

			if (result == null)
			{
				lock (_sync)
				{
					result = Excel.Log.Log.GetLoggerFor(objectName);
					_dictionary.Add(objectName, result);
				}
			}
			
			return result;
		}
	}
}
