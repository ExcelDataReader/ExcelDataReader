﻿using ExcelDataReader.Log.Logger;

namespace ExcelDataReader.Log;

/// <summary>
/// Logger type initialization.
/// </summary>
public static class Log
{
    private static Lazy<ILogFactory> factoryInstance = new(() => default(NullLogFactory));

    /// <summary>
    /// Sets up logging to be with a certain type.
    /// </summary>
    /// <typeparam name="T">The type of ILog for the application to use.</typeparam>
    public static void InitializeWith<T>() 
        where T : ILogFactory, new()
    {
        factoryInstance = new Lazy<ILogFactory>(() => new T());
    }

    /// <summary>
    /// Initializes a new instance of a logger for an object.
    /// This should be done only once per object name.
    /// </summary>
    /// <param name="loggingType">The type to get a logger for.</param>
    /// <returns>ILog instance for an object if log type has been initialized; otherwise a null logger.</returns>
    public static ILog GetLoggerFor(Type loggingType) => factoryInstance.Value.Create(loggingType);
}
