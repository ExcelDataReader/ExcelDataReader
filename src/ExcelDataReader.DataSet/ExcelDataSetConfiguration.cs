﻿namespace ExcelDataReader;

/// <summary>
/// Processing configuration options and callbacks for IExcelDataReader.AsDataSet().
/// </summary>
public class ExcelDataSetConfiguration
{
    /// <summary>
    /// Gets or sets a value indicating whether to set the DataColumn.DataType property in a second pass.
    /// </summary>
    public bool UseColumnDataType { get; set; } = true;

    /// <summary>
    /// Gets or sets a callback to obtain configuration options for a DataTable. 
    /// </summary>
    public Func<IExcelDataReader, ExcelDataTableConfiguration> ConfigureDataTable { get; set; }

    /// <summary>
    /// Gets or sets a callback to determine whether to include the current sheet in the DataSet. Called once per sheet before ConfigureDataTable.
    /// </summary>
    public Func<IExcelDataReader, int, bool> FilterSheet { get; set; }
}
