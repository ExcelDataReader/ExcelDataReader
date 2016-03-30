using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Excel
{
    public interface IExcelWorker : IDisposable
    {
        /// <summary>
        /// Gets a value indicating whether this instance is valid.
        /// </summary>
        /// <value><c>true</c> if this instance is valid; otherwise, <c>false</c>.</value>
        bool IsValid { get; }

        /// <summary>
        /// Gets the exception message.
        /// </summary>
        /// <value>The exception message.</value>
        string ExceptionMessage { get; }

        /// <summary>
        /// Extracts the specified zip file stream.
        /// </summary>
        /// <param name="fileStream">The zip file stream.</param>
        /// <returns></returns>
        bool Extract(Stream fileStream);

        /// <summary>
        /// Gets the shared strings stream.
        /// </summary>
        /// <returns></returns>
        Stream GetSharedStringsStream();

        /// <summary>
        /// Gets the styles stream.
        /// </summary>
        /// <returns></returns>
        Stream GetStylesStream();

        /// <summary>
        /// Gets the workbook stream.
        /// </summary>
        /// <returns></returns>
        Stream GetWorkbookStream();

        /// <summary>
        /// Gets the worksheet stream.
        /// </summary>
        /// <param name="sheetId">The sheet id.</param>
        /// <returns></returns>
        Stream GetWorksheetStream(int sheetId);

        /// <summary>
        /// Gets the worksheet stream.
        /// </summary>
        /// <param name="sheetId">The sheet path.</param>
        /// <returns></returns>
        Stream GetWorksheetStream(string sheetPath);

        /// <summary>
        /// Gets the workbook rels stream.
        /// </summary>
        /// <returns></returns>
        Stream GetWorkbookRelsStream();
    }
}
