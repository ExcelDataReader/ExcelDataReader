using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataReader.Portable.Data
{
    /// <summary>
    /// Implement this
    /// </summary>
    public interface IDatasetHelper
    {
        /// <summary>
        /// Is the operation valid
        /// </summary>
        bool IsValid     { get; set; }

        /// <summary>
        /// Create new dataset
        /// </summary>
        void CreateNew();

        /// <summary>
        /// Create new table
        /// </summary>
        /// <param name="name"></param>
        void CreateNewTable(string name);

        /// <summary>
        /// End loading data in to the table
        /// </summary>
        void EndLoadTable();

        /// <summary>
        /// Add a column to the table
        /// </summary>
        /// <param name="columnName"></param>
        void AddColumn(string columnName);

        /// <summary>
        /// Start loading data in to the table
        /// </summary>
        void BeginLoadData();

        /// <summary>
        /// Add a row to the current table with the supplied values
        /// </summary>
        /// <param name="values"></param>
        void AddRow(params object[] values);

        /// <summary>
        /// Dataset loading is finished
        /// </summary>
        void DatasetLoadComplete();
    }

    //public interface ITableHelper
    //{
    //    void CreateNew(string sheetName);
    //}

    //public interface ITableHelperFactory
    //{
    //    IDatasetHelper CreateDataset();
    //    ITableHelper CreateTable();
    //}
}
