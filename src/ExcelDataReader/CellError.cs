namespace ExcelDataReader
{
#pragma warning disable SA1629 // Documentation text should end with a period
    /// <summary>
    /// Formula error. 
    /// </summary>
    public enum CellError : byte
    {
        /// <summary>
        /// #NULL!
        /// </summary>
        NULL = 0x00,

        /// <summary>
        /// #DIV/0!
        /// </summary>
        DIV0 = 0x07,

        /// <summary>
        /// #VALUE!
        /// </summary>
        VALUE = 0x0F,

        /// <summary>
        /// #REF!
        /// </summary>
        REF = 0x17,

        /// <summary>
        /// #NAME?
        /// </summary>
        NAME = 0x1D,

        /// <summary>
        /// #NUM!
        /// </summary>
        NUM = 0x24,

        /// <summary>
        /// #N/A
        /// </summary>
        NA = 0x2A,

        /// <summary>
        /// #GETTING_DATA
        /// </summary>
#pragma warning disable CA1707 // Identifiers should not contain underscores
        GETTING_DATA = 0x2B,
#pragma warning restore CA1707 // Identifiers should not contain underscores
    }
#pragma warning restore SA1629 // Documentation text should end with a period
}
