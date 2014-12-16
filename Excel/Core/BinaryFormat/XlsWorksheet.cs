
namespace ExcelDataReader.Portable.Core.BinaryFormat
{
	/// <summary>
	/// Represents Worksheet section in workbook
	/// </summary>
	internal class XlsWorksheet
	{
		private readonly uint m_dataOffset;
		private readonly int m_Index;
		private readonly string m_Name = string.Empty;
	    private readonly string m_VisibleState = string.Empty;
		private XlsBiffSimpleValueRecord m_CalcCount;
		private XlsBiffSimpleValueRecord m_CalcMode;
		private XlsBiffRecord m_Delta;
		private XlsBiffDimensions m_Dimensions;
		private XlsBiffSimpleValueRecord m_Iteration;
		private XlsBiffSimpleValueRecord m_RefMode;
		private XlsBiffRecord m_Window;

		public XlsWorksheet(int index, XlsBiffBoundSheet refSheet)
		{
			m_Index = index;
			m_Name = refSheet.SheetName;
			m_dataOffset = refSheet.StartOffset;

		    switch (refSheet.VisibleState)
		    {
		        case XlsBiffBoundSheet.SheetVisibility.Hidden:
		            m_VisibleState = "hidden";
		            break;
                case XlsBiffBoundSheet.SheetVisibility.VeryHidden:
		            m_VisibleState = "veryhidden";
		            break;
                default:
		            m_VisibleState = "visible";
		            break;
		    }
		}

		/// <summary>
		/// Name of worksheet
		/// </summary>
		public string Name
		{
			get { return m_Name; }
		}

        /// <summary>
        /// Visibility of worksheet
        /// </summary>
        public string VisibleState
        {
            get { return m_VisibleState; }
        }

		/// <summary>
		/// Zero-based index of worksheet
		/// </summary>
		public int Index
		{
			get { return m_Index; }
		}

		/// <summary>
		/// Offset of worksheet data
		/// </summary>
		public uint DataOffset
		{
			get { return m_dataOffset; }
		}

		public XlsBiffSimpleValueRecord CalcMode
		{
			get { return m_CalcMode; }
			set { m_CalcMode = value; }
		}

		public XlsBiffSimpleValueRecord CalcCount
		{
			get { return m_CalcCount; }
			set { m_CalcCount = value; }
		}

		public XlsBiffSimpleValueRecord RefMode
		{
			get { return m_RefMode; }
			set { m_RefMode = value; }
		}

		public XlsBiffSimpleValueRecord Iteration
		{
			get { return m_Iteration; }
			set { m_Iteration = value; }
		}

		public XlsBiffRecord Delta
		{
			get { return m_Delta; }
			set { m_Delta = value; }
		}

		/// <summary>
		/// Dimensions of worksheet
		/// </summary>
		public XlsBiffDimensions Dimensions
		{
			get { return m_Dimensions; }
			set { m_Dimensions = value; }
		}

		public XlsBiffRecord Window
		{
			get { return m_Window; }
			set { m_Window = value; }
		}

	}
}