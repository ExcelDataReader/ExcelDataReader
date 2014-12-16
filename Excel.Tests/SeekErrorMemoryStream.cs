using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

#if LEGACY
namespace Excel.Tests
#else
namespace ExcelDataReader.Tests
#endif
{
	public class SeekErrorMemoryStream : MemoryStream
	{
		private bool canSeek = false;

		private SeekErrorMemoryStream()
		{
			
		}
		/// <summary>
		/// Creates SeekErrorMemoryStream copy data from the source
		/// </summary>
		/// <param name="source"></param>
		public static SeekErrorMemoryStream CreateFromStream(Stream source)
		{
			var forwardStream = new SeekErrorMemoryStream();
			forwardStream.canSeek = true;

			Helper.CopyStream(source, forwardStream);
			forwardStream.Seek(0, SeekOrigin.Begin);
			
			//now disable seek
			forwardStream.canSeek = false;

			return forwardStream;
		}

		public override long Seek(long offset, SeekOrigin loc)
		{
			if (canSeek)
				return base.Seek(offset, loc);

			//throw offset error to simuate problem we had with HttpInputStream
			throw new ArgumentOutOfRangeException("offset");
		}

		public override bool CanSeek
		{
			get
			{
				return canSeek;
			}
		}
	}

}
