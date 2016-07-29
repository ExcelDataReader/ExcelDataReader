using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelDataReader.Tests
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

			CopyStream(source, forwardStream);
			forwardStream.Seek(0, SeekOrigin.Begin);
			
			//now disable seek
			forwardStream.canSeek = false;

			return forwardStream;
		}

		// Merged From linked CopyStream below and Jon Skeet's ReadFully example
		public static void CopyStream(Stream input, Stream output) {
			byte[] buffer = new byte[16 * 1024];
			int read;
			while ((read = input.Read(buffer, 0, buffer.Length)) > 0) {
				output.Write(buffer, 0, read);
			}
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
