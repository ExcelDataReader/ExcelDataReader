#if NET20
using System;
using System.IO;
using System.Collections.Generic;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;

namespace ExcelDataReader.Core {
    public class ZipEntry {
		ICSharpCode.SharpZipLib.Zip.ZipFile Handle;
		ICSharpCode.SharpZipLib.Zip.ZipEntry Entry;

		internal ZipEntry(ZipFile handle, ICSharpCode.SharpZipLib.Zip.ZipEntry entry) {
			Handle = handle;
			Entry = entry;
		}

		public Stream Open() {
            return Handle.GetInputStream(Entry);
        }
    }

    public class ZipArchive : IDisposable {
		ZipFile Handle;

		public ZipArchive(Stream stream) {
			Handle = new ZipFile(stream);
		}

        public ZipEntry GetEntry(string name) {
			var entry = Handle.GetEntry(name);
			if (entry == null) 
				return null;
			return new ZipEntry(Handle, entry);
        }

		#region IDisposable Support
		private bool disposedValue = false; // To detect redundant calls

		protected virtual void Dispose(bool disposing) {
			if (!disposedValue) {
				if (disposing) {
					// TODO: dispose managed state (managed objects).
				}

				// TODO: free unmanaged resources (unmanaged objects) and override a finalizer below.
				// TODO: set large fields to null.

				disposedValue = true;
			}
		}

		// TODO: override a finalizer only if Dispose(bool disposing) above has code to free unmanaged resources.
		// ~ZipArchive() {
		//   // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
		//   Dispose(false);
		// }

		// This code added to correctly implement the disposable pattern.
		public void Dispose() {
			// Do not change this code. Put cleanup code in Dispose(bool disposing) above.
			Dispose(true);
			// TODO: uncomment the following line if the finalizer is overridden above.
			// GC.SuppressFinalize(this);
		}
		#endregion
	}
}

#endif
