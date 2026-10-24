using System.Runtime.ExceptionServices;

namespace ExcelDataReader.Core;

internal static class ResourceCleanup
{
    public static void RestorePosition(Stream stream, long position, Exception? failure)
    {
        try
        {
            stream.Position = position;
        }
        catch (Exception) when (failure != null)
        {
            // Preserve the operation failure already being rethrown by the caller.
        }
    }

    public static void DisposeAll(Exception? failure, params IDisposable?[] resources)
    {
        Exception? firstError = null;
        foreach (var resource in resources)
        {
            try
            {
                resource?.Dispose();
            }
            catch (Exception error)
            {
                firstError ??= error;
            }
        }

        // Callers rethrow the operation failure after cleanup.
        if (failure == null && firstError != null)
            ExceptionDispatchInfo.Capture(firstError).Throw();
    }
}