using System.Runtime.ExceptionServices;

namespace ExcelDataReader.Core;

internal static class ResourceCleanup
{
    public static void DisposeAll(Exception? failure, params IDisposable?[] resources)
    {
        List<Exception>? errors = null;
        foreach (var resource in resources)
        {
            try
            {
                resource?.Dispose();
            }
            catch (Exception error)
            {
                (errors ??= []).Add(error);
            }
        }

        if (errors == null)
            return;
        if (failure != null)
            errors.Insert(0, failure);
        if (errors.Count == 1)
            ExceptionDispatchInfo.Capture(errors[0]).Throw();
        throw new AggregateException("Resource cleanup failed.", errors);
    }
}
