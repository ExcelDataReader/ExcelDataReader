#if !NET8_0_OR_GREATER
namespace System;

internal static class ObjectDisposedExceptionExtensions
{
    extension(ObjectDisposedException)
    {
        public static void ThrowIf(bool condition, object instance)
        {
            if (condition)
                throw new ObjectDisposedException(instance.GetType().FullName);
        }
    }
}
#endif
