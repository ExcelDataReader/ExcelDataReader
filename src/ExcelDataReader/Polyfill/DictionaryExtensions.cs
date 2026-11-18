#if !(NETSTANDARD2_1_OR_GREATER || NET8_0_OR_GREATER)
namespace System.Collections.Generic;

internal static class DictionaryExtensions
{
    public static TValue? GetValueOrDefault<TKey, TValue>(this IReadOnlyDictionary<TKey, TValue> dictionary, TKey key)
        where TKey : notnull =>
        dictionary.TryGetValue(key, out var value) ? value : default;

    public static bool TryAdd<TKey, TValue>(this Dictionary<TKey, TValue> dictionary, TKey key, TValue value)
        where TKey : notnull
    {
        if (dictionary.ContainsKey(key))
            return false;

        dictionary.Add(key, value);
        return true;
    }
}
#endif
