#if NET20
namespace ExcelDataReader
{
    /// <summary>
    /// Encapsulates a method that has one parameter and returns a value of the type specified by the TResult parameter.
    /// </summary>
    /// <typeparam name="T1">The type of the parameter of the method that this delegate encapsulates.</typeparam>
    /// <typeparam name="TResult">The type of the return value of the method that this delegate encapsulates.</typeparam>
    /// <param name="arg1">The parameter of the method that this delegate encapsulates.</param>
    /// <returns>The return value of the method that this delegate encapsulates.</returns>
    public delegate TResult Func<T1, TResult>(T1 arg1);
}
#endif