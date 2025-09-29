namespace ExcelDataReader.Tests;

[SetUpFixture]
public sealed class SetUpFixture
{
    [OneTimeSetUp]
    public static void SetUp()
    {
        Log.Log.InitializeWith<NunitLogFactory>();

#if NETCOREAPP1_0_OR_GREATER
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
#endif
    }
}
