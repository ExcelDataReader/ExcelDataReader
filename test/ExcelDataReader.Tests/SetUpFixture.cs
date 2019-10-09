using System;
using System.Collections.Generic;
using System.Text;

using NUnit.Framework;

namespace ExcelDataReader.Tests
{
    [SetUpFixture]
    public sealed class SetUpFixture
    {
        [OneTimeSetUp]
        public static void SetUp()
        {
            Log.Log.InitializeWith<NunitLogFactory>();

#if NETCOREAPP1_0 || NETCOREAPP2_0
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
#endif
        }
    }
}
