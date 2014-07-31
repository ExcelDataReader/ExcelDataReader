using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Excel.Tests
{
    [TestClass]
    public class ReferenceHelperTest
    {
        [TestMethod]
        public void ReferenceToColumnAndRow()
        {
            var A1 = ReferenceHelper.ReferenceToColumnAndRow("A1");
            Assert.AreEqual(1, A1[0]);
            Assert.AreEqual(1, A1[1]);

            var a1 = ReferenceHelper.ReferenceToColumnAndRow("a1");
            Assert.AreEqual(1, a1[0]);
            Assert.AreEqual(1, a1[1]);

            var b1 = ReferenceHelper.ReferenceToColumnAndRow("B1");
            Assert.AreEqual(1, b1[0]);
            Assert.AreEqual(2, b1[1]);

            var z2 = ReferenceHelper.ReferenceToColumnAndRow("Z2");
            Assert.AreEqual(2, z2[0]);
            Assert.AreEqual(26, z2[1]);

            var aa99 = ReferenceHelper.ReferenceToColumnAndRow("AA99");
            Assert.AreEqual(99, aa99[0]);
            Assert.AreEqual(27, aa99[1]); // (26)^1 * 1 +  26^0 * 1

            var ba99 = ReferenceHelper.ReferenceToColumnAndRow("BA99");
            Assert.AreEqual(99, ba99[0]);
            Assert.AreEqual((26 * 2) + 1, ba99[1]);

            var aaa99 = ReferenceHelper.ReferenceToColumnAndRow("AAA99");
            Assert.AreEqual(99, aaa99[0]);
            Assert.AreEqual(26*26 + 26 + 1, aaa99[1]);

            var zzz99 = ReferenceHelper.ReferenceToColumnAndRow("ZZZ99");
            Assert.AreEqual(99, zzz99[0]);
            Assert.AreEqual((26*26)*26 + 26*26 + 26, zzz99[1]);

        }

    }
}
