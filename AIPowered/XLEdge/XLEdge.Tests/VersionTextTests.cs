using Microsoft.VisualStudio.TestTools.UnitTesting;
using XLEdge.Shared;

namespace XLEdge.Tests
{
    /// <summary>Versions are shown as 11.1.3, not 11.1.3.0 (the assembly/manifest form stays the data).</summary>
    [TestClass]
    public class VersionTextTests
    {
        [TestMethod]
        [DataRow("11.1.3.0", "11.1.3")]
        [DataRow(" 11.1.3.0 ", "11.1.3")]
        [DataRow("11.1.3", "11.1.3")]
        [DataRow("11.1.3.5", "11.1.3.5")]
        [DataRow("11.1.0.0", "11.1.0")]
        [DataRow("11.1", "11.1")]
        [DataRow("not a version", "not a version")]
        [DataRow("", "")]
        [DataRow(null, null)]
        public void Display_DropsAZeroRevisionOnly(string version, string expected)
        {
            Assert.AreEqual(expected, VersionText.Display(version));
        }
    }
}
