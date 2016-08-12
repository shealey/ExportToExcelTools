using FluentAssertions;
using Xunit;

namespace ExportToExcelTools.UnitTests
{
    public class UtilsTests
    {
        [Theory]
        [InlineData("'", "&apos;")]
        [InlineData("\"", "&quot;")]
        [InlineData(">", "&gt;")]
        [InlineData("<", "&lt;")]
        [InlineData("&", "&amp;")]
        public void EscapeXMLValue_Default_ReplacesExpectedCharacters(string text, string expectedXml)
        {            
            Utils.EscapeXMLValue(text).Should().Be(expectedXml);
        }
    }
}
