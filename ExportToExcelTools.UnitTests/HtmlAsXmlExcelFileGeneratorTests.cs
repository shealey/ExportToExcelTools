using FluentAssertions;
using Xunit;
using System.Xml.Linq;
using ExportToExcelTools.Interfaces;
using System.Collections.Generic;

namespace ExportToExcelTools.UnitTests
{
    public class HtmlAsXmlExcelFileGeneratorTests
    {
        private const int BlankWorkbookFilesize = 4096;
        [Fact]
        public void Convert_NoRows_ProducesEmptyFile()
        {            
            var fileGenerator = new ExcelFromXmlFileGenerator();
            var xml = XDocument.Parse("<table></table>");

            var output = fileGenerator.CreateFile(xml, new List<IExcelStyleBuilder>());

            //blank workbook size
            output.Length.Should().Be(BlankWorkbookFilesize);
        }

        [Fact]
        public void Convert_HasRows_ProducesEmptyFile()
        {
            var fileGenerator = new ExcelFromXmlFileGenerator();
            var xml = XDocument.Parse("<table><tr><th>Heading</th></tr></table>");

            var output = fileGenerator.CreateFile(xml, new List<IExcelStyleBuilder>());
            
            output.Length.Should().BeGreaterThan(BlankWorkbookFilesize);
        }
    }
}
