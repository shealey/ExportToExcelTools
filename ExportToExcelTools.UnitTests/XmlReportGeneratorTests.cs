using ExportToExcelTools.Interfaces;
using FluentAssertions;
using NSubstitute;
using System.Collections.Generic;
using System.Xml.Linq;
using Xunit;

namespace ExportToExcelTools.UnitTests
{
    public class XmlReportconverterTests
    {        
        [Fact]
        public void Generate_WhenCalled_ProducesExpectedOutput()
        {
            var converter = new ExportToXmlReportConverter();
            var fakeReport = Substitute.For<IExportToExcelReport>();
            var headerRow = Substitute.For<IExcelReportRow>();
            var headerCell = Substitute.For<IExcelReportCell>();
            var dataRow = Substitute.For<IExcelReportRow>();
            var dataCell = Substitute.For<IExcelReportCell>();

            fakeReport.HeaderRow.Returns(headerRow);           
            headerRow.Cells.Returns(new List<IExcelReportCell> { headerCell });
            headerRow.Style.Returns(Constants.DefaultHeaderStyle);

            headerCell.Style.Returns(string.Empty);
            headerCell.Text.Returns("Header");

            fakeReport.ReportRows.Returns(new List<IExcelReportRow> { dataRow });
            dataRow.Cells.Returns(new List<IExcelReportCell> { dataCell });
            dataRow.Style.Returns(string.Empty);

            dataCell.Style.Returns(Constants.DefaultCellStyle);
            dataCell.Text.Returns("Data");

            var htmlString = "<table>\r\n" +
            "<tr class=\"" + Constants.DefaultHeaderStyle + "\">\r\n" +
            "<th>Header</th>\r\n" +
            "</tr>\r\n" +
            "<tr>\r\n" +
            "<td class=\"" + Constants.DefaultCellStyle + "\">Data</td>\r\n" +
            "</tr>\r\n" +
            "</table>\r\n";

            var htmlAsXml = XDocument.Parse(htmlString).ToString();

            converter.Convert(fakeReport).ToString().Should().Be(htmlAsXml);
        }

        [Fact]
        public void Generate_CellsHaveColspansSet_ContainsColspanAttributes()
        {
            var converter = new ExportToXmlReportConverter();
            var fakeReport = Substitute.For<IExportToExcelReport>();
            var headerRow = Substitute.For<IExcelReportRow>();
            var headerCell = Substitute.For<IExcelReportCell>();
            var dataRow = Substitute.For<IExcelReportRow>();
            var dataCell = Substitute.For<IExcelReportCell>();

            fakeReport.HeaderRow.Returns(headerRow);
            headerRow.Cells.Returns(new List<IExcelReportCell> { headerCell });
            headerRow.Style.Returns(Constants.DefaultHeaderStyle);

            headerCell.Style.Returns(string.Empty);
            headerCell.Text.Returns("Header");            

            fakeReport.ReportRows.Returns(new List<IExcelReportRow> { dataRow });
            dataRow.Cells.Returns(new List<IExcelReportCell> { dataCell });
            dataRow.Style.Returns(string.Empty);

            dataCell.Style.Returns(Constants.DefaultCellStyle);
            dataCell.Text.Returns("Data");
            dataCell.ColSpan.Returns(1);

            var xml = converter.Convert(fakeReport).ToString();

            xml.Should().Contain("colspan");
        }

        [Fact]
        public void Generate_NoCellsHaveColspansSet_DoesNotContainColspanAttributes()
        {
            var converter = new ExportToXmlReportConverter();
            var fakeReport = Substitute.For<IExportToExcelReport>();
            var headerRow = Substitute.For<IExcelReportRow>();
            var headerCell = Substitute.For<IExcelReportCell>();
            var dataRow = Substitute.For<IExcelReportRow>();
            var dataCell = Substitute.For<IExcelReportCell>();

            fakeReport.HeaderRow.Returns(headerRow);
            headerRow.Cells.Returns(new List<IExcelReportCell> { headerCell });
            headerRow.Style.Returns(Constants.DefaultHeaderStyle);

            headerCell.Style.Returns(string.Empty);
            headerCell.Text.Returns("Header");

            fakeReport.ReportRows.Returns(new List<IExcelReportRow> { dataRow });
            dataRow.Cells.Returns(new List<IExcelReportCell> { dataCell });
            dataRow.Style.Returns(string.Empty);

            dataCell.Style.Returns(Constants.DefaultCellStyle);
            dataCell.Text.Returns("Data");

            var xml = converter.Convert(fakeReport).ToString();
            xml.Should().NotContain("colspan");
        }
    }
}
