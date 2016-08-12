using ExportToExcelTools.Interfaces;
using FluentAssertions;
using NSubstitute;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace ExportToExcelTools.UnitTests
{
    public class ExportToExcelReportBaseTests
    {
        [Fact]
        public void ExportToExcelReport_Default_ReportRowsIsEmptyList()
        {
            var report = new ExportToExcelReportBase();
            report.ReportRows.Should().NotBeNull().And.HaveCount(0);
        }

        [Fact]
        public void ExportToExcelReport_Default_HasHeaderRowInitialised()
        {
            var report = new ExportToExcelReportBase();
            report.HeaderRow.Should().NotBeNull();
        }

        [Fact]
        public void ExportToExcelReport_Default_HeaderRowHasDefaultStyle()
        {
            var report = new ExportToExcelReportBase();
            report.HeaderRow.Style.Should().Be(Constants.DefaultHeaderStyle);
        }

        [Fact]
        public void ExportToExcelReport_Default_StylesContainsDefaultHeaderStyle()
        {
            var report = new ExportToExcelReportBase();

            var style = report.Styles.FirstOrDefault(s => s.Name == Constants.DefaultHeaderStyle);

            style.Should().NotBeNull();
        }

        [Fact]
        public void ExportToExcelReport_Default_StylesContainsDefaultCellStyle()
        {
            var report = new ExportToExcelReportBase();

            var style = report.Styles.FirstOrDefault(s => s.Name == Constants.DefaultCellStyle);

            style.Should().NotBeNull();
        }

        [Fact]
        public void GenerateFile_WhenCalled_CallsConvertFromReportConverter()
        {
            var fakeReportConverter = Substitute.For<IExcelReportConverter<XDocument>>();
            var report = new ExportToExcelReportBase(fakeReportConverter, Substitute.For<IExcelFileGenerator<XDocument>>());
            var reportXml = XDocument.Parse("<table></table>");

            fakeReportConverter.Convert(Arg.Any<IExportToExcelReport>()).Returns(reportXml);
            report.GenerateFile();

            fakeReportConverter.Received().Convert(Arg.Any<IExportToExcelReport>());
        }

        [Fact]
        public void GenerateFile_WhenCalled_CallsCreateFileFromFileGenerator()
        {
            var fakeReportConverter = Substitute.For<IExcelReportConverter<XDocument>>();
            var fakeFileGenerator = Substitute.For<IExcelFileGenerator<XDocument>>();
            var report = new ExportToExcelReportBase(fakeReportConverter, fakeFileGenerator);
            var reportXml = XDocument.Parse("<table></table>");

            fakeReportConverter.Convert(Arg.Any<IExportToExcelReport>()).Returns(reportXml);
            report.GenerateFile();

            fakeFileGenerator.Received().CreateFile(reportXml, report.Styles);            
        }
    }
}
