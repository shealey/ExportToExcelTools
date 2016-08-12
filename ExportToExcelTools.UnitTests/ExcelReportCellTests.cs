using FluentAssertions;
using Xunit;

namespace ExportToExcelTools.UnitTests
{    
    public class ExcelReportCellTests
    {
        [Fact]
        public void CellCreated_WithoutClassName_DefaultClassNameUsed()
        {
            var cell = new ExcelReportCell(string.Empty);
            cell.Style.Should().Be(Constants.DefaultCellStyle);            
        }

        [Fact]
        public void CellCreated_WithClassName_ClassNameMatches()
        {
            var cell = new ExcelReportCell(string.Empty, "Style");
            cell.Style.Should().Be("Style");
        }

        [Fact]
        public void CellCreated_ValueNotNullOrWhitespace_SetsTextAsValue()
        {
            var cell = new ExcelReportCell("Value");
            cell.Text.Should().Be("Value");
        }

        [Theory]
        [InlineData(null)]
        [InlineData("   ")]
        public void CellCreated_ValueNullOrWhitespace_AddsEmptyText(string text)
        {
            var cell = new ExcelReportCell(null);
            cell.Text.Should().Be(string.Empty);
        }

        [Fact]
        public void CellCreated_WithoutColSpan_DefaultColspanNull()
        {
            var cell = new ExcelReportCell(string.Empty);
            cell.ColSpan.Should().Be(null);
        }

        [Fact]
        public void CellCreated_WithColSpan_ColSpanAssigned()
        {
            var cell = new ExcelReportCell(string.Empty, 2);
            cell.ColSpan.Should().Be(2);
        }

    }
}
