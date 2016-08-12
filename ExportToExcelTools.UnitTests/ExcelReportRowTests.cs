using FluentAssertions;
using Xunit;

namespace ExportToExcelTools.UnitTests
{
    public class ExcelReportRowTests
    {
        [Fact]
        public void RowCreated_WithoutStyle_StyleIsEmpty()
        {
            var row = new ExcelReportRow();
            row.Style.Should().Be(string.Empty);
        }

        [Fact]
        public void RowCreated_WithStyle_StyleAdded()
        {
            var row = new ExcelReportRow("style");
            row.Style.Should().Be("style");
        }

        [Fact]
        public void RowCreated_Default_CellsInitializedAsEmptyList()
        {
            var row = new ExcelReportRow();
            row.Cells.Should().NotBeNull().And.HaveCount(0);
        }

        [Fact]
        public void AddCell_WhenCalled_AddSingleCellToRows()
        {
            var row = new ExcelReportRow();

            row.AddCell(string.Empty);

            row.Cells.Should().NotBeNull().And.HaveCount(1);
        }

        [Fact]
        public void SetStyle_Default_ChangesStyle()
        {
            var row = new ExcelReportRow();

            row.SetStyle("style");

            row.Style.Should().Be("style");
        }

        [Fact]
        public void SetStyle_StyleIsNull_StyleIsEmptyString()
        {
            var row = new ExcelReportRow("style");

            row.SetStyle(null);

            row.Style.Should().Be(string.Empty);
        }
    }
}
