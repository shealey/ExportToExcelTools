using FluentAssertions;
using System.Data;
using System.Linq;
using Telerik.Web.UI;
using Xunit;

namespace ExportToExcelTools.UnitTests
{
    public class RadGridExportToExcelReportTests
    {
        #region Helper Methods
        private RadGrid CreateRadGrid()
        {
            var table = new DataTable();
            table.Columns.Add(new DataColumn("Header"));
            table.Columns.Add(new DataColumn("InvisibleHeader"));
            table.Rows.Add("SomeValue", "AnotherValue");            

            var grid = new RadGrid();
            grid.MasterTableView.Columns.Add(new GridBoundColumn { HeaderText = "Header", Visible = true, UniqueName = "Header" });            
            grid.MasterTableView.Columns.Add(new GridBoundColumn { HeaderText = "Invisible Header", Visible = false, UniqueName = "InvisibleHeader" });
            
            grid.DataSource = table;
            grid.DataBind();            
            
            return grid;
        }
        #endregion

        [Fact]
        public void RadGridExportToExcelReport_OnCreate_AddsRowsFromGrid()
        {
            var grid = CreateRadGrid();
            var report = new RadGridExportToExcelReport(grid);

            report.ReportRows.Should().HaveCount(1);
        }

        [Fact]
        public void RadGridExportToExcelReport_OnCreate_AddsVisibleHeadersFromGrid()
        {
            var grid = CreateRadGrid();
            var report = new RadGridExportToExcelReport(grid);

            report.HeaderRow.Cells.Should().HaveCount(1);
        }

        [Fact]
        public void RadGridExportToExcelReport_OnCreate_AddsHeaderTextFromGrid()
        {
            var grid = CreateRadGrid();
            var report = new RadGridExportToExcelReport(grid);

            report.HeaderRow.Cells.First().Text.Should().Be("Header");
        }

        [Fact]
        public void RadGridExportToExcelReport_OnCreate_AddsRowTextFromGrid()
        {
            var grid = CreateRadGrid();
            var report = new RadGridExportToExcelReport(grid);

            report.ReportRows.First().Cells.Any(c => c.Text == "SomeValue");
        }        
    }
}
