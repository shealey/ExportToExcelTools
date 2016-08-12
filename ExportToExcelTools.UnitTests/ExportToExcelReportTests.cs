using FluentAssertions;
using Xunit;

namespace ExportToExcelTools.UnitTests
{
    public class ExportToExcelReportTests
    {      
        [Fact]
        public void AddRow_WhenCalled_AddsRowToReportRows()
        {
            var report = new ExportToExcelReport();
            
            report.AddRow();

            report.ReportRows.Should().HaveCount(1);
        }        
    }
}
