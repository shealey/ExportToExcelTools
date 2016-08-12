using System.Collections.Generic;
using ExportToExcelTools.Interfaces;
using Telerik.Web.UI;
using System.Xml.Linq;

namespace ExportToExcelTools
{
    public class RadGridExportToExcelReport : ExportToExcelReportBase, IExportToExcelReport
    {        
        public RadGridExportToExcelReport(RadGrid grid) : base()
        {            
            ExtractRadGridHeadersAndRows(grid);
        }
        public RadGridExportToExcelReport(
            RadGrid grid, 
            IExcelReportConverter<XDocument> xmlReportGenerator,
            IExcelFileGenerator<XDocument> xmlFileGenerator) : base(xmlReportGenerator, xmlFileGenerator)
        {           
            ExtractRadGridHeadersAndRows(grid);
        }

        private void ExtractRadGridHeadersAndRows(RadGrid grid)
        {
            var columnKeys = new List<string>();
            foreach (GridColumn header in grid.Columns)
            {
                if (header.Visible)
                {
                    columnKeys.Add(header.UniqueName);                    
                    HeaderRow.AddCell(header.HeaderText);
                }
            }

            foreach (GridDataItem row in grid.Items)
            {
                var reportRow = new ExcelReportRow();

                foreach (var column in columnKeys)
                {
                    var cellText = !string.IsNullOrWhiteSpace(row[column].Text) ? System.Net.WebUtility.HtmlDecode(row[column].Text) : string.Empty;
                    reportRow.AddCell(cellText);
                }

                ReportRows.Add(reportRow);          
            }
        }             
    }
}
