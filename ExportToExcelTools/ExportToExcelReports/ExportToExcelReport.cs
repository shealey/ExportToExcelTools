using ExportToExcelTools.Interfaces;
using System.Xml.Linq;

namespace ExportToExcelTools
{
    public class ExportToExcelReport : ExportToExcelReportBase, IExportToExcelReport
    {       
        public ExportToExcelReport() : base()
        {
               
        }        

        public ExportToExcelReport(
            IExcelReportConverter<XDocument> xmlReportGenerator,
            IExcelFileGenerator<XDocument> xmlFileGenerator) : base(xmlReportGenerator, xmlFileGenerator)
        {
            
        }       

        public IExcelReportRow AddRow()
        {
            var row = new ExcelReportRow();
            this.ReportRows.Add(row);
            return row;
        }             
    }
}
