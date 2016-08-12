using System.Collections.Generic;
using System.Xml.Linq;

namespace ExportToExcelTools.Interfaces
{
    public interface IExportToExcelReport
    {      
        IExcelReportRow HeaderRow { get; }
        List<IExcelReportRow> ReportRows { get; }     
        ICollection<IExcelStyleBuilder> Styles { get; }         
        byte[] GenerateFile();
    }
}
