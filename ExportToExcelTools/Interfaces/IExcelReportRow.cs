using System.Collections.Generic;

namespace ExportToExcelTools.Interfaces
{
    public interface IExcelReportRow
    {
        string Style { get; }
        List<IExcelReportCell> Cells { get; }
        IExcelReportRow AddCell(string value);
        IExcelReportRow AddCell(string value, string style);
        IExcelReportRow AddCell(string value, int? colSpan);
        IExcelReportRow AddCell(string value, string style, int? colSpan);
        IExcelReportRow SetStyle(string style);

    }
}
