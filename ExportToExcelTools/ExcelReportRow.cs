using System.Collections.Generic;
using ExportToExcelTools.Interfaces;

namespace ExportToExcelTools
{
    public class ExcelReportRow : IExcelReportRow
    {
        public string Style { get; private set; }
        public List<IExcelReportCell> Cells { get; private set; }

        public ExcelReportRow(string style = null)
        {
            Style = !string.IsNullOrWhiteSpace(style) ? style : string.Empty;
            Cells = new List<IExcelReportCell>();
        }

        public IExcelReportRow AddCell(string value)
        {
            this.Cells.Add(new ExcelReportCell(value));
            return this;
        }

        public IExcelReportRow AddCell(string value, int? colSpan)
        {
            this.Cells.Add(new ExcelReportCell(value, colSpan));
            return this;
        }

        public IExcelReportRow AddCell(string value, string style)
        {
            this.Cells.Add(new ExcelReportCell(value, style));
            return this;
        }

        public IExcelReportRow AddCell(string value, string style, int? colSpan)
        {
            this.Cells.Add(new ExcelReportCell(value, style, colSpan));
            return this;
        }

        public IExcelReportRow SetStyle(string style)
        {
            Style = !string.IsNullOrWhiteSpace(style) ? style : string.Empty;
            return this;
        }
    }
}
