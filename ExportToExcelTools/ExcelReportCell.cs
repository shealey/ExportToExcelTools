using ExportToExcelTools.Interfaces;

namespace ExportToExcelTools
{
    public class ExcelReportCell : IExcelReportCell
    {
        public int? ColSpan { get; private set; }
        public string Text { get; private set; }
        public string Style { get; private set; }

        public ExcelReportCell(string value)
        {            
            Text = !string.IsNullOrWhiteSpace(value) ? value : string.Empty;
            Style = Constants.DefaultCellStyle;
        }

        public ExcelReportCell(string value, int? colSpan)
        {
            ColSpan = colSpan;
            Text = !string.IsNullOrWhiteSpace(value) ? value : string.Empty;
            Style = Constants.DefaultCellStyle;
        }

        public ExcelReportCell(string value, string style)
        {            
            Text = !string.IsNullOrWhiteSpace(value) ? value : string.Empty;
            Style = !string.IsNullOrWhiteSpace(style) ? style : Constants.DefaultCellStyle;
        }

        public ExcelReportCell(string value, string style, int? colSpan)
        {
            ColSpan = colSpan;
            Text = !string.IsNullOrWhiteSpace(value) ? value : string.Empty;
            Style = !string.IsNullOrWhiteSpace(style) ? style : Constants.DefaultCellStyle;
        }
    }
}
