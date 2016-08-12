using ExportToExcelTools.Interfaces;
using NPOI.SS.UserModel;

namespace ExportToExcelTools
{
    public class DefaultCellStyleBuilder : IExcelStyleBuilder
    {
        public string Name { get { return Constants.DefaultCellStyle; } }

        public ICellStyle CreateStyle(IWorkbook workbook)
        {
            var font = workbook.CreateFont();
            font.FontHeightInPoints = 11;
            font.FontName = "Calibri";

            var style = workbook.CreateCellStyle();
            style.SetFont(font);

            return style;
        }
    }
}
