using ExportToExcelTools.Interfaces;
using NPOI.SS.UserModel;

namespace ExportToExcelTools
{
    public class DefaultHeaderStyleBuilder : IExcelStyleBuilder
    {
        public string Name { get { return Constants.DefaultHeaderStyle; } }

        public ICellStyle CreateStyle(IWorkbook workbook)
        {
            var font = workbook.CreateFont();
            font.FontHeightInPoints = 11;
            font.FontName = "Calibri";
            font.Boldweight = (short)FontBoldWeight.Bold;

            var style = workbook.CreateCellStyle();
            style.SetFont(font);

            return style;
        }
    }
}
