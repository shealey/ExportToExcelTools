using NPOI.SS.UserModel;

namespace ExportToExcelTools.Interfaces
{
    public interface IExcelStyleBuilder
    {
        string Name { get; }
        ICellStyle CreateStyle(IWorkbook workbook);
    }
}
