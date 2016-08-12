namespace ExportToExcelTools.Interfaces
{
    public interface IExcelReportConverter<T> where T : class
    {
        T Convert(IExportToExcelReport report);
    }
}
