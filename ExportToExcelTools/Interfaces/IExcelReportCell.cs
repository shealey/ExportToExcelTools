namespace ExportToExcelTools.Interfaces
{
    public interface IExcelReportCell
    {
        string Text { get; }
        string Style { get; }
        int? ColSpan { get; }
    }
}
