using System.Collections.Generic;

namespace ExportToExcelTools.Interfaces
{
    public interface IExcelFileGenerator<T> where T : class
    {
        byte[] CreateFile(T generatedReport, IEnumerable<IExcelStyleBuilder> styles);
    }
}
