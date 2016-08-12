using ExportToExcelTools.Interfaces;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Linq;

namespace ExportToExcelTools
{
    internal class ExcelStyleFactory
    {
        private readonly Dictionary<string, ICellStyle> _StyleCache = new Dictionary<string, ICellStyle>();
        private readonly IWorkbook _Workbook;
        private readonly IEnumerable<IExcelStyleBuilder> _Styles;

        public ExcelStyleFactory(IWorkbook workbook, IEnumerable<IExcelStyleBuilder> styles)
        {
            _Workbook = workbook;
            _Styles = styles;
        }

        public ICellStyle CreateStyle(string name)
        {
            if (name == null) return null;
            if (_StyleCache.ContainsKey(name)) return _StyleCache[name];

            var styleBuilder = _Styles.SingleOrDefault(s => s.Name == name);

            if (styleBuilder == null) return null;

            var style = styleBuilder.CreateStyle(_Workbook);

            _StyleCache.Add(styleBuilder.Name, style);

            return style;
        }
    }
}
