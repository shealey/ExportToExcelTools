using ExportToExcelTools.Interfaces;
using System.Collections.Generic;
using System.Xml.Linq;

namespace ExportToExcelTools
{
    public class ExportToExcelReportBase : IExportToExcelReport
    {
        private IExcelReportConverter<XDocument> _xmlReportGenerator;
        private IExcelFileGenerator<XDocument> _xmlFileGenerator;
        public IExcelReportRow HeaderRow { get; private set; }
        public ICollection<IExcelStyleBuilder> Styles { get; private set; }
        public List<IExcelReportRow> ReportRows { get; private set; }

        public ExportToExcelReportBase()
        {
            _xmlReportGenerator = new ExportToXmlReportConverter();
            _xmlFileGenerator = new ExcelFromXmlFileGenerator();

            HeaderRow = new ExcelReportRow(Constants.DefaultHeaderStyle);
            ReportRows = new List<IExcelReportRow>();
            Styles = new List<IExcelStyleBuilder>
            {
                new DefaultHeaderStyleBuilder(),
                new DefaultCellStyleBuilder()
            };
        }

        public ExportToExcelReportBase(
            IExcelReportConverter<XDocument> xmlReportGenerator,
            IExcelFileGenerator<XDocument> xmlFileGenerator) : base()
        {
            _xmlFileGenerator = xmlFileGenerator;
            _xmlReportGenerator = xmlReportGenerator;

            HeaderRow = new ExcelReportRow(Constants.DefaultHeaderStyle);
            ReportRows = new List<IExcelReportRow>();
            Styles = new List<IExcelStyleBuilder>
            {
                new DefaultHeaderStyleBuilder(),
                new DefaultCellStyleBuilder()
            };
        }        

        public byte[] GenerateFile()
        {
            var xml = _xmlReportGenerator.Convert(this);
            return _xmlFileGenerator.CreateFile(xml, Styles);
        }        
    }
}
