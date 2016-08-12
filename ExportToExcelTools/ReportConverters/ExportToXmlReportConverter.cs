using ExportToExcelTools.Interfaces;
using System.Web.UI;
using System.IO;
using System.Xml.Linq;

namespace ExportToExcelTools
{
    public class ExportToXmlReportConverter : IExcelReportConverter<XDocument>
    {               
        public XDocument Convert(IExportToExcelReport report)
        {
            var _stringWriter = new StringWriter();
            var _htmlWriter = new HtmlTextWriter(_stringWriter);            

            _htmlWriter.WriteLine("<table>");

            if (report.HeaderRow.Cells.Count > 0)
            {
                _htmlWriter.WriteLine("<tr class=\"" + report.HeaderRow.Style + "\">");

                foreach (var heading in report.HeaderRow.Cells)
                {
                    if (heading.ColSpan.HasValue)
                    {
                        _htmlWriter.WriteLine("<th colspan=\"" + heading.ColSpan + "\">" + Utils.EscapeXMLValue(heading.Text) + "</th>");
                    }
                    else
                    {
                        _htmlWriter.WriteLine("<th>" + Utils.EscapeXMLValue(heading.Text) + "</th>");
                    }
                }

                _htmlWriter.WriteLine("</tr>");

                foreach (var row in report.ReportRows)
                {
                    if (row.Style.Length > 0)
                    {
                        _htmlWriter.WriteLine("<tr class=\"" + row.Style + "\">");
                    }
                    else
                    {
                        _htmlWriter.WriteLine("<tr>");
                    }

                    foreach (var cell in row.Cells)
                    {
                        if (cell.ColSpan.HasValue)
                        {
                            _htmlWriter.WriteLine("<td class=\"" + cell.Style + "\" colspan=\"" + cell.ColSpan + "\">" + Utils.EscapeXMLValue(cell.Text) + "</td>");
                        }
                        else
                        {
                            _htmlWriter.WriteLine("<td class=\"" + cell.Style + "\">" + Utils.EscapeXMLValue(cell.Text) + "</td>");
                        }                        
                    }

                    _htmlWriter.WriteLine("</tr>");
                }
            }

            _htmlWriter.WriteLine("</table>");

            return XDocument.Parse(_stringWriter.ToString());
        }
    }
}
