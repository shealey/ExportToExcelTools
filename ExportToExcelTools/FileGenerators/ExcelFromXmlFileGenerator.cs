using ExportToExcelTools.Interfaces;
using NPOI.HSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

namespace ExportToExcelTools
{
    public class ExcelFromXmlFileGenerator : IExcelFileGenerator<XDocument>
    {       
        public byte[] CreateFile(XDocument generatedReport, IEnumerable<IExcelStyleBuilder> styles)
        {
            var workbook = HSSFWorkbook.Create(InternalWorkbook.CreateWorkbook());
            var sheet = (HSSFSheet)workbook.CreateSheet("Sheet1");
            var totalNumberOfColumns = 0;

            var styleFactory = new ExcelStyleFactory(workbook, styles);

            var rowIndex = 0;

            foreach (var rowElement in generatedReport.Descendants("tr"))
            {
                var row = sheet.CreateRow(rowIndex);
                var rowStyle = styleFactory.CreateStyle(rowElement.Attribute("class") != null ? rowElement.Attribute("class").Value : null);

                var cellIndex = 0;

                foreach (var cellElement in rowElement.Elements())
                {
                    var cell = row.CreateCell(cellIndex);
                    cell.SetCellValue(cellElement.Value);

                    if (cellElement.Attribute("class") != null &&
                        cellElement.Attribute("class").Value != null)
                    {
                        cell.CellStyle = styleFactory.CreateStyle(cellElement.Attribute("class").Value);
                    }
                    else
                    {
                        cell.CellStyle = rowStyle;
                    }                    

                    int colSpan;
                    if (cellElement.Attribute("colspan") != null &&
                        cellElement.Attribute("colspan").Value != null &&
                        int.TryParse(cellElement.Attribute("colspan").Value, out colSpan))
                    {
                        sheet.AddMergedRegion(new CellRangeAddress(rowIndex, rowIndex, cellIndex, cellIndex + colSpan - 1));
                        cellIndex += colSpan - 1;
                    }

                    cellIndex++;
                }

                totalNumberOfColumns = cellIndex > totalNumberOfColumns ? cellIndex : totalNumberOfColumns;

                rowIndex++;
            }

            for (var columnIndex = 0; columnIndex < totalNumberOfColumns; columnIndex++)
            {
                sheet.AutoSizeColumn(columnIndex);
            }

            using (var memoryStream = new MemoryStream())
            {
                workbook.Write(memoryStream);
                memoryStream.Position = 0;
                return memoryStream.ToArray();
            }
        }
    }
}
