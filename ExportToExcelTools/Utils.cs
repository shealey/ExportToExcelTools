using System;
using System.Security;

namespace ExportToExcelTools
{
    public static class Utils
    {
        public static string EscapeXMLValue(string xmlString)
        {
            return SecurityElement.Escape(xmlString);            
        }
    }
}
