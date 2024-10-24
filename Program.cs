using System;
using OfficeOpenXml;
using log4net;
using log4net.Config;

namespace ExcelReportGenerator
{
    class Program
    {
        private static readonly ILog log = LogManager.GetLogger(typeof(Program));
        // Load log4net configuration from external file

        static void Main(string[] args)
        {
            var excelService = new ExcelService();
            excelService.Run();
        }

    }
}