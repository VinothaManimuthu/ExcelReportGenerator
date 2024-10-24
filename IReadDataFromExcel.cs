using System;
using System.Collections.Generic;
namespace  ExcelReportGenerator
{
    public interface IReadDataFromExcel
    {
        List<DataRecord> ReadData(string inputFilePath, string stateFilter);
    }
}
